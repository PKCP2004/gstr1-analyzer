import io, re, zipfile
from pathlib import Path
from datetime import datetime
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# -----------------------------------------------------------------------------
# Excel template mapping
# -----------------------------------------------------------------------------
# Each output group is mapped to the corresponding filed GSTR-1 section.
# IMPORTANT: The PDF's "Value" column is treated as TAXABLE VALUE.
# Invoice Value is always calculated as:
#     Taxable Value + IGST + CGST + SGST + Cess
#
# The output template does not expose every tax column for every section, so
# the source tax fields are retained internally and only the template headers
# are written to Excel.

GROUP_CONFIG = {
    '4A, 4B, 6C - B2B, DE Invoices': {
        'type': 'aggregate',
        'sections': [
            ('4A - Taxable outward supplies made to registered persons', 'full', 'Total'),
            ('4B - Taxable outward supplies made to registered persons attracting tax on reverse charge', 'full', 'Total'),
            ('6C - Deemed Exports', 'full', 'Total'),
        ],
    },
    '5A, 5B - B2C (Large) Invoices': {
        'heading': '5 - Taxable outward inter-state supplies made to unregistered persons',
        'layout': 'value_igst_cess',
    },
    '6A - Exports Invoices': {
        'heading': '6A – Exports (with/without payment)',
        'heading_alt': '6A - Exports (with/without payment)',
        'layout': 'value_igst_cess',
    },
    '6B- SEZ WOP': {
        'heading': '6B - Supplies made to SEZ unit or SEZ developer',
        'row_label': '- SEZWOP',
        'layout': 'value_igst_cess',
    },
    '6B- SEZ WP': {
        'heading': '6B - Supplies made to SEZ unit or SEZ developer',
        'row_label': '- SEZWP',
        'layout': 'value_igst',
    },
    '7 - B2C (Others)': {
        'heading': '7- Taxable supplies (Net of debit and credit notes) to unregistered persons',
        'heading_alt': '7- Taxable supplies',
        'layout': 'full',
        # Table 7 has an extra "Document Type" text column (e.g. "Net Value")
        # sitting between the record count and the ₹ figures, and a rate-wise
        # breakdown (5%, 12%, 18%...) above the real total. Anchor precisely
        # on the row that starts with "Total" followed by a record count,
        # rather than any line merely containing the word "Total" somewhere,
        # so a rate-wise row or stray text can't be picked up by mistake.
        'row_pattern': r'^Total\s+[0-9]',
    },
    '8 - Nil rated, exempted and non GST outward supplies': {'type': 'table8'},
    '9A - Amended B2B Invoices': {
        'type': 'aggregate',
        'sections': [
            ('9A - Amendment to taxable outward supplies made to registered person in returns of earlier tax periods in table 4 - B2B Regular', 'full', 'Amended amount - Total'),
            ('9A - Amendment to taxable outward supplies made to registered person in returns of earlier tax periods in table 4 - B2B Reverse charge', 'full', 'Amended amount - Total'),
            ('9A - Amendment to Deemed Exports in returns of earlier tax periods in table 6C', 'full', 'Amended amount - Total'),
        ],
    },
    '9A - Amended B2C (Large) Invoices': {
        'heading': '9A - Amendment to Inter-State supplies made to unregistered person',
        'layout': 'value_igst_cess',
        'row_pattern': 'Amended amount - Total',
    },
    '9A - Amended Exports Invoices': {
        'heading': '9A - Amendment to Export supplies in returns of earlier tax periods in table 6A',
        'layout': 'value_igst_cess',
        'row_pattern': 'Amended amount - Total',
    },
    '9B - Credit/Debit Notes (Registered)': {
        'heading': '9B - Credit/Debit Notes (Registered)',
        'layout': 'full',
        'row_pattern': 'Total - Net off debit/credit notes',
    },
    '9B - Credit / Debit Notes (Unregistered)': {
        'heading': '9B - Credit/Debit Notes (Unregistered)',
        'layout': 'value_igst_cess',
        'row_pattern': 'Total - Net off debit/credit notes',
    },
    '9C - Amended Credit/Debit Notes (Registered)': {
        'heading': '9C - Amended Credit/Debit Notes (Registered)',
        'layout': 'full',
        'row_pattern': 'Amended amount - Total',
    },
    '9C - Amended Credit/Debit Notes (Unregistered': {
        'heading': '9C - Amended Credit/Debit Notes (Unregistered)',
        'layout': 'value_igst_cess',
        'row_pattern': 'Amended amount - Total',
    },
    '10 - Amended B2C(Others)': {
        'heading': '10 - Amendment to taxable outward supplies made to unregistered person',
        'layout': 'full',
        'row_pattern': 'Amended amount - Total',
    },
    '11A - Amended Tax Liability (Advance Received)': {
        'heading': '11A - Amendment to advances received',
        'layout': 'full',
        'row_pattern': 'Amended amount - Total',
    },
    '11B - Amendment of Adjustment of Advances': {
        'heading': '11B - Amendment to advances adjusted',
        'layout': 'full',
        'row_pattern': 'Amended amount - Total',
    },
    '11A(1), 11A(2) - Tax Liability (Advances Received)': {
        'heading': '11A(1), 11A(2) - Advances received',
        'layout': 'full',
        'row_pattern': 'Total',
    },
    '11B(1), 11B(2) - Adjustment of Advances': {
        'heading': '11B(1), 11B(2) - Advance amount received in earlier tax period',
        'layout': 'full',
        'row_pattern': 'Total',
    },
    '12 - HSN-wise summary of outward supplies': {
        'heading': '12 - HSN-wise summary of outward supplies',
        'layout': 'full',
        'row_pattern': r'^Total\s+\d+\s+NA\s+',
        'hsn': True,
    },
}


def load_template(template_path):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    groups = []
    month_col = 2
    for c in range(2, ws.max_column + 1):
        title = ws.cell(3, c).value
        if title:
            # The "Month" header (column B) is a single label column, not a
            # data group with sub-headers. Treating it as a group was the
            # cause of the Month column being overwritten with 0 on every
            # row (its blank sub-header fell into the default "0" branch of
            # write_group_row right after the real month value was set).
            if str(title).strip().lower() == 'month':
                month_col = c
                continue
            end = c
            while end + 1 <= ws.max_column and ws.cell(3, end + 1).value is None:
                end += 1
            groups.append({
                'start': c,
                'end': end,
                'title': str(title),
                'headers': [str(ws.cell(4, x).value or '').strip() for x in range(c, end + 1)],
            })
    return wb, ws, groups, month_col


def norm(s):
    if s is None:
        return ''
    s = str(s)
    s = s.replace('\u2013', '-').replace('\u2014', '-').replace('\u2011', '-')
    s = s.replace('₹', ' ')
    return re.sub(r'[ \t]+', ' ', s).strip()


def row_numbers(line):
    line = norm(line)
    line = line.replace('I0', '0').replace('N0.00', '0.00').replace('F0', '0')
    return re.findall(r'(?<![A-Za-z])[-]?[0-9][0-9,]*(?:\.[0-9]+)?', line)


def num(s):
    try:
        s = str(s).replace(',', '').strip()
        return float(s) if s else 0.0
    except (TypeError, ValueError):
        return 0.0


def zero():
    return {
        'records': 0,
        'taxable': 0.0,
        'igst': 0.0,
        'cgst': 0.0,
        'sgst': 0.0,
        'cess': 0.0,
        'invoice': 0.0,
        'found': False,
        'source_line': '',
    }


def parse_row(line, layout='full', hsn=False):
    """Parse a single filed-summary row using the exact PDF column layout."""
    if not line:
        return zero()
    ns = row_numbers(line)
    if not ns:
        return zero()

    # The first number is the record count in all mapped summary rows.
    records = int(round(num(ns[0])))
    values = [num(x) for x in ns[1:]]
    d = zero()
    d['records'] = records
    d['found'] = True
    d['source_line'] = line

    # "Value" in the GSTR-1 summary is mapped as taxable value for this tool.
    d['taxable'] = values[0] if len(values) >= 1 else 0.0
    if layout == 'value_igst_cess':
        d['igst'] = values[1] if len(values) >= 2 else 0.0
        d['cess'] = values[2] if len(values) >= 3 else 0.0
    elif layout == 'value_igst':
        d['igst'] = values[1] if len(values) >= 2 else 0.0
    else:  # value + IGST + CGST + SGST + Cess
        d['igst'] = values[1] if len(values) >= 2 else 0.0
        d['cgst'] = values[2] if len(values) >= 3 else 0.0
        d['sgst'] = values[3] if len(values) >= 4 else 0.0
        d['cess'] = values[4] if len(values) >= 5 else 0.0

    d['invoice'] = d['taxable'] + d['igst'] + d['cgst'] + d['sgst'] + d['cess']
    return d


def find_heading(lines, heading, heading_alt=None):
    targets = [norm(heading).lower()]
    if heading_alt:
        targets.append(norm(heading_alt).lower())
    for i, line in enumerate(lines):
        low = line.lower()
        if any(t in low for t in targets):
            return i
    return None


def find_row_after_heading(lines, heading, row_pattern=None, heading_alt=None, row_label=None, max_window=12):
    """Find a data row only inside the relevant section window."""
    idx = find_heading(lines, heading, heading_alt)
    if idx is None:
        return None

    start = idx + 1
    end = min(len(lines), start + max_window)
    window = lines[start:end]

    if row_label:
        for line in window:
            if norm(row_label).lower() in line.lower():
                return line

    if row_pattern:
        pattern = re.compile(row_pattern, re.I) if not isinstance(row_pattern, re.Pattern) else row_pattern
        for line in window:
            if pattern.search(line):
                return line

    # Default: first summary Total row. Avoid differential rows where possible.
    for line in window:
        low = line.lower()
        if ('net differential amount' in low or 'net differential' in low) and 'total' not in low.split('net differential')[0]:
            continue
        if re.search(r'\bTotal\b|\bNet Total\b|Amended amount', line, re.I):
            return line
    return None


def aggregate(parsed_rows):
    d = zero()
    found = False
    sources = []
    for p in parsed_rows:
        if not p.get('found'):
            continue
        found = True
        for k in ['records', 'taxable', 'igst', 'cgst', 'sgst', 'cess', 'invoice']:
            d[k] += p.get(k, 0) or 0
        if p.get('source_line'):
            sources.append(p['source_line'])
    d['found'] = found
    d['source_line'] = ' || '.join(sources)
    return d


def extract_pdf(pdf_bytes):
    lines = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            raw = page.extract_text() or ''
            for raw_line in raw.splitlines():
                line = norm(raw_line)
                if line:
                    lines.append(line)

    text = '\n'.join(lines)
    fy = re.search(r'Financial year\s+(\d{4}-\d{2})', text, re.I)
    tp = re.search(r'Tax period\s+([A-Za-z]+)', text, re.I)
    gstin = re.search(r'1 GSTIN\s+([0-9A-Z]{15})', text, re.I)
    arn = re.search(r'\(c\) ARN\s+([A-Z0-9]+)', text, re.I)
    arn_date = re.search(r'\(d\) ARN date\s+(\d{2}/\d{2}/\d{4})', text, re.I)

    month = None
    if fy and tp:
        try:
            month_num = datetime.strptime(tp.group(1)[:3], '%b').month
            start_year = int(fy.group(1)[:4])
            year = start_year if month_num >= 4 else start_year + 1
            month = datetime(year, month_num, 1)
        except Exception:
            month = None

    data = {
        'fy': fy.group(1) if fy else None,
        'tax_period': tp.group(1) if tp else None,
        'gstin': gstin.group(1) if gstin else None,
        'arn': arn.group(1) if arn else None,
        'arn_date': arn_date.group(1) if arn_date else None,
        'month': month,
        'sections': {},
        'audit': [],
    }

    # --- Section mappings ----------------------------------------------------
    for group, cfg in GROUP_CONFIG.items():
        if cfg.get('type') == 'table8':
            values = {}
            idx = find_heading(lines, '8 - Nil rated, exempted and non GST outward supplies')
            if idx is not None:
                for label, key in [('- Nil', 'nil'), ('- Exempted', 'exempted'), ('- Non-GST', 'non_gst')]:
                    found_line = None
                    for line in lines[idx + 1:min(len(lines), idx + 10)]:
                        if line.lower().startswith(label.lower()):
                            found_line = line
                            ns = row_numbers(line)
                            values[key] = num(ns[-1]) if ns else 0.0
                            break
                    if found_line:
                        data['audit'].append((group, found_line, 'special Table 8 value'))
            data['sections'][group] = {
                'nil': values.get('nil', 0.0),
                'exempted': values.get('exempted', 0.0),
                'non_gst': values.get('non_gst', 0.0),
                'found': idx is not None,
            }
            continue

        if cfg.get('type') == 'aggregate':
            parts = []
            for heading, layout, row_pattern in cfg['sections']:
                line = find_row_after_heading(lines, heading, row_pattern=row_pattern)
                parsed = parse_row(line, layout=layout) if line else zero()
                parts.append(parsed)
                if line:
                    data['audit'].append((group, line, heading))
            data['sections'][group] = aggregate(parts)
            continue

        heading = cfg.get('heading')
        line = find_row_after_heading(
            lines,
            heading,
            row_pattern=cfg.get('row_pattern'),
            heading_alt=cfg.get('heading_alt'),
            row_label=cfg.get('row_label'),
            max_window=15,
        )
        parsed = parse_row(line, layout=cfg.get('layout', 'full'), hsn=cfg.get('hsn', False)) if line else zero()
        data['sections'][group] = parsed
        if line:
            data['audit'].append((group, line, heading))

    # Filed-copy reconciliation: Table 6A total should agree with the final
    # outward-supply liability when no other taxable outward tables contain
    # values. Keep this as an audit flag rather than changing extracted data.
    total_liability = None
    for line in lines:
        m = re.search(r'Total Liability \(Outward supplies other than Reverse charge\)\s+([-]?[0-9,]+(?:\.[0-9]+)?)', line, re.I)
        if m:
            total_liability = num(m.group(1))
            break
    data['total_liability'] = total_liability
    six_a = data['sections'].get('6A - Exports Invoices', zero())
    data['reconciliation'] = {
        'total_liability': total_liability,
        '6A_invoice_value': six_a.get('invoice', 0.0),
        'difference': (total_liability - six_a.get('invoice', 0.0)) if total_liability is not None else None,
    }

    return data


def copy_row_style(ws, source_row, target_row):
    # NOTE: openpyxl's cell._style is a mutable array shared by reference,
    # not copied by value. Assigning `dst._style = src._style` directly (as
    # earlier versions of this helper did) makes the two cells point at the
    # SAME underlying style object, so a later change to one row's format
    # (e.g. the Total row) silently mutates every other row that ever
    # shared that reference. Copying each style component individually
    # avoids that bleed-through.
    for c in range(1, ws.max_column + 1):
        src = ws.cell(source_row, c)
        dst = ws.cell(target_row, c)
        if src.has_style:
            if src.font:
                dst.font = src.font.copy()
            if src.border:
                dst.border = src.border.copy()
            if src.fill:
                dst.fill = src.fill.copy()
            if src.alignment:
                dst.alignment = src.alignment.copy()
            if src.protection:
                dst.protection = src.protection.copy()
        if src.number_format:
            dst.number_format = src.number_format
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def write_group_row(ws, row, group, d):
    headers = group['headers']
    for i, header in enumerate(headers):
        c = group['start'] + i
        h = header.strip().lower()
        if h.startswith('no. of records'):
            value = d.get('records', 0)
        elif h == 'invoice value':
            # ALWAYS use all source tax components, even if some tax columns
            # are not exposed in this particular output group.
            value = d.get('invoice', 0)
        elif h == 'taxable value':
            value = d.get('taxable', 0)
        elif h == 'igst':
            value = d.get('igst', 0)
        elif h == 'cgst':
            value = d.get('cgst', 0)
        elif h == 'sgst':
            value = d.get('sgst', 0)
        elif h == 'cess':
            value = d.get('cess', 0)
        elif h == 'nil':
            value = d.get('nil', 0)
        elif h == 'exempted':
            value = d.get('exempted', 0)
        elif h == 'non-gst':
            value = d.get('non_gst', 0)
        else:
            value = 0
        ws.cell(row, c).value = value


def build_workbook(records, template_path):
    wb, ws, groups, month_col = load_template(template_path)

    # Clear old data rows but preserve the template headers/formatting.
    start_row = 5
    old_max_row = ws.max_row
    for r in range(start_row, old_max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    row = start_row
    audit_rows = []
    for rec in sorted(records, key=lambda x: x['month'] or datetime.max):
        if row > start_row:
            copy_row_style(ws, start_row, row)
        # Write the month as a real Excel date so it sorts/filters correctly
        # and displays via the template's own accounting-style "mmm-yy"
        # format (e.g. Apr-25) rather than as plain text.
        month_date = rec.get('month')
        ws.cell(row, month_col).value = month_date if month_date else None
        ws.cell(row, month_col).number_format = 'mmm-yy'

        for g in groups:
            d = rec['sections'].get(g['title'], zero())
            write_group_row(ws, row, g, d)

        for group, source_line, heading in rec.get('audit', []):
            audit_rows.append([
                rec.get('file_name', ''),
                rec.get('tax_period', ''),
                group,
                heading,
                source_line,
            ])
        row += 1

    total_row = row
    if records:
        copy_row_style(ws, start_row, total_row)
    ws.cell(total_row, month_col).value = 'Total'
    ws.cell(total_row, month_col).number_format = '@'

    # Total each output column, including record counts and Table 8 values.
    for g in groups:
        for i, header in enumerate(g['headers']):
            c = g['start'] + i
            h = header.strip().lower()
            if h == 'month':
                continue
            total = 0.0
            for r in range(start_row, total_row):
                v = ws.cell(r, c).value
                if isinstance(v, (int, float)):
                    total += v
            ws.cell(total_row, c).value = int(total) if h.startswith('no. of records') else total

    # --- Professional styling ------------------------------------------------
    # Make the Total row stand out: bold text, a heavier top border, and a
    # light highlight fill that complements the template's maroon header.
    if records:
        total_font = Font(name='Trebuchet MS', size=10, bold=True, color='FF98002E')
        total_fill = PatternFill(fill_type='solid', fgColor='FFF2E2E6')
        top_border = Border(top=Side(style='double', color='FF98002E'))
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(total_row, c)
            cell.font = total_font
            cell.fill = total_fill
            cell.border = Border(
                top=top_border.top,
                bottom=cell.border.bottom,
                left=cell.border.left,
                right=cell.border.right,
            )

    # Right-align + apply thousands separators to numeric data cells so large
    # figures are easy to scan, and keep column widths readable.
    number_fmt = '#,##0.00;[RED]-#,##0.00'
    count_fmt = '#,##0'
    for g in groups:
        for i, header in enumerate(g['headers']):
            c = g['start'] + i
            h = header.strip().lower()
            fmt = count_fmt if h.startswith('no. of records') else number_fmt
            for r in range(start_row, total_row + 1):
                cell = ws.cell(r, c)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = fmt
        width = 16 if any('value' in h.lower() for h in g['headers']) else 13
        for c in range(g['start'], g['end'] + 1):
            ws.column_dimensions[get_column_letter(c)].width = max(
                ws.column_dimensions[get_column_letter(c)].width or 0, width
            )
    ws.column_dimensions[get_column_letter(month_col)].width = 12

    # The template ships with 12 pre-formatted blank month rows plus its own
    # leftover "Total" row/fill. Once we've written the real Total row,
    # remove any now-unused template rows below it so the sheet doesn't show
    # stray empty bordered rows or an orphaned Total-style band underneath.
    if old_max_row > total_row:
        ws.delete_rows(total_row + 1, old_max_row - total_row)

    ws.freeze_panes = f'{get_column_letter(month_col + 1)}{start_row}'
    ws.auto_filter.ref = (
        f'{get_column_letter(month_col)}4:{get_column_letter(ws.max_column)}{total_row}'
    )

    # Add a transparent audit sheet so every extracted value can be traced back
    # to the exact source row in the filed GSTR-1 PDF.
    if 'Extraction Audit' in wb.sheetnames:
        del wb['Extraction Audit']
    audit = wb.create_sheet('Extraction Audit')
    headers = ['PDF File', 'Tax Period', 'Excel Group', 'GSTR-1 Section', 'Source Row']
    for c, h in enumerate(headers, start=1):
        audit.cell(1, c).value = h
        audit.cell(1, c).font = Font(bold=True)
    for r_idx, values in enumerate(audit_rows, start=2):
        for c_idx, value in enumerate(values, start=1):
            audit.cell(r_idx, c_idx).value = value
    audit.freeze_panes = 'A2'
    audit.auto_filter.ref = f'A1:E{max(1, len(audit_rows)+1)}'
    for col, width in {'A':35, 'B':15, 'C':45, 'D':70, 'E':90}.items():
        audit.column_dimensions[col].width = width

    if 'Read Me' in wb.sheetnames:
        del wb['Read Me']
    meta = wb.create_sheet('Read Me', 0)
    meta.sheet_view.showGridLines = False

    maroon = 'FF98002E'
    title_font = Font(name='Trebuchet MS', size=16, bold=True, color=maroon)
    subtitle_font = Font(name='Trebuchet MS', size=10, italic=True, color='FF666666')
    label_font = Font(name='Trebuchet MS', size=10, bold=True, color=maroon)
    body_font = Font(name='Trebuchet MS', size=10, color='FF333333')
    link_font = Font(name='Trebuchet MS', size=10, bold=True, color='FF1155CC', underline='single')

    meta['A1'] = 'GSTR-1 PDF to Excel Analyzer'
    meta['A1'].font = title_font
    meta['A2'] = 'Powered by pushpakkumar.com'
    meta['A2'].font = subtitle_font
    meta['A2'].hyperlink = 'https://pushpakkumar.com'

    # Surface the GSTIN(s) found across the uploaded PDFs. If more than one
    # distinct GSTIN is present, that's a real accuracy risk (PDFs from two
    # different companies accidentally batched together), so flag it clearly
    # instead of silently combining their totals.
    gstins = sorted({r.get('gstin') for r in records if r.get('gstin')})
    if len(gstins) == 1:
        gstin_label = 'GSTIN'
        gstin_value = gstins[0]
        gstin_font = label_font
    elif len(gstins) > 1:
        gstin_label = 'GSTIN — WARNING'
        gstin_value = f'{len(gstins)} different GSTINs found in this upload: ' + ', '.join(gstins) + \
            ' — totals below mix multiple companies. Re-run with one company\'s PDFs at a time.'
        gstin_font = Font(name='Trebuchet MS', size=10, bold=True, color='FFCC0000')
    else:
        gstin_label = 'GSTIN'
        gstin_value = 'Not found in the uploaded PDF(s)'
        gstin_font = label_font

    rows = [
        ('Invoice Value rule', 'Taxable Value + IGST + CGST + SGST + Cess'),
        ('Source', 'Filed GSTR-1 PDFs uploaded in the ZIP'),
        ('Mapping rule', 'Each Excel group is mapped to a specific GSTR-1 section/row. The PDF Value field is treated as Taxable Value.'),
        ('Audit', 'Extraction Audit sheet records the exact source row used for each mapped section.'),
        ('Reconciliation', 'The tool also captures the filed Total Liability and the difference against extracted taxable-outward invoice value; it does not overwrite source data.'),
    ]
    r = 4
    meta.cell(r, 1, gstin_label).font = gstin_font
    meta.cell(r, 2, gstin_value).font = gstin_font
    meta.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='top')
    r += 1
    for label, body in rows:
        meta.cell(r, 1, label).font = label_font
        meta.cell(r, 2, body).font = body_font
        meta.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='top')
        r += 1

    r += 1
    meta.cell(r, 1, 'Website').font = label_font
    link_cell = meta.cell(r, 2, 'pushpakkumar.com')
    link_cell.font = link_font
    link_cell.hyperlink = 'https://pushpakkumar.com'

    meta.column_dimensions['A'].width = 20
    meta.column_dimensions['B'].width = 90
    for rr in range(4, r + 1):
        meta.row_dimensions[rr].height = 18

    # Small branded banner on the main data sheet header area so the
    # workbook is clearly attributed wherever it's opened/printed.
    ws['A1'] = 'pushpakkumar.com'
    ws['A1'].font = Font(name='Trebuchet MS', size=9, italic=True, color='FF98002E')
    ws['A1'].hyperlink = 'https://pushpakkumar.com'

    return wb


def main():
    import streamlit as st

    st.set_page_config(
        page_title="Pushpak Kumar | GSTR-1 Analyzer",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="collapsed",
    )

    # Professional Pushpak Kumar branding
    st.markdown("""
    <style>
    .stApp { background:#f7f8fb; }
    .main .block-container {
        max-width:1180px;
        padding-top:2rem;
        padding-bottom:3rem;
    }

    .hero {
        background:linear-gradient(135deg,#98002e 0%,#65001f 100%);
        border-radius:20px;
        padding:30px 34px;
        color:#fff;
        margin-bottom:22px;
        box-shadow:0 10px 30px rgba(90,0,30,.16);
    }

    .brand-row {
        display:flex;
        align-items:center;
        gap:12px;
        margin-bottom:20px;
    }

    .brand-logo {
        width:44px;
        height:44px;
        border-radius:12px;
        background:rgba(255,255,255,.14);
        border:1px solid rgba(255,255,255,.25);
        display:flex;
        align-items:center;
        justify-content:center;
        font-size:17px;
        font-weight:800;
        letter-spacing:-1px;
    }

    .brand-name {
        font-size:14px;
        font-weight:750;
        letter-spacing:.4px;
        color:#fff;
    }

    .brand-tagline {
        font-size:11px;
        color:rgba(255,255,255,.72);
        margin-top:2px;
    }

    .hero h1 {
        margin:0 0 8px 0;
        font-size:31px;
        font-weight:750;
        letter-spacing:-.6px;
    }

    .hero p {
        margin:0;
        color:rgba(255,255,255,.88);
        font-size:14px;
    }

    .card {
        background:#fff;
        border:1px solid #e7e9ef;
        border-radius:16px;
        padding:22px;
        margin-bottom:18px;
        box-shadow:0 3px 14px rgba(20,20,40,.05);
    }

    .section-title {
        font-size:18px;
        font-weight:700;
        color:#242733;
        margin-bottom:4px;
    }

    .section-subtitle {
        color:#707582;
        font-size:13px;
        margin-bottom:14px;
    }

    .workflow {
        display:flex;
        gap:10px;
        align-items:center;
        color:#424752;
        font-size:13px;
    }

    .step {
        display:flex;
        align-items:center;
        gap:8px;
        flex:1;
    }

    .step-num {
        width:28px;
        height:28px;
        border-radius:50%;
        background:#f3e6eb;
        color:#98002e;
        display:flex;
        align-items:center;
        justify-content:center;
        font-weight:750;
    }

    .workflow-line {
        height:1px;
        background:#dddfe6;
        flex:1;
    }

    .status-card {
        background:#fff;
        border-left:4px solid #98002e;
        border-radius:10px;
        padding:13px 16px;
        margin:12px 0 16px;
        box-shadow:0 2px 9px rgba(20,20,40,.04);
    }

    .brand-chip {
        display:inline-flex;
        align-items:center;
        gap:7px;
        background:#f3e6eb;
        color:#98002e;
        border-radius:999px;
        padding:6px 11px;
        font-size:11px;
        font-weight:750;
        margin-bottom:11px;
    }

    .brand-dot {
        width:6px;
        height:6px;
        border-radius:50%;
        background:#98002e;
    }

    .footer {
        text-align:center;
        color:#858994;
        font-size:12px;
        padding-top:22px;
    }

    .footer-name {
        color:#98002e;
        font-weight:750;
    }

    div[data-testid="stFileUploader"] {
        background:#fbfbfd;
        border:1px dashed #c9ccd6;
        border-radius:13px;
        padding:8px;
    }

    .stButton > button,
    .stDownloadButton > button {
        border-radius:10px;
        min-height:44px;
        font-weight:650;
    }
    </style>
    """, unsafe_allow_html=True)

    # Header
    st.markdown("""
    <div class="hero">
        <div class="brand-row">
            <div class="brand-logo">PK</div>
            <div>
                <div class="brand-name">PUSHPAK KUMAR</div>
                <div class="brand-tagline">GSTR-1 Analysis &amp; Reporting</div>
            </div>
        </div>
        <h1>GSTR-1 PDF → Excel Analyzer</h1>
        <p>Automate filed GSTR-1 analysis, section mapping and Excel reporting.</p>
    </div>
    """, unsafe_allow_html=True)

    # Workflow
    st.markdown("""
    <div class="card">
        <div class="workflow">
            <div class="step"><span class="step-num">1</span><b>Upload ZIP</b></div>
            <div class="workflow-line"></div>
            <div class="step"><span class="step-num">2</span><b>Analyze PDFs</b></div>
            <div class="workflow-line"></div>
            <div class="step"><span class="step-num">3</span><b>Download Excel</b></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Upload
    st.markdown("""
    <div class="card">
        <div class="brand-chip"><span class="brand-dot"></span> PUSHPAK KUMAR • ANALYZER</div>
        <div class="section-title">Upload GSTR-1 files</div>
        <div class="section-subtitle">
            Upload a ZIP containing filed GSTR-1 PDF copies. One PDF per tax period is recommended.
        </div>
    """, unsafe_allow_html=True)

    template = Path(__file__).with_name("Sample_format.xlsx")
    zip_file = st.file_uploader(
        "Upload ZIP containing GSTR-1 filed PDFs",
        type=["zip"],
        label_visibility="collapsed",
    )

    st.markdown("</div>", unsafe_allow_html=True)

    if not zip_file:
        st.info("Tip: Keep PDFs for the same GSTIN together in one ZIP for clean monthly reporting.")
        st.markdown("""
        <div class="footer">
            Invoice Value = Taxable Value + IGST + CGST + SGST + Cess
            <br><br>
            Designed &amp; developed by <span class="footer-name">Pushpak Kumar</span>
        </div>
        """, unsafe_allow_html=True)
        return

    # Inspect ZIP before processing
    try:
        with zipfile.ZipFile(zip_file) as z_preview:
            pdf_names = [
                n for n in z_preview.namelist()
                if n.lower().endswith(".pdf") and not n.endswith("/")
            ]
    except zipfile.BadZipFile:
        st.error("The uploaded file is not a valid ZIP file.")
        return

    if not pdf_names:
        st.error("No PDF files were found in the ZIP.")
        return

    size_mb = len(zip_file.getvalue()) / (1024 * 1024)

    st.markdown(
        f"""
        <div class="status-card">
            <b>Ready to analyze</b><br>
            <span style="color:#707582;font-size:13px;">
                {len(pdf_names)} PDF file(s) &nbsp; • &nbsp; {size_mb:.2f} MB
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button("🔍  Analyze GSTR-1 ZIP", type="primary", use_container_width=True):
        records = []
        errors = []

        progress = st.progress(0, text="Starting analysis…")
        status = st.empty()

        total_files = len(pdf_names)

        with zipfile.ZipFile(zip_file) as z:
            for i, name in enumerate(pdf_names, start=1):
                status.markdown(
                    f"**Processing {i} of {total_files}:** `{Path(name).name}`"
                )

                try:
                    d = extract_pdf(z.read(name))

                    if not d.get("month"):
                        raise ValueError(
                            "Could not identify Financial Year / Tax Period"
                        )

                    d["file_name"] = name
                    records.append(d)

                except Exception as exc:
                    errors.append((name, str(exc)))

                progress.progress(
                    i / total_files,
                    text=f"Analyzing PDF {i} of {total_files}",
                )

        progress.progress(1.0, text="Analysis complete ✓")
        status.empty()

        if errors:
            st.warning(f"{len(errors)} file(s) could not be parsed.")
            with st.expander("View parsing errors"):
                st.dataframe(
                    [{"File": n, "Error": e} for n, e in errors],
                    use_container_width=True,
                    hide_index=True,
                )

        if not records:
            st.error("No GSTR-1 PDFs could be successfully parsed.")
            return

        st.success(
            f"Successfully parsed {len(records)} of {len(pdf_names)} PDF(s)."
        )

        # Summary cards
        gstins = sorted(
            {r.get("gstin") for r in records if r.get("gstin")}
        )
        fys = sorted(
            {r.get("fy") for r in records if r.get("fy")}
        )

        a, b, c = st.columns(3)
        with a:
            st.metric("PDFs Parsed", len(records))
        with b:
            st.metric("GSTINs Found", len(gstins))
        with c:
            st.metric("Financial Year(s)", len(fys))

        if len(gstins) > 1:
            st.warning(
                "Multiple GSTINs were found in this ZIP. For clean reporting, "
                "analyze one GSTIN at a time."
            )

        # Period preview
        st.markdown(
            """
            <div class="section-title" style="margin-top:22px;">
                Extracted periods
            </div>
            <div class="section-subtitle">
                Review the detected tax periods before downloading.
            </div>
            """,
            unsafe_allow_html=True,
        )

        preview = []
        for r in sorted(
            records,
            key=lambda x: x["month"] or datetime.max
        ):
            preview.append({
                "Month": r["month"].strftime("%b-%y")
                    if r.get("month") else "",
                "GSTIN": r.get("gstin"),
                "Tax Period": r.get("tax_period"),
                "Financial Year": r.get("fy"),
                "File": Path(r["file_name"]).name,
            })

        st.dataframe(
            preview,
            use_container_width=True,
            hide_index=True,
        )

        # Excel build progress
        build_progress = st.progress(
            0,
            text="Preparing Excel workbook…"
        )

        build_progress.progress(
            25,
            text="Loading Excel template…"
        )

        build_progress.progress(
            50,
            text="Applying GSTR-1 mappings…"
        )

        wb = build_workbook(records, str(template))

        build_progress.progress(
            75,
            text="Creating extraction audit trail…"
        )

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)

        build_progress.progress(
            100,
            text="Excel report ready ✓"
        )

        st.markdown(
            """
            <div class="card">
                <div class="section-title">Your report is ready</div>
                <div class="section-subtitle">
                    The workbook includes the mapped report, Extraction Audit
                    and Read Me sheets.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.download_button(
            "⬇️  Download GSTR1_Analyzed.xlsx",
            out.getvalue(),
            file_name="GSTR1_Analyzed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

    st.markdown("""
    <div class="footer">
        GSTR-1 PDF → Excel Analyzer
        &nbsp; • &nbsp;
        Invoice Value = Taxable Value + IGST + CGST + SGST + Cess
        <br><br>
        Designed &amp; developed by
        <span class="footer-name">Pushpak Kumar</span>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
