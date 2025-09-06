
import io
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from odoo import http
from odoo.http import request, content_disposition
from collections import defaultdict


class StockInOutReportController(http.Controller):

    @http.route('/web/binary/export_xlsx', type='http', auth='user')

    def export_xlsx(self, wizard_id=None, **kwargs):
        wizard = request.env['bk.stock.inout.wizard'].browse(int(wizard_id))
        if not wizard.exists():
            return request.not_found()

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        sheet = workbook.add_worksheet("Stock Report")

        # Formats
        title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#D9D9D9', 'border': 1,
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        })
        number_fmt = workbook.add_format({'border': 1, 'font_size': 9, 'num_format': '#,##0.00', 'align': 'right'})
        alt_row_fmt = workbook.add_format({'border': 1, 'font_size': 9, 'num_format': '#,##0.00',
                                           'align': 'right', 'bg_color': '#F9F9F9'})
        total_fmt = workbook.add_format({'bold': True, 'bg_color': '#E0E0E0', 'border': 1,
                                         'num_format': '#,##0.00', 'align': 'right'})
        text_fmt = workbook.add_format({'border': 1, 'font_size': 9, 'align': 'left'})

        # 1. Report Title
        sheet.merge_range('A1:H1', "Stock Movement Balance Report", title_fmt)

        # 2. Filter Conditions
        sheet.write('A2', f"Period: {wizard.date_start} → {wizard.date_end}")
        sheet.write('A3', f"Locations: {', '.join([loc.display_name for loc in wizard.location_ids]) or 'All'}")
        sheet.write('A4', f"Products/Categories: "
                          f"{', '.join([p.display_name for p in wizard.product_ids]) if wizard.product_ids else ''}"
                          f"{', '.join([c.display_name for c in wizard.categ_ids]) if wizard.categ_ids else 'All'}")

        # 3. Data
        if wizard.report_type == 'detailed':
            headers, lines = self._compute_detailed_lines(wizard)
        else:  # summary
            headers, lines = self._compute_summary_lines(wizard)

        # write headers
        for col, h in enumerate(headers):
            sheet.write(5, col, h, header_fmt)
            sheet.set_column(col, col, 14)  # set default width

        # write data rows
        row = 6
        for idx, line in enumerate(lines):
            row_fmt = alt_row_fmt if idx % 2 else number_fmt
            for col, val in enumerate(line):
                if isinstance(val, (int, float)):
                    sheet.write(row, col, val, row_fmt)
                else:
                    sheet.write(row, col, val, text_fmt)
            row += 1

        # optional totals row at end
        if lines and any(isinstance(v, (int, float)) for v in lines[0][3:]):  # detect numeric cols
            sheet.write(row, 0, "TOTAL", total_fmt)
            for col in range(3, len(headers)):
                col_letter = chr(65 + col)
                sheet.write_formula(row, col, f"SUM({col_letter}7:{col_letter}{row})", total_fmt)

        workbook.close()
        output.seek(0)

        file_name = "stock_report.xlsx"
        return request.make_response(
            output.read(),
            headers=[
                ('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
                ('Content-Disposition', content_disposition(file_name))
            ]
        )

    def _compute_lines(self, wizard):
        domain = [
            ("date", "<=", wizard.date_end),
        ]
        if wizard.state != 'all':
            domain.append(("state", "=", wizard.state))

        if wizard.product_ids:
            domain.append(("product_id", "in", wizard.product_ids.ids))
        elif wizard.categ_ids:
            domain.append(("product_id.categ_id", "in", wizard.categ_ids.ids))
        if wizard.location_ids:
            domain += ["|", ("location_id", "in", wizard.location_ids.ids),
                       ("location_dest_id", "in", wizard.location_ids.ids)]

        moves = request.env["stock.move"].sudo().search(domain, order="product_id, location_id, date")

        data = defaultdict(lambda: defaultdict(float))

        for move in moves:
            product = move.product_id
            categ = product.categ_id.display_name
            location = move.location_id if move.location_id.usage == "internal" else move.location_dest_id
            location_name = location.display_name if location else "-"

            key = (product.id, location.id)
            d = data[key]
            d["product"] = product.display_name
            d["category"] = categ
            d["location"] = location_name

            qty = move.product_uom_qty

            # Initial balance (before period)
            if move.date.date() < wizard.date_start:
                if move.location_dest_id.usage == "internal":
                    d["initial"] += qty
                if move.location_id.usage == "internal":
                    d["initial"] -= qty
                continue

            # Classification within period
            if move.location_id.usage == "supplier" and move.location_dest_id.usage == "internal":
                d["purchased"] += qty
            elif move.location_id.usage == "internal" and move.location_dest_id.usage == "supplier":
                d["return_sup"] += qty
            elif move.location_id.usage == "internal" and move.location_dest_id.usage == "customer":
                d["sold"] += qty
            elif move.location_id.usage == "customer" and move.location_dest_id.usage == "internal":
                d["return_cust"] += qty
            elif move.location_id.usage == "internal" and move.location_dest_id.usage in ("inventory", "production"):
                d["losses"] += qty
            elif move.location_id.usage in ("inventory", "production") and move.location_dest_id.usage == "internal":
                d["gains"] += qty

        # Compute ending and running balance
        lines = []
        for k, d in data.items():
            # Ending balance
            d["ending"] = (
                    d.get("initial", 0)
                    + d.get("purchased", 0)
                    + d.get("return_cust", 0)
                    + d.get("gains", 0)
                    - d.get("sold", 0)
                    - d.get("return_sup", 0)
                    - d.get("losses", 0)
            )
            # Running balance = initial + net movements per period
            d["running_balance"] = d["ending"]  # or keep same as ending

            # NEW: Net Incoming / Net Outgoing
            # NEW: Net Incoming / Net Outgoing
            # d["net_incoming"] = d.get("purchased", 0) + d.get("gains", 0) - d.get("losses", 0) - d.get("return_sup", 0)
            # d["net_outgoing"] = d.get("sold", 0) + d.get("losses", 0) - d.get("return_cust", 0)
            d["net_incoming"] = d.get("purchased", 0) + d.get("gains", 0) - d.get("return_sup", 0)
            d["net_outgoing"] = -1 * (d.get("sold", 0) + d.get("losses", 0) - d.get("return_cust", 0))

            lines.append(d)

        return lines

    def _compute_detailed_lines(self, wizard):
        """Detailed: per product + internal location with full breakdown."""
        # --- Build domain
        domain = [("date", "<=", wizard.date_end)]
        # state filter (ready→assigned)
        state_val = getattr(wizard, "state", "done")
        print(f"state val:{state_val}")
        if state_val and state_val != "all":
            domain.append(("state", "=", "assigned" if state_val == "ready" else state_val))

        # product/category filter
        if wizard.product_ids:
            domain.append(("product_id", "in", wizard.product_ids.ids))
        elif wizard.categ_ids:
            domain.append(("product_id.categ_id", "in", wizard.categ_ids.ids))
        # location filter
        if wizard.location_ids:
            domain += ["|", ("location_id", "in", wizard.location_ids.ids),
                       ("location_dest_id", "in", wizard.location_ids.ids)]
        moves = request.env["stock.move"].sudo().search(domain, order="product_id, location_id, location_dest_id, date, id")
        # accumulators per (product, internal_location)
        agg = defaultdict(lambda: {
            "product": "",
            "category": "",
            "location": "",
            "initial": 0.0,
            "purchased": 0.0,
            "return_sup": 0.0,
            "sold": 0.0,
            "return_cust": 0.0,
            "losses": 0.0,
            "gains": 0.0,
        })

        def pick_internal_loc(m):
            """Choose the internal location relevant for this move row."""
            src, dst = m.location_id, m.location_dest_id
            # If user filtered locations, prefer the one among them
            if wizard.location_ids:
                if src.usage == "internal" and src in wizard.location_ids:
                    return src
                if dst.usage == "internal" and dst in wizard.location_ids:
                    return dst
            # Otherwise pick whichever side is internal (if any)
            if src.usage == "internal":
                return src
            if dst.usage == "internal":
                return dst
            return None

        for mv in moves:
            loc = pick_internal_loc(mv)
            # Skip moves that don't touch an internal location at all
            if not loc:
                continue

            prod = mv.product_id
            key = (prod.id, loc.id)
            d = agg[key]
            d["product"] = prod.display_name
            d["category"] = prod.categ_id.display_name
            d["location"] = loc.display_name
            qty = mv.product_uom_qty

            # Initial balance (all moves before start date)
            if mv.date.date() < wizard.date_start:
                if mv.location_dest_id.usage == "internal":
                    d["initial"] += qty
                if mv.location_id.usage == "internal":
                    d["initial"] -= qty
                continue

            # Period classifications
            src_u = mv.location_id.usage
            dst_u = mv.location_dest_id.usage

            if src_u == "supplier" and dst_u == "internal":
                d["purchased"] += qty
            elif src_u == "internal" and dst_u == "supplier":
                d["return_sup"] += qty
            elif src_u == "internal" and dst_u == "customer":
                d["sold"] += qty
            elif src_u == "customer" and dst_u == "internal":
                d["return_cust"] += qty
            elif src_u == "internal" and dst_u in ("inventory", "production"):
                d["losses"] += qty
            elif src_u in ("inventory", "production") and dst_u == "internal":
                d["gains"] += qty

        # finalize rows
        headers = [
            "Product", "Category", "Location","UoM",
            "Initial Balance",
            "Purchased", "Customer Returns", "Adjustments (Gain)",
            "Sold Qty", "Supplier Returns", "Adjustments (Loss)",
            "Total Incoming", "Total Outgoing",
            "Ending Balance",
            "Unit Cost", "Valuation"
        ]

        lines = []
        running_total = 0.0  # overall running (sheet-level); keep if you want a per-line running figure
        for (pid, lid), d in agg.items():  # iterate with keys + values
            net_incoming = d["purchased"] + d["gains"] - d["losses"] - d["return_sup"]
            net_outgoing = d["sold"] + d["losses"] - d["return_cust"]
            ending = d["initial"] + net_incoming - net_outgoing
            running_total += ending

            # fetch correct product (one record only)

            product = request.env["product.product"].browse(pid) if pid else False
            unit_cost = product.standard_price if product else 0.0
            uom = product.uom_id.name if product else ""
            valuation = ending * unit_cost

            lines.append([
                d["product"], d["category"], d["location"],
                uom,
                d["initial"],
                d["purchased"], d["return_cust"], d["gains"],
                d["sold"], d["return_sup"], d["losses"],
                net_incoming, net_outgoing,
                ending,
                unit_cost, valuation,
            ])

        return headers, lines

    def _compute_summary_lines(self, wizard):
        """Summary: per product (no location), with Initial, Net In, Net Out, Forecast."""
        # --- Build domain
        domain = [("date", "<=", wizard.date_end)]
        # state filter (ready→assigned)
        state_val = getattr(wizard, "state", "done")
        if state_val and state_val != "all":
            domain.append(("state", "=", "assigned" if state_val == "ready" else state_val))
        # product/category filter
        if wizard.product_ids:
            domain.append(("product_id", "in", wizard.product_ids.ids))
        elif wizard.categ_ids:
            domain.append(("product_id.categ_id", "in", wizard.categ_ids.ids))

        # location filter (still applied for which moves are considered)
        if wizard.location_ids:
            domain += ["|", ("location_id", "in", wizard.location_ids.ids),
                            ("location_dest_id", "in", wizard.location_ids.ids)]

        moves = request.env["stock.move"].sudo().search(domain, order="product_id, date, id")

        # accumulators per product
        agg = defaultdict(lambda: {
            "product": "",
            "category": "",
            "initial": 0.0,
            "purchased": 0.0,
            "return_sup": 0.0,
            "sold": 0.0,
            "return_cust": 0.0,
            "losses": 0.0,
            "gains": 0.0,
        })

        for mv in moves:
            # consider only moves that touch any internal location so we don't count external-to-external noise
            if mv.location_id.usage != "internal" and mv.location_dest_id.usage != "internal":
                continue

            prod = mv.product_id
            d = agg[prod.id]
            d["product"] = prod.display_name
            d["category"] = prod.categ_id.display_name
            qty = mv.product_uom_qty
            src_u = mv.location_id.usage
            dst_u = mv.location_dest_id.usage

            # Initial balance (before start date)
            if mv.date.date() < wizard.date_start:
                if dst_u == "internal":
                    d["initial"] += qty
                if src_u == "internal":
                    d["initial"] -= qty
                continue

            # Period classifications (same as detailed)
            if src_u == "supplier" and dst_u == "internal":
                d["purchased"] += qty
            elif src_u == "internal" and dst_u == "supplier":
                d["return_sup"] += qty
            elif src_u == "internal" and dst_u == "customer":
                d["sold"] += qty
            elif src_u == "customer" and dst_u == "internal":
                d["return_cust"] += qty
            elif src_u == "internal" and dst_u in ("inventory", "production"):
                d["losses"] += qty
            elif src_u in ("inventory", "production") and dst_u == "internal":
                d["gains"] += qty

        headers = [
            "Product", "Category", "UoM",
            "Initial Balance",
            "Total Incoming", "Total Outgoing",
            "Forecast Qty",
        ]

        lines = []
        for pid, d in agg.items():
            net_incoming = d["purchased"] + d["gains"] - d["losses"] - d["return_sup"]
            net_outgoing = d["sold"] + d["losses"] - d["return_cust"]
            forecast = d["initial"] + net_incoming - net_outgoing

            # fetch product to get UoM and cost
            product = request.env["product.product"].browse(pid) if pid else False
            uom = product.uom_id.name if product else ""
            unit_cost = product.standard_price if product else 0.0
            valuation = forecast * unit_cost

            lines.append([
                d["product"], d["category"],
                uom,  # 👈 add UoM here
                d["initial"],
                net_incoming, net_outgoing,
                forecast,
                unit_cost, valuation,  # 👈 added for completeness
            ])
        return headers, lines

