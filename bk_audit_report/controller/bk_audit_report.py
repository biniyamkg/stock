from odoo import http
from odoo.http import request
import io
import xlsxwriter
from datetime import datetime


class BkAuditReportController(http.Controller):
    @http.route(['/web/export_excel'], type='http', auth='user')
    def export_excel(self, wizard_id, **kwargs):
        wizard = request.env['bk.audit.report_wizard'].browse(int(wizard_id))
        if not wizard:
            return request.not_found()

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet("Purchase Order")

        # Formats
        bold = workbook.add_format({'bold': True})
        money_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1, 'font_size':9})
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
        section_fmt = workbook.add_format({'bold': True, 'bg_color': '#BDD7EE', 'border': 1})
        text_fmt = workbook.add_format({ "font_size": 9, 'border': 1})

        row = 0

        # ----------------------------
        # Helper: sum journal lines by account + currency + reference
        # ----------------------------
        def sum_journal_lines(moves, reference=None):
            debit_sum = {}
            credit_sum = {}
            currency_map = {}
            for move in moves.filtered(lambda m: m.state == 'posted'):
                for l in move.line_ids:
                    key = (l.account_id.name, l.currency_id.name or move.currency_id.name or 'Company Currency')
                    debit_sum[key] = debit_sum.get(key, 0.0) + l.debit
                    credit_sum[key] = credit_sum.get(key, 0.0) + l.credit
                    currency_map[key] = l.currency_id.name or move.currency_id.name or 'Company Currency'
            accounts = set(debit_sum.keys()) | set(credit_sum.keys())
            return [
                {
                    'account': a[0],
                    # 'currency': currency_map[a],
                    'debit': debit_sum.get(a, 0.0),
                    'credit': credit_sum.get(a, 0.0),
                    'remark': reference or ', '.join(moves.mapped('name'))
                }
                for a in accounts
            ]

        # ----------------------------
        # Helper: write journal summary
        # ----------------------------
        def write_summary(title, summary):
            nonlocal row
            if summary:
                row += 1
                worksheet.write(row, 0, title, section_fmt)
                row += 1
                # headers = ["Account", "Currency", "Debit", "Credit", "Remark"]
                headers = ["Account", "Debit", "Credit", "Remark"]
                for col, h in enumerate(headers):
                    worksheet.write(row, col, h, header_fmt)
                row += 1
                for j in summary:
                    worksheet.write(row, 0, j['account'], text_fmt)
                    # worksheet.write(row, 1, j['currency'])
                    worksheet.write(row, 1, j['debit'], money_fmt)
                    worksheet.write(row, 2, j['credit'], money_fmt)
                    worksheet.write(row, 3, j['remark'], text_fmt)
                    row += 1

        # ----------------------------
        # PURCHASE REPORT
        # ----------------------------
        if wizard.document_type == 'purchase' and wizard.document_id:
            po = wizard.document_id
            company_currency = po.company_id.currency_id
            worksheet.write(row, 0, f"Purchase Order: {po.name}", section_fmt)
            po_currency = po.currency_id
            row += 2

            headers = ["Product", "Qty Ordered", "Qty Received", "Qty Invoiced",
                       f"Ordered Value({po_currency.name})",
                       f"Invoiced Value({po_currency.name})",
                       f"Stock Value({company_currency.name})",
                       f"Unit Cost({company_currency.name})"
                       # "Margin %",
                       # "Variance"
                    ]
            for col, h in enumerate(headers):
                worksheet.write(row, col, h, header_fmt)
                worksheet.set_column(col, col, 14)
            row += 1

            lines = self._compute_purchase_lines(po)
            for line in lines:
                worksheet.write(row, 0, line['product'],text_fmt)
                worksheet.write(row, 1, line.get('qty_ordered', 0), text_fmt)
                worksheet.write(row, 2, line.get('qty_received', 0), text_fmt)
                worksheet.write(row, 3, line.get('qty_invoiced', 0), text_fmt)
                worksheet.write(row, 4, line.get('value_ordered', 0), money_fmt)
                worksheet.write(row, 5, line.get('value_invoiced', 0), money_fmt)
                worksheet.write(row, 6, line.get('value_stock', 0), money_fmt)
                worksheet.write(row, 7, line.get('unit_cost', 0), money_fmt)
                # worksheet.write(row, 10, line.get('margin', 0))
                # worksheet.write(row, 11, line.get('variance', 0), money_fmt)
                row += 1

            # 1) Billing entries
            write_summary("Billing Entries Summary", sum_journal_lines(po.invoice_ids, reference="".join(po.invoice_ids.mapped("name"))))
            # 2) Inventory entries
            inventory_moves = po.order_line.mapped("move_ids.account_move_ids")
            write_summary("Inventory Entries Summary", sum_journal_lines(inventory_moves, reference="; ".join(po.order_line.mapped('move_ids.picking_id.name'))))

            # 3) Landed costs entries
            landed_moves = request.env['account.move']
            for picking in po.picking_ids:
                landed_costs = request.env['stock.landed.cost'].search([('picking_ids', 'in', [picking.id])])
                landed_moves |= landed_costs.mapped("account_move_id")

            write_summary("Landed Cost Entries Summary",
                          sum_journal_lines(landed_moves, reference="; ".join(landed_moves.mapped("ref"))))
        # ----------------------------
        # RECEIPT REPORT
        # ----------------------------
        if wizard.document_type == 'picking' and wizard.picking_id:
            picking = wizard.picking_id
            worksheet.write(row, 0, f"Receipt / Delivery: {picking.name}", section_fmt)
            row += 2

            headers = ["Product", "Qty Moved", "Stock Value", "Unit Cost"]
            for col, h in enumerate(headers):
                worksheet.write(row, col, h, header_fmt)
            row += 1

            lines = self._compute_picking_lines(picking)
            for line in lines:
                worksheet.write(row, 0, line['product'])
                worksheet.write(row, 1, line.get('qty_received', 0))
                worksheet.write(row, 2, line.get('value_stock', 0), money_fmt)
                worksheet.write(row, 3, line.get('unit_cost', 0), money_fmt)
                row += 1

            # 1) Receipt journal summary → use stock receipt reference
            write_summary("Receipt Entries Summary",
                          sum_journal_lines(picking.account_move_ids, reference=picking.name))

            # 2) Landed cost entries with cross-pending references
            landed_moves = request.env['account.move']
            landed_costs = request.env['stock.landed.cost'].search([('picking_ids', 'in', [picking.id])])
            landed_moves |= landed_costs.mapped("account_move_id")

            write_summary("Landed Cost Entries Summary",
                          sum_journal_lines(landed_moves, reference="; ".join(landed_moves.mapped("ref"))))

        workbook.close()
        output.seek(0)
        filename = f"Purchase_Reconcile_{wizard.document_id.mapped('name')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        return request.make_response(
            output.read(),
            headers=[
                ('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
                ('Content-Disposition', f'attachment; filename={filename}')
            ]
        )

    # -----------------------------
    # Helper computation methods
    # -----------------------------
    def _compute_purchase_lines(self, po):
        res = []
        for line in po.order_line:
            product = line.product_id
            qty_ordered = line.product_qty

            invoice_lines = line.invoice_lines.filtered(lambda l: l.move_id.state == 'posted')
            qty_invoiced = sum(invoice_lines.mapped('quantity'))
            value_invoiced = sum(invoice_lines.mapped('price_total'))

            stock_moves = line.move_ids.filtered(lambda m: m.state == 'done')
            qty_received = sum(stock_moves.mapped('product_qty'))
            value_stock = sum(sum(l.value for l in move.stock_valuation_layer_ids) for move in stock_moves)

            # Accounts
            account_lines = invoice_lines.mapped('move_id.line_ids').filtered(lambda l: l.move_id.state == 'posted')
            debit_by_account = {}
            credit_by_account = {}
            for l in account_lines:
                debit_by_account[l.account_id.name] = debit_by_account.get(l.account_id.name,0.0) + l.debit
                credit_by_account[l.account_id.name] = credit_by_account.get(l.account_id.name,0.0) + l.credit

            unit_cost = value_stock / qty_received if qty_received else 0.0
            variance = value_invoiced - value_stock # Difference

            res.append({
                'product': product.display_name,
                'qty_ordered': qty_ordered,
                'qty_received': qty_received,
                'qty_invoiced': qty_invoiced,
                'value_ordered': line.price_subtotal,
                'value_stock': value_stock,
                'value_invoiced': value_invoiced,
                'unit_cost': unit_cost,
                'debit_account': ', '.join(f"{k}: {v}" for k,v in debit_by_account.items()),
                'credit_account': ', '.join(f"{k}: {v}" for k,v in credit_by_account.items()),
                'margin': 0,
                'variance': variance,
            })
        return res


