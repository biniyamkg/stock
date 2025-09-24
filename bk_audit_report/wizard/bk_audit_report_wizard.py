from odoo import models, fields, api, _
from collections import defaultdict


class BkAuditReportWizard(models.TransientModel):
    _name = "bk.audit.report_wizard"
    _description = "Audit Report Wizard"

    document_type = fields.Selection([
        ("purchase", "Purchase Order"),
    ], string="Document Type", default="purchase")

    document_id = fields.Many2one("purchase.order", string="Purchase Order",
                                  domain="[('state','not in', ('draft','sent','cancel'))]")
    # sale_order_id = fields.Many2one("sale.order", string="Sales Order")
    # picking_id = fields.Many2one("stock.picking", string="Incoming Receipt")
    # date_start = fields.Date(string="Start Date")
    # date_end = fields.Date(string="End Date")

    def action_generate_excel(self):
        return {
            "type": "ir.actions.act_url",
            "url": f"/web/export_excel?wizard_id={self.id}",
            "target": "self",
        }

