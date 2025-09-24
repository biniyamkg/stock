{
    "name": "Document Based Audit Report",
    "summary": "Audit verification per Purchase / Sale / Picking — HTML & Excel",
    "description": """
    """,
    "version": "1.0.0",
    "author": "Your Company / Your Name",

    "category": "Accounting/Reporting",
    "license": "LGPL-3",
    "depends": ["base", "purchase", "stock_landed_costs",
    "sale_management", "stock", "account",
     "stock_account", "web"],
    "data": [
        "security/ir.model.access.csv",
        "wizard/bk_audit_report_wizard_view.xml",
        "report/audit_report_templates.xml",
    ],
    "assets": {},
    "application": False,
    "installable": True,
    "auto_install": False,
}
