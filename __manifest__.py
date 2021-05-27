# -*- coding: utf-8 -*-
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl.html).
{
    "name": "Extensión de Reportes Contables",
    "version": "11.0.1.0.0",
    "category": "Report",
    "website": "http://argil.mx",
    "author": "<Argil>, German Ponce Dominguez (Desarrollador)",
    "license": "AGPL-3",
    "application": False,
    "installable": True,
    "description": """

Modificaciones de Reportes Contables.

    
    """,
    "external_dependencies": {
        "python": ['xlsxwriter'],
        "bin": [],
    },
    "depends": [
        "account",
        "financial_reports",
        "account_financial_report",
    ],
    "data": [
        # 'reports/report_account.xml',
        'account.xml',
    ],
    "demo": [
    ],
    "qweb": [
    ]
}
