# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.
{
    'name': 'Debt Management System',
    'version': '19.1.0',
    'category': 'Finance',
    'summary': 'This module is used to manage Debt.',
    'author': "Namah Softech Private Limited",
    'maintainer': 'Namah Softech Pvt Ltd',
    'description': """
    This Debt Management System helps streamline the management of debts and loans for individuals and
        organizations. It provides a centralized platform for tracking outstanding debts, monitoring payment schedules,
        and automating reminders for upcoming payments. 
    """,
    'depends': ['mail', 'base'],
    'sequence': 0,
    'data': [
        'security/ir.model.access.csv',
        'data/email_template.xml',
        'data/schedule_action.xml',
        'views/debt_details_views.xml',
        'views/emi_payment_views.xml',
        'views/debt_details_menuitem.xml',
    ],
    'images': ['static/description/banner.png'],
    'installable': True,
    'application': True,
    'auto_install': False,
    'license': 'OPL-1',
    'post_init_hook': 'import_bank_names',
}
