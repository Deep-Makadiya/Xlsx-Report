
# -*- coding: utf-8 -*-
{
    'name': "xlsx Report",

    'summary': "Printing the xlsx report in purchase.order",

    'description': """
IN this model i have maded a xlsx report in the purchase.order modeule.

    """,

    'author': "My Company",
    'website': "https://www.yourcompany.com",

    # Categories can be used to filter modules in modules listing
    # Check https://github.com/odoo/odoo/blob/15.0/odoo/addons/base/data/ir_module_category_data.xml
    # for the full list
    'category': 'Uncategorized',
    'version': '0.1',

    # any module necessary for this one to work correctly
    'depends': ['base', 'sale_management', 'account', 'sale', 'purchase', 'stock', 'product'],

    # always loaded
    'data': [

        'security/ir.model.access.csv',
        'views/purchase_order_views.xml',
        'wizard/purchase_order_wizard_views.xml',


    ],
    'installable': True,
    'application': True,
}
