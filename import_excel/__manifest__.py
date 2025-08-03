{
    'name': 'Import Excel Files',
    'version': '1.0',
    'author': 'Your Name',
    'category': 'Tools',
    'summary': 'Import Excel files and create products in Odoo.',
    'depends': ['base', 'product', 'account'],
    'data': [
        'security/ir.model.access.csv',
        'views/import_wizard_view.xml',
        'views/product_template_view.xml',
    ],
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
}
