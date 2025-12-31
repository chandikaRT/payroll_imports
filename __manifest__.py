{
    'name': 'Payroll Imports',
    'version': '1.0',
    'summary': 'Import various payroll components into payslips',
    'description': 'Import any payslip input type (transport, incentives, deductions, etc.) from Excel',
    'author': 'Chandika Rathnayake',
    'depends': ['hr_payroll'],
    'data': [
        'security/ir.model.access.csv',
        'data/payroll_input_types.xml',
        'views/payroll_import_views.xml',
        'views/import_wizard_views.xml',
    ],
    'installable': True,
    'application': True,
}