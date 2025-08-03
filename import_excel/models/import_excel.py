from odoo import models, fields


class ExcelData(models.Model):
    _name = 'excel.data'
    _description = 'Excel Data'


class ProductTemplate(models.Model):
    _inherit = 'product.template'

    series = fields.Char(string="Series")
    min_price = fields.Float(string="Minimum Price")
    dimensions = fields.Char(string="Dimensions")
    finishing = fields.Char(string="Finishing")
    brand = fields.Char(string="Brand")
    op_stock = fields.Float(string="Opening Stock")
    op_amt = fields.Float(string="Opening Amount")
    retail_price = fields.Float(string="Retail Price")
