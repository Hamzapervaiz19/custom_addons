# import base64
# import xlrd
# from odoo import models, fields
# from odoo.exceptions import UserError
#
# class ImportWizard(models.TransientModel):
#     _name = 'import.wizard'
#     _description = 'Import Excel Wizard'
#
#     file = fields.Binary(string='Upload File', required=True)
#     filename = fields.Char()
#
#     def import_file(self):
#         if not self.file:
#             raise UserError("Please upload a file.")
#
#         file_data = base64.b64decode(self.file)
#
#         try:
#             workbook = xlrd.open_workbook(file_contents=file_data)
#             sheet = workbook.sheet_by_index(0)
#         except Exception as e:
#             raise UserError(f"Error reading Excel file: {e}")
#
#         create_vals_batch = []
#         BATCH_SIZE = 500
#
#         for row in range(1, sheet.nrows):  # Assuming first row is header
#             item_name = str(sheet.cell(row, 0).value).strip()
#             default_code = str(sheet.cell(row, 1).value).strip()
#             series = str(sheet.cell(row, 2).value).strip()
#             finishing = str(sheet.cell(row, 4).value).strip()
#             brand = str(sheet.cell(row, 5).value).strip()
#             usage_unit = str(sheet.cell(row, 6).value).strip()
#
#             # Tax cleaning
#             tax_raw = str(sheet.cell(row, 9).value).replace("%", "").strip()
#             try:
#                 tax_value = float(tax_raw)
#             except:
#                 tax_value = 0.0
#
#             sales_price_raw = str(sheet.cell(row, 10).value).replace("OMR", "").strip()
#             try:
#                 list_price = float(sales_price_raw)
#             except:
#                 list_price = 0.0
#
#             min_price_raw = sheet.cell(row, 11).value
#             try:
#                 min_price = float(min_price_raw)
#             except:
#                 min_price = 0.0
#
#             purchase_price_raw = str(sheet.cell(row, 12).value).replace("OMR", "").strip()
#             try:
#                 standard_price = float(purchase_price_raw)
#             except:
#                 standard_price = 0.0
#
#             description = str(sheet.cell(row, 14).value).strip()
#
#             # UOM handling
#             uom = self.env['uom.uom'].search([('name', '=', usage_unit)], limit=1)
#             uom_id = uom.id if uom else 1
#
#             # Tax handling
#             if tax_value:
#                 tax = self.env['account.tax'].search([('amount', '=', tax_value)], limit=1)
#                 if not tax:
#                     tax = self.env['account.tax'].create({
#                         'name': f"{tax_value}%",
#                         'amount': tax_value,
#                         'type_tax_use': 'sale'
#                     })
#                 taxes_id = [(6, 0, [tax.id])]
#             else:
#                 taxes_id = [(6, 0, [])]
#
#             # Tags: brand + finishing
#             tag_names = [brand, finishing]
#             tag_ids = []
#             for tag_name in tag_names:
#                 if tag_name:
#                     tag = self.env['product.tag'].search([('name', '=', tag_name)], limit=1)
#                     if not tag:
#                         tag = self.env['product.tag'].create({'name': tag_name})
#                     tag_ids.append(tag.id)
#
#             product_tag_ids = [(6, 0, tag_ids)] if tag_ids else [(6, 0, [])]
#
#             vals = {
#                 "default_code": default_code,
#                 "name": item_name,
#                 "series": series,  # custom field
#                 "standard_price": standard_price,
#                 "list_price": list_price,
#                 "uom_id": uom_id,
#                 "uom_po_id": uom_id,
#                 "taxes_id": taxes_id,
#                 "product_tag_ids": product_tag_ids,
#                 "description_sale": description,
#                 "min_price": min_price,  # custom field
#             }
#
#             # Check if product exists → update, else batch create
#             product = self.env['product.template'].search([('default_code', '=', default_code)], limit=1)
#             if product:
#                 product.write(vals)
#             else:
#                 create_vals_batch.append(vals)
#
#             # Batch create every 500 products
#             if len(create_vals_batch) >= BATCH_SIZE:
#                 self.env['product.template'].create(create_vals_batch)
#                 create_vals_batch = []
#
#         # Create remaining products
#         if create_vals_batch:
#             self.env['product.template'].create(create_vals_batch)

import base64
import xlrd
from odoo import models, fields
from odoo.exceptions import UserError


class ImportWizard(models.TransientModel):
    _name = 'import.wizard'
    _description = 'Import Excel Wizard'

    file = fields.Binary(string='Upload File', required=True)
    filename = fields.Char()

    def import_file(self):
        if not self.file:
            raise UserError("Please upload a file.")

        file_data = base64.b64decode(self.file)

        try:
            workbook = xlrd.open_workbook(file_contents=file_data)
            sheet = workbook.sheet_by_index(0)
        except Exception as e:
            raise UserError(f"Error reading Excel file: {e}")

        create_vals_batch = []
        BATCH_SIZE = 500

        for row in range(1, sheet.nrows):  # Assuming first row is header
            item_name = str(sheet.cell(row, 0).value).strip()
            default_code = str(sheet.cell(row, 1).value).strip()
            series = str(sheet.cell(row, 2).value).strip()
            dimensions = str(sheet.cell(row, 3).value).strip()
            finishing = str(sheet.cell(row, 4).value).strip()
            brand = str(sheet.cell(row, 5).value).strip()
            usage_unit = str(sheet.cell(row, 6).value).strip()

            # OP Stock
            try:
                op_stock = float(sheet.cell(row, 7).value)
            except:
                op_stock = 0.0

            # OP Amt
            try:
                op_amt = float(sheet.cell(row, 8).value)
            except:
                op_amt = 0.0

            # Tax cleaning
            tax_raw = str(sheet.cell(row, 9).value).replace("%", "").strip()
            try:
                tax_value = float(tax_raw)
            except:
                tax_value = 0.0

            # Sales price
            sales_price_raw = str(sheet.cell(row, 10).value).replace("OMR", "").strip()
            try:
                list_price = float(sales_price_raw)
            except:
                list_price = 0.0

            # Min price
            try:
                min_price = float(sheet.cell(row, 11).value)
            except:
                min_price = 0.0

            # Purchase price
            purchase_price_raw = str(sheet.cell(row, 12).value).replace("OMR", "").strip()
            try:
                standard_price = float(purchase_price_raw)
            except:
                standard_price = 0.0

            # Retail price
            try:
                retail_price = float(sheet.cell(row, 13).value)
            except:
                retail_price = 0.0

            description = str(sheet.cell(row, 14).value).strip()

            # UOM handling
            uom = self.env['uom.uom'].search([('name', '=', usage_unit)], limit=1)
            uom_id = uom.id if uom else 1

            # Tax handling
            if tax_value:
                tax = self.env['account.tax'].search([('amount', '=', tax_value)], limit=1)
                if not tax:
                    tax = self.env['account.tax'].create({
                        'name': f"{tax_value}%",
                        'amount': tax_value,
                        'type_tax_use': 'sale'
                    })
                taxes_id = [(6, 0, [tax.id])]
            else:
                taxes_id = [(6, 0, [])]

            vals = {
                "default_code": default_code,
                "name": item_name,
                "series": series,  # custom field
                "dimensions": dimensions,  # custom field
                "finishing": finishing,  # custom field
                "brand": brand,  # custom field
                "op_stock": op_stock,  # custom field
                "op_amt": op_amt,  # custom field
                "retail_price": retail_price,  # custom field
                "standard_price": standard_price,
                "list_price": list_price,
                "uom_id": uom_id,
                "uom_po_id": uom_id,
                "taxes_id": taxes_id,
                "description_sale": description,
                "min_price": min_price,  # custom field
            }

            product = self.env['product.template'].search([('default_code', '=', default_code)], limit=1)
            if product:
                product.write(vals)
            else:
                create_vals_batch.append(vals)

            # Batch create every 500 products
            if len(create_vals_batch) >= BATCH_SIZE:
                self.env['product.template'].create(create_vals_batch)
                create_vals_batch = []

        # Create remaining products
        if create_vals_batch:
            self.env['product.template'].create(create_vals_batch)

