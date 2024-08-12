# -*- coding: utf-8 -*-
from odoo import models, fields
import xlwt
import base64
from io import BytesIO


class PurchaseOrderWizard(models.TransientModel):
    _name = 'purchase.order.wizard'

    file_name = fields.Binary(string="Report")
    data_fname = fields.Char(string="File Name")
    po_wizard = fields.Many2many('purchase.order', string="Purchase Wizard")

    def button_purchase_order_wizard(self):
        # Create a new workbook
        workbook = xlwt.Workbook(encoding='utf-8')

        # Keep track of sheet names to ensure uniqueness
        used_sheet_names = set()

        for po in self.po_wizard:
            # Generate a unique sheet name
            base_name = po.partner_id.name[:31]  # Excel sheet name max length is 31 characters
            sheet_name = base_name
            suffix = 1
            while sheet_name in used_sheet_names:
                sheet_name = f"{base_name}_{suffix}"
                suffix += 1

            # Add the new sheet with the unique name
            used_sheet_names.add(sheet_name)
            sheet = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)

            # Define styles
            format1 = xlwt.easyxf(
                'align:horiz center; font:color black,bold True; border:top_color black,bottom_color black')
            format2 = xlwt.easyxf(
                'font:bold True; pattern: pattern solid, fore_colour gray25; align: horiz left; borders: top_color black, bottom_color black, right_color black, left_color black, left thin, right thin, top thin, bottom thin;')
            format4 = xlwt.easyxf(
                'align:horiz center; font:color black,bold True, height 250; border:top_color black,bottom_color black; borders: top_color black, bottom_color black, right_color black, left_color black, left thin, right thin, top thin, bottom thin;')

            # Define column widths
            column_width = 7000
            for col in range(9):
                sheet.col(col).width = column_width
            sheet.row(0).height = 200

            # Merge cells for the Purchase Order header
            sheet.write_merge(0, 1, 0, 8, "Purchase Order", format4)

            # Write Purchase Order Summary
            sheet.write(2, 0, "Order Reference:", format2)
            sheet.write(2, 1, "Vendor Name:", format2)
            sheet.write(2, 2, "Order Date:", format2)
            sheet.write(2, 3, "Total Amount:", format2)
            sheet.write(3, 0, po.name or '', format1)
            sheet.write(3, 1, po.partner_id.name or '', format1)
            sheet.write(3, 2, po.date_order.strftime('%Y-%m-%d') if po.date_order else '', format1)
            sheet.write(3, 3, f"{po.amount_total:.2f}" if po.amount_total else '', format1)

            # Add a blank row before starting the Purchase Order Line details
            row = 8

            # Write header row for Purchase Order Lines
            headers = ["Vendor Name", "Product ID", "Product Name", "Order Reference", "Quantity", "Unit Price",
                       "Subtotal", "Date Ordered", "Delivery Date"]
            for col, header in enumerate(headers):
                sheet.write(7, col, header, format2)

            # Write order details
            for line in po.order_line:
                sheet.write(row, 0, po.partner_id.name or '', format1)
                sheet.write(row, 1, line.product_id.default_code or '', format1)
                sheet.write(row, 2, line.product_id.name or '', format1)
                sheet.write(row, 3, po.name or '', format1)
                sheet.write(row, 4, f"{line.product_qty:.2f}" if line.product_qty else '', format1)
                sheet.write(row, 5, f"{line.price_unit:.2f}" if line.price_unit else '', format1)
                sheet.write(row, 6, f"{line.price_subtotal:.2f}" if line.price_subtotal else '', format1)
                sheet.write(row, 7, po.date_order.strftime('%Y-%m-%d') if po.date_order else '', format1)
                sheet.write(row, 8, line.date_planned.strftime('%Y-%m-%d') if line.date_planned else '', format1)
                row += 1

            # Merge cells for the Purchase Order Lines header
            sheet.write_merge(5, 6, 0, 8, 'Purchase Order Lines', format4)

        # Save to BytesIO stream and encode
        stream = BytesIO()
        workbook.save(stream)
        file_data = stream.getvalue()
        out = base64.b64encode(file_data)

        # Assign encoded data to fields
        self.file_name = out
        self.data_fname = 'Purchase_Report.xls'  # Ensure the filename ends with .xls

        return {
            "res_id": self.id,
            "name": 'Purchase Order Report',
            "view_type": 'form',
            "view_mode": 'form',
            "res_model": 'purchase.order.wizard',
            "view_id": False,
            "type": 'ir.actions.act_window'
        }
