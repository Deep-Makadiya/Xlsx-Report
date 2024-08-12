# -*- coding: utf-8 -*-
from odoo import models, fields
import xlwt
import base64
from io import BytesIO


class PurchaseOrder(models.Model):
    _inherit = 'purchase.order'

    file_name = fields.Binary(string="Report")
    data_fname = fields.Char(string="File Name")

    def print_excel(self):
        filename = f"{self.name}.xls"
        workbook = xlwt.Workbook(encoding='utf-8')
        sheet1 = workbook.add_sheet(self.partner_id.name, cell_overwrite_ok=True)

        format1 = xlwt.easyxf(
            'align:horiz center; font:color black,bold True; border:top_color black,bottom_color black')
        format2 = xlwt.easyxf(
            'font:bold True;pattern: pattern solid, fore_colour gray25;align: horiz left; borders: top_color black, bottom_color black, right_color black, left_color black, left thin, right thin, top thin, bottom thin;')
        format3 = xlwt.easyxf('font:bold True;align: horiz left')
        format4 = xlwt.easyxf(
            'align:horiz center; font:color black,bold True ,height 250; border:top_color black,bottom_color black; borders: top_color black, bottom_color black, right_color black, left_color black, left thin, right thin, top thin, bottom thin;')

        # Define column widths
        column_width = 7000
        for col in range(9):
            sheet1.col(col).width = column_width
        sheet1.row(0).height = 200

        # Merge cells for the Purchase Order header
        sheet1.write_merge(0, 1, 0, 8, "Purchase Order", format4)

        # Write Purchase Order Summary
        sheet1.write(2, 0, "Order Reference:", format2)
        sheet1.write(2, 1, "Vendor Name:", format2)
        sheet1.write(2, 2, "Order Date:", format2)
        sheet1.write(2, 3, "Total Amount:", format2)
        sheet1.write(3, 0, self.name or '', format1)
        sheet1.write(3, 1, self.partner_id.name or '', format1)
        sheet1.write(3, 2, self.date_order.strftime('%Y-%m-%d') if self.date_order else '', format1)
        sheet1.write(3, 3, f"{self.amount_total:.2f}" if self.amount_total else '', format1)

        # Add a blank row before starting the Purchase Order Line details
        row = 8

        # Write header row for Purchase Order Lines
        headers = ["Vendor Name", "Product ID", "Product Name", "Order Reference", "Quantity", "Unit Price", "Subtotal",
                   "Date Ordered", "Delivery Date"]
        for col, header in enumerate(headers):
            sheet1.write(7, col, header, format2)

        # Write order details
        for line in self.order_line:
            sheet1.write(row, 0, self.partner_id.name or '', format1)
            sheet1.write(row, 1, line.product_id.default_code or '', format1)
            sheet1.write(row, 2, line.product_id.name or '', format1)
            sheet1.write(row, 3, self.name or '', format1)
            sheet1.write(row, 4, f"{line.product_qty:.2f}" if line.product_qty else '', format1)
            sheet1.write(row, 5, f"{line.price_unit:.2f}" if line.price_unit else '', format1)
            sheet1.write(row, 6, f"{line.price_subtotal:.2f}" if line.price_subtotal else '', format1)
            sheet1.write(row, 7, self.date_order.strftime('%Y-%m-%d') if self.date_order else '', format1)
            sheet1.write(row, 8, line.date_planned.strftime('%Y-%m-%d') if line.date_planned else '', format1)
            row += 1

        # Merge cells for the Purchase Order Lines header
        sheet1.write_merge(5, 6, 0, 8, 'Purchase Order Lines', format4)

        # Save to BytesIO stream and encode
        stream = BytesIO()
        workbook.save(stream)
        out = base64.b64encode(stream.getvalue())
        self.file_name = out
        self.data_fname = filename

        return {
            "res_id": self.id,
            "name": 'Purchase Order Report',
            "view_type": 'form',
            "view_mode": 'form',
            "res_model": 'purchase.order',
            "view_id": False,
            "type": 'ir.actions.act_window'
        }
