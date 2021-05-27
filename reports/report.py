# -*- coding: utf-8 -*-
##############################################################################
#
#    Cybrosys Technologies Pvt. Ltd.
#    Copyright (C) 2017-TODAY Cybrosys Technologies(<http://www.cybrosys.com>).
#    Author: Jesni Banu(<https://www.cybrosys.com>)
#    you can modify it under the terms of the GNU LESSER
#    GENERAL PUBLIC LICENSE (LGPL v3), Version 3.
#
#    It is forbidden to publish, distribute, sublicense, or sell copies
#    of the Software or modified copies of the Software.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU LESSER GENERAL PUBLIC LICENSE (LGPL v3) for more details.
#
#    You should have received a copy of the GNU LESSER GENERAL PUBLIC LICENSE
#    GENERAL PUBLIC LICENSE (LGPL v3) along with this program.
#    If not, see <http://www.gnu.org/licenses/>.
#
##############################################################################
import datetime
from odoo.addons.report_xlsx.report.report_xlsx import ReportXlsxAbstract
from odoo import models
from openerp.modules import module

class PartnerXlsx(models.AbstractModel):
    _name = 'report.export_mro_pickings.stock_report_mro_xls'
    _inherit = 'report.report_xlsx.abstract'

    def generate_xlsx_report(self, workbook, data, lines):
        ## texto continuo excel ##
        # 'text_wrap': True

        ### browse iterator ###
        rec = lines[0]

        sheet = workbook.add_worksheet('Reporte Entrega de Producto')
        ## Expandiendo las Columnas ##
        sheet.set_column('A:A', 8)
        sheet.set_column('B:B', 30)
        sheet.set_column('C:C', 20)
        sheet.set_column('D:D', 30)
        sheet.set_column('E:I', 20)
        sheet.set_column('J:K', 30)
        # sheet.set_column('D:J', 14)

        ## Format para moneda
        num_format = '$ #,##0.00'
        bg_gray = '#D8D8D8'

        module_path = module.get_module_path('export_mro_pickings')
        #print "########## MODULE PATH >>> ",module_path
        image_module_path = module_path+'/images/logo.png'
        sheet.insert_image('A2', image_module_path)

        format_period_title = workbook.add_format({
                                'bold':     True,
                                'align':    'center',
                                'valign':   'vcenter',
                            })

        format_period_title.set_font_size(18)
        format_period_title.set_bottom(2)
        format_period_title.set_top(2)

        format_bold_border = workbook.add_format({'bold': True, 'valign':   'vcenter'})
        format_bold_border.set_border(2)

        format_bold_border_tit_gray = workbook.add_format({'bold': True, 'valign':   'vcenter'})
        format_bold_border_tit_gray.set_border(2)
        format_bold_border_tit_gray.set_bg_color(bg_gray)

        format_bold_border2 = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center'})
        format_bold_border2.set_border(2)

        format_bold_border_bg_yllw = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center'})
        format_bold_border_bg_yllw.set_border(2)
        format_bold_border_bg_yllw.set_bg_color("#F0FF5B")

        format_bold_border_bg_gray = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center'})
        format_bold_border_bg_gray.set_border(2)
        format_bold_border_bg_gray.set_bg_color(bg_gray)

        format_header_border_bg_gray = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center', 'text_wrap': True})
        format_header_border_bg_gray.set_border(2)
        format_header_border_bg_gray.set_bg_color(bg_gray)
        format_header_border_bg_gray.set_font_size(12)

        format_bold_border_bg_yllw_line = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center'})
        format_bold_border_bg_yllw_line.set_border(1)
        format_bold_border_bg_yllw_line.set_bg_color("#F0FF5B")
        format_bold_border_bg_yllw_line.set_font_size(9)


        format_bold_border_bg_wht_line = workbook.add_format({'valign':   'vcenter', 'align':   'center'})
        format_bold_border_bg_wht_line.set_border(1)
        format_bold_border_bg_wht_line.set_font_size(9)


        format_bold_border_bg_wht_line_boxes = workbook.add_format({'valign':   'vcenter', 'align':   'center'})
        format_bold_border_bg_wht_line_boxes.set_border(1)
        format_bold_border_bg_wht_line_boxes.set_bg_color(bg_gray)
        format_bold_border_bg_wht_line_boxes.set_font_size(9)

        format_bold_border_bg_wht_line_left = workbook.add_format({'valign':   'vcenter', 'align':   'left'})
        format_bold_border_bg_wht_line_left.set_border(1)
        format_bold_border_bg_wht_line_left.set_font_size(9)

        format_bold_border_bg_wht_signs = workbook.add_format({'align':   'center', 'bold': True})
        format_bold_border_bg_wht_signs.set_border(1)
        format_bold_border_bg_wht_signs.set_font_size(9)


        sheet.write('I3','CÓDIGO',format_bold_border_tit_gray)
        sheet.write('I4','REV',format_bold_border_tit_gray)

        sheet.write('I6','FECHA:',format_bold_border_tit_gray)
        sheet.write('I7','FOLIO:',format_bold_border_tit_gray)
        sheet.write('I8','CLIENTE:',format_bold_border_tit_gray)

        sheet.merge_range('J3:K3','RC02-PSC-8.5',format_bold_border2)
        sheet.merge_range('J4:K4','2',format_bold_border2)
        # sheet.merge_range('J4:k4',format_bold_border2)

        fecha_creacion = rec.create_date_report
        fecha_creacion_sp = fecha_creacion.split('-')
        fecha_creacion = fecha_creacion_sp[2]+'/'+fecha_creacion_sp[1]+'/'+fecha_creacion_sp[0]
        sheet.merge_range('J6:K6',str(fecha_creacion),format_bold_border_bg_gray)
        sheet.merge_range('J7:K7',rec.sequence_name,format_bold_border_bg_yllw)
        sheet.merge_range('J8:K8',rec.partner_id.name if rec.partner_id else "N/A",format_bold_border2)

        sheet.merge_range('C4:G6', 'REPORTE DE ENTREGA DE PRODUCTO',format_period_title)

        sheet.merge_range('A11:A12', '#',format_header_border_bg_gray)
        sheet.merge_range('B11:B12', 'N/P',format_header_border_bg_gray)
        sheet.merge_range('C11:C12', 'REV.',format_header_border_bg_gray)
        sheet.merge_range('D11:D12', 'LÍNEA DE PRODUCTO',format_header_border_bg_gray)
        sheet.merge_range('E11:E12', 'CANTIDAD TOTAL',format_header_border_bg_gray)
        sheet.merge_range('F11:F12', 'UM',format_header_border_bg_gray)
        sheet.merge_range('G11:G12', 'CANTIDAD POR EMPAQUE',format_header_border_bg_gray)
        sheet.merge_range('H11:H12', 'CANTIDAD CAJAS',format_header_border_bg_gray)
        sheet.merge_range('I11:I12', 'CAJAS PZS INCOMPLETAS',format_header_border_bg_gray)
        sheet.merge_range('J11:J12', 'PO / FACTURA',format_header_border_bg_gray)
        sheet.merge_range('K11:K12', 'MO',format_header_border_bg_gray)

        posicion_fila = 13
        ## format_bold_border_bg_yllw_line
        i=1
        for line in rec.line_ids:
            sheet.write('A'+str(posicion_fila), str(i), format_bold_border_bg_yllw_line)
            sheet.write('B'+str(posicion_fila), line.n_p if line.n_p else "", format_bold_border_bg_wht_line)
            sheet.write('C'+str(posicion_fila), line.rev if line.rev else "", format_bold_border_bg_wht_line)
            sheet.write('D'+str(posicion_fila), line.partner_id.name if line.partner_id else "", format_bold_border_bg_wht_line)
            sheet.write('E'+str(posicion_fila), str(line.product_qty) if line.product_qty else "", format_bold_border_bg_wht_line)
            sheet.write('F'+str(posicion_fila), str(line.product_uom.name) if line.product_uom else "", format_bold_border_bg_wht_line)
            sheet.write('G'+str(posicion_fila), str(line.product_qty_pack) if line.product_qty_pack else "", format_bold_border_bg_wht_line)
            sheet.write('H'+str(posicion_fila), str(line.product_qty_box) if line.product_qty_box else "", format_bold_border_bg_wht_line_boxes)
            sheet.write('I'+str(posicion_fila), str(line.product_qty_trash) if line.product_qty_trash else "", format_bold_border_bg_wht_line_boxes)
            sheet.write('J'+str(posicion_fila), str(line.invoice_ref) if line.invoice_ref else "", format_bold_border_bg_wht_line)
            sheet.write('K'+str(posicion_fila), str(line.name) if line.name else "", format_bold_border_bg_wht_line)
            posicion_fila += 1
            i += 1
        posicion_fila+=1
        sheet.write('D'+str(posicion_fila), "TOTAL", format_header_border_bg_gray)
        sheet.write('E'+str(posicion_fila), str(rec.product_qty_total) if rec.product_qty_total else "0.0", format_bold_border2)
        sheet.write('H'+str(posicion_fila), str(rec.product_qty_boxes) if rec.product_qty_total else "0.0", format_bold_border2)
        
        posicion_fila+=2


        sheet.merge_range('A%s:K%s' % (posicion_fila,posicion_fila), 'Observaciones',format_bold_border2)
        posicion_fila += 1
        tmp_p = posicion_fila
        posicion_fila+=2
        sheet.merge_range('A%s:K%s' % (tmp_p,posicion_fila), '',format_header_border_bg_gray)
        posicion_fila += 2
        sheet.merge_range('A%s:K%s' % (posicion_fila,posicion_fila), 'Nota: Marcar con X los rubros necesarios y los innecesarios N/A (No Aplica), nomenclaturas O (ocurre) D (domicilio).',format_bold_border_bg_wht_line_left)
        
        posicion_fila += 2
        tmp_p2 = posicion_fila
        posicion_fila += 5
        sheet.merge_range('A%s:D%s' % (tmp_p2,posicion_fila), 'FIRMA DE TRANSPORTISTA',format_bold_border_bg_wht_signs)
        sheet.merge_range('E%s:H%s' % (tmp_p2,posicion_fila), 'SELLO DE ALMACEN',format_bold_border_bg_wht_signs)
        sheet.merge_range('I%s:K%s' % (tmp_p2,posicion_fila), 'FIRMA Y SELLO DE RECIBIDO',format_bold_border_bg_wht_signs)

        # posicion_fila +=1
        # # for detail in lines[0].line_ids:
        # #     sheet.write(posicion_fila, 0, detail.name, format23)
        # #     sheet.write(posicion_fila, 1, detail.default_code, format23)
        # #     sheet.write(posicion_fila, 2, detail.product_id.name, format23)
        # #     sheet.write(posicion_fila, 3, detail.product_qty, format23)
        # #     sheet.write(posicion_fila, 4, detail.price_unit, format23)
        # #     sheet.write(posicion_fila, 5, detail.product_cost, format24)
        # #     sheet.write(posicion_fila, 6, detail.sale_total, format24)
        # #     sheet.write(posicion_fila, 7, detail.cost_total, format24)
        # #     sheet.write(posicion_fila, 8, detail.utility_unit, format24)
        # #     sheet.write(posicion_fila, 9, detail.utility_total, format24)
        # #     posicion_fila+=1

    
            

# StockReportXlsMRO('export_mro_pickings.stock_report_mro_xls.xlsx', 'mro.report.stock.warehouse')
