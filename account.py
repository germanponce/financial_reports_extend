# -*- coding: utf-8 -*-
#############################################################################
## Codigo de Cherman - German Ponce Dominguez - cherman.seingalt@gmail.com ##
#############################################################################

import time
from datetime import datetime, timedelta
from odoo import api, fields, models, _
from . excel_styles import ExcelStyles
from odoo.exceptions import UserError, ValidationError

import xlwt
import xlsxwriter

from io import BytesIO

import io
import base64
import pdb

from odoo.modules import module
import tempfile
import os

import logging
_logger = logging.getLogger(__name__)

def indice_to_column_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def convert_date_to_MX(date_sting):
    fecha_convert = str(date_sting)
    fecha_convert_sp = fecha_convert.split('-')
    fecha_convert = fecha_convert_sp[2]+'/'+fecha_convert_sp[1]+'/'+fecha_convert_sp[0]
    return fecha_convert

class AccountFinancialReport(models.Model):
    _name = 'account.financial.report'
    _inherit ='account.financial.report'

    # general_balance_2_columns  = fields.Boolean('Balance General 2 columnas')

    # estado_resultados_extended  = fields.Boolean('Estado de Resultados con Porcentaje')

    special_output_report = fields.Selection(
                                             [
                                                ('general_balance_2_columns', 'Balance General 2 columnas'),
                                                ('estado_resultados_extended', 'Estado de Resultados con Porcentaje'),

                                             ], 'Tipo de Salida XLSX')

    @api.constrains('special_output_report')
    def _constraint_special_output_report(self):
        for rec in self:
            if rec.special_output_report:
                other_ids = self.search([('special_output_report','=',rec.special_output_report),('id','!=',rec.id)])
                if other_ids:
                    raise ValidationError(_('Solo puede existir un reporte marcado con esta salida a Excel.'))
        return True
    

    # @api.constrains('estado_resultados_extended')
    # def _constraint_estado_resultados_extended(self):
    #     for rec in self:
    #         if self.estado_resultados_extended:
    #             other_ids = self.search([('estado_resultados_extended','=',True),('id','!=',rec.id)])
    #             if other_ids:
    #                 raise ValidationError(_('Este registro solo debe estar habilitado dentro del Estato de Resultados.'))
    #     return True

class AccountingReport(models.TransientModel):
    _inherit = "account.common.report"
    _description = "Accounting Report"


    def b64str_to_tempfile(self, b64_str=None, file_suffix=None, file_prefix=None):
        """
        @param b64_str : Text in Base_64 format for add in the file
        @param file_suffix : Sufix of the file
        @param file_prefix : Name of file in TempFile
        """
        (fileno, fname) = tempfile.mkstemp(file_suffix, file_prefix)
        f = open(fname, 'wb')
        f.write(base64.decodestring(b64_str or str.encode('')))
        f.close()
        os.close(fileno)
        return fname

    def logo_b64_str_to_physical_file(self, b64_str, file_extension='png', prefix='company_logo'):
        ###### Debemos verificar la extensión del logo si es JPG, PNG, JPEG, WEBPG, ETC.... ######
        _logger.info("\n####################### logo_b64_str_to_physical_file >>>>>>>>>>> ")
        _logger.info("\n####################### file_extension %s " % file_extension)
        _logger.info("\n####################### prefix %s " % prefix)
        b64_temporal_route = self.b64str_to_tempfile(base64.encodestring(b''), 
                                                          file_suffix='.%s' % file_extension, 
                                                          file_prefix='odoo__%s__' % prefix)
        _logger.info("\n### b64_temporal_route %s " % b64_temporal_route)
        ### Guardando el Logo  ###
        f = open(b64_temporal_route, 'wb')
        f.write(base64.decodestring(b64_str or str.encode('')))
        f.close()

        file_result = open(b64_temporal_route, 'rb').read()
        
        return file_result, b64_temporal_route


    @api.multi
    def print_excel_report(self):
        print ("###### print_excel_report  >>>>>>>>>>>> ")
        print ("###### self  >>>>>>>>>>>> ", self)
        self.ensure_one()
        print ("###### account_report_id  >>>>>>>>>>>> ", self.account_report_id)
        
        # general_balance_2_columns
        # estado_resultados_extended
        tables, where_clause, where_params = self.env['account.move.line']._query_get()
        print ("######### tables >>>>>> ", tables)
        print ("######### where_clause >>>>>> ", where_clause)
        print ("######### where_params >>>>>> ", where_params)
        ####### Validación del Balance General ##########
        if self.account_report_id and not self.account_report_id.special_output_report:
            return super(AccountingReport, self).print_excel_report()
        if self.account_report_id.special_output_report == 'general_balance_2_columns':
            return self.print_excel_report_balance_general()
        elif self.account_report_id.special_output_report == 'estado_resultados_extended':
            return self.print_excel_report_estado_resultados()
            
    @api.multi
    def print_excel_report_balance_general(self):
        ################ DATA REPORT #################
        data = {}
        data['ids'] = self.env.context.get('active_ids', [])
        data['model'] = self.env.context.get('active_model', 'ir.ui.menu')
        data['form'] = self.read(['date_from', 'date_to', 'journal_ids', 'target_move', 'company_id'])[0]
        used_context = self._build_contexts(data)
        data['form']['used_context'] = dict(used_context, lang=self.env.context.get('lang', 'en_US'))
        # res = self.check_report()
        data['form'].update(self.read(['debit_credit', 'enable_filter', 'label_filter', 'account_report_id', 'date_from_cmp', 'date_to_cmp', 'journal_ids', 'filter_cmp', 'target_move', 'hierarchy_type', 'other_currency'])[0])
#        for field in ['account_report_id']:
#            if isinstance(data['form'][field], tuple):
#                data['form'][field] = data['form'][field][0]
        comparison_context = self._build_comparison_context(data)
        data['form']['comparison_context'] = comparison_context
        print ("###### comparison_context  >>>>>>>>>>>> ", comparison_context)
#        if data['form']['hierarchy_type'] == 'hierarchy':
#            report_lines = self.get_account_lines_hierarchy(data['form'])
#        else:
        report_lines = self.get_account_lines(data['form'])
        print ("###### report_lines  >>>>>>>>>>>> ", report_lines)

        ### Las lineas se identifican por el campo Level #####
        ### Si el Level es 1 indica que debe ir en la siguiente columna ####

        print ("###### data  >>>>>>>>>>>> ", data)
        
        date_to = self.date_to
        date_from = self.date_from

        # print ("\n\n\n\n\nreport_lines",report_lines)

        report_name = data['form']['account_report_id'][1]

        ################ Con XLSXWRITER ##############

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output)

        sheet = workbook.add_worksheet(report_name)

        ################ Logo a Archivo Temporal ##############

        company_logo = self.env.user.company_id.logo
        if company_logo:
            file_result_b64, logo_path_b64 = self.logo_b64_str_to_physical_file(company_logo, 'png', 'company_logo')
            image_module_path = logo_path_b64
        else:
            module_path = module.get_module_path('financial_reports_extend')
            image_module_path = module_path+'/static/img/logo.jpg'

        ############# ESTILOS ###########

        num_format = '$ #,##0.00'
        bg_gray = '#D8D8D8'

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

        f_wh_detail_save = workbook.add_format({'bold': True, 'valign':   'vcenter',  'align':   'center'})

        f_gray_detail_save = workbook.add_format({'bold': True, 'valign':   'vcenter',})
        f_gray_detail_save.set_border(2)
        f_gray_detail_save.set_bg_color(bg_gray)


        f_gray_detail_save_center = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center'})
        f_gray_detail_save_center.set_border(2)
        f_gray_detail_save_center.set_bg_color(bg_gray)

        f_blue_detail_save_center = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center'})
        f_blue_detail_save_center.set_border(2)
        f_blue_detail_save_center.set_font_color('white')
        f_blue_detail_save_center.set_bg_color("#3465a4")

        format_bold_border2 = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center'})
        format_bold_border2.set_border(2)

        format_bold_border_bg_yllw = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center'})
        format_bold_border_bg_yllw.set_border(2)
        format_bold_border_bg_yllw.set_bg_color("#F0FF5B")

        format_bold_border_bg_gray = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center'})
        format_bold_border_bg_gray.set_border(2)
        format_bold_border_bg_gray.set_bg_color(bg_gray)

        format_header_border_bg_gray = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center', 'text_wrap': True})
        format_header_border_bg_gray.set_border(4)
        format_header_border_bg_gray.set_bg_color(bg_gray)
        format_header_border_bg_gray.set_font_size(12)

        ############ Estilos para el Detalle ##########

        format_header_border_bg = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center', 'text_wrap': True})
        format_header_border_bg.set_border(4)
        format_header_border_bg.set_font_size(12)

        format_header_border_bg_left = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'left', 'text_wrap': True})
        format_header_border_bg_left.set_border(4)
        format_header_border_bg_left.set_font_size(12)

        format_header_border_bg_right = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'right', 'text_wrap': True, 'num_format': '$ #,##0.00'})
        format_header_border_bg_right.set_border(4)
        format_header_border_bg_right.set_font_size(12)

        format_header_border_bg_left_yll = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'left', 'text_wrap': True, 'num_format': '$ #,##0.00'})
        format_header_border_bg_left_yll.set_border(4)
        format_header_border_bg_left_yll.set_bg_color("#f7f4be")
        format_header_border_bg_left_yll.set_font_size(12)

        format_header_border_bg_right_yll = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'right', 'text_wrap': True, 'num_format': '$ #,##0.00'})
        format_header_border_bg_right_yll.set_border(4)
        format_header_border_bg_right_yll.set_bg_color("#f7f4be")
        format_header_border_bg_right_yll.set_font_size(12)

        format_header_border_bg_right_gray = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'right', 'text_wrap': True})
        format_header_border_bg_right_gray.set_border(4)
        format_header_border_bg_right_gray.set_bg_color(bg_gray)
        format_header_border_bg_right_gray.set_font_size(12)

        format_header_border_bg_left_gray = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'left', 'text_wrap': True, 'num_format': '$ #,##0.00'})
        format_header_border_bg_left_gray.set_border(4)
        format_header_border_bg_left_gray.set_bg_color(bg_gray)
        format_header_border_bg_left_gray.set_font_size(12)

        ############ Fin del Detalle #################

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

        format_bold_border_center = workbook.add_format({'bold': True})
        format_bold_border_center.set_border(2)
        format_bold_border_center.set_align('center_across')

        ################### FIN DE ESTILOS #####################

        ################### CABECERA DEL REPORTE ###############
        sheet.insert_image('A1', image_module_path, {'x_scale': 0.1, 'y_scale': 0.1})
        sheet.write('F3','COMPAÑIA:',format_bold_border_tit_gray)
        sheet.write('F4','RFC:',format_bold_border_tit_gray)
        sheet.write('F5','FECHA:',format_bold_border_tit_gray)

        sheet.write('F6','MOVS. DESTINO:',format_bold_border_tit_gray)

        sheet.merge_range('G3:H3',self.env.user.company_id.name,format_bold_border2)
        sheet.merge_range('G4:H4',self.env.user.company_id.vat,format_bold_border2)
        if self.date_to and self.date_from:
            sheet.write('F7','PERIODO:',format_bold_border_tit_gray)
            date_from_c = convert_date_to_MX(str(date_from))
            date_to_c = convert_date_to_MX(str(date_to))
            sheet.merge_range('G7:H7', date_from_c+' A '+date_to_c, format_bold_border2)
        # # sheet.merge_range('G4:H4',format_bold_border2)

        fecha_creacion = convert_date_to_MX(str(fields.Date.today()))
        sheet.merge_range('G5:H5',fecha_creacion,format_bold_border2)
        target_move = ""
        if data['form']['target_move'] == 'all':
            target_move =  "Todas las Entradas"
        if data['form']['target_move'] == 'posted':
            target_move =  "Todas las Entradas Asentadas"

        sheet.merge_range('G6:H6',target_move,format_bold_border2)

        # sheet.merge_range('G6:H6',str(fecha_creacion),format_bold_border_bg_gray)
        # sheet.merge_range('G7:H7', date_from+' al ' +date_to,format_bold_border_bg_yllw)
        # sheet.merge_range('G8:H8', department.name,format_bold_border2)
        sheet.set_column('A:A', 40)
        sheet.set_column('B:B', 20)

        sheet.set_column('D:D', 40)
        sheet.set_column('E:E', 20)

        sheet.set_column('F:F', 20)
        sheet.set_column('G:H', 17)

        sheet.merge_range('B4:D6', report_name.upper(), format_period_title)
        i = 10
        letra_i = 1
        detail_start_data = i+1
        ######################### COMIENZA CON EL PROCESO ############################

        letra_c = indice_to_column_string(letra_i)
        sheet.write(letra_c+str(i), 'Descripción', format_header_border_bg_gray)

        letra_i += 1        
        letra_c = indice_to_column_string(letra_i)
        sheet.write(letra_c+str(i), 'Balance', format_header_border_bg_gray)

        letra_i += 2        
        letra_c = indice_to_column_string(letra_i)
        sheet.write(letra_c+str(i), 'Descripción', format_header_border_bg_gray)

        letra_i += 1        
        letra_c = indice_to_column_string(letra_i)
        sheet.write(letra_c+str(i), 'Balance', format_header_border_bg_gray)

        sheet.freeze_panes(10,9)

        ######################### GRABAMOS LAS LINEAS ############################

        ################## ---- INICIO DE LA PRIMER PRUEBA ---- #################
        #### Saltos de Columna ########
        i+=1
        initial_index =  i
        next_column = False
        count_level_11 = 0
        #### Conteo Letra ####
        letra_i = 1

        last_sum_index = 0
        prev_indexsum = 0

        sumatory_report  =  0.0
        total_left, total_left_cmp, total_right, total_right_cmp = 0.00, 0.00, 0.00, 0.00
        total_general = 0.0
        #### Fin de Saltos de C.#######
        for each in report_lines:
            #### Saltos de Columna ######## 
            each_level = each['level']
            #### Fin de Saltos de C.#######
            print ("#### 01 LEVEL >>>>>>>> ", each['level'])
            print ("######### each >>>>>>> ", each)
            if each['level'] != 0:
                # #### Saltos de Columna ######## 
                if each_level == 1:
                    prev_indexsum = 1
                    if count_level_11 > 0:
                        i = initial_index
                        letra_i += 3

                else:
                    prev_indexsum += 1

                # #### Fin de Saltos de C.#######

                name = ""
                gap = " "
                name = each['name']
                print ("#---- NAME >>>> ", name)

                letra_c = indice_to_column_string(letra_i)
                letra_c_02 = indice_to_column_string(letra_i+1)
                line_type = each['type']
                if each['report_side'] != 'right':
                    #### Saltos de Columna ########
                    print ("### != RIGHT >>>>>>>>>> ")
                    if line_type == 'report':
                        sheet.write(letra_c+str(i), name, format_header_border_bg_left_yll)
                        sheet.write(letra_c_02+str(i), each['balance'], format_header_border_bg_right_yll)
                    else:
                        sheet.write(letra_c+str(i), name, format_header_border_bg_left)
                        sheet.write(letra_c_02+str(i), each['balance'], format_header_border_bg_right)
                    #### Fin de Saltos de C.#######
#                        sheet1.write(row, 2, self.env.user.company_id.currency_id.symbol, left)
                    if each['level'] == 1:
                        total_left += each['balance']
                elif each['report_side'] == 'right':
                    print ("### RIGHT >>>>>>>>>> ")
                    if line_type == 'report':
                        sheet.write(letra_c+str(i), name, format_header_border_bg_left_yll)
                        sheet.write(letra_c_02+str(i), each['balance'], format_header_border_bg_right_yll)
                    else:
                        sheet.write(letra_c+str(i), name, format_header_border_bg_left)
                        sheet.write(letra_c_02+str(i), each['balance'], format_header_border_bg_right)
                    
                #### Saltos de Columna ######## 
                if each_level == 1:
                    count_level_11 += 1 
                if prev_indexsum > last_sum_index:
                    last_sum_index = prev_indexsum
                i+=1
                print ("#::::::::::::::::::: last_sum_index >>>>>>>>>>>>>>> ", last_sum_index)
                #### Fin de Saltos de C.#######
            else:
                if each_level == 0:
                    total_general = each['balance']

        # #### Saltos de Columna ######## 
        last_sum_index += 1
        total_i = initial_index + last_sum_index
        sheet.write('A'+str(total_i), 'Total', format_header_border_bg_left_gray)
        sheet.write('B'+str(total_i), total_general, format_header_border_bg_right)

        #### Fin de Saltos de C.#######

        ################## ---- FIN DE LA PRIMER PRUEBA ---- #################

        ######################### GUARDAMOS EL RESULTADO FINAL ############################


        workbook.close()
        xlsx_data = output.getvalue()

        datas_fname = report_name+".xlsx" # Nombre del Archivo

        self.env.cr.execute(""" DELETE FROM accounting_report_output""")

        attach_id = self.env['accounting.report.output'].create({'name': datas_fname, 'output': base64.encodestring(xlsx_data)})
        return {
                'name': _('Notification'),
                'view_type': 'form',
                'view_mode': 'form',
                'res_model': 'accounting.report.output',
                'res_id': attach_id.id,
                'type': 'ir.actions.act_window',
                'target': 'new'
                }

    

    @api.multi
    def print_excel_report_estado_resultados(self):
        ################ DATA REPORT #################
        data = {}
        data['ids'] = self.env.context.get('active_ids', [])
        data['model'] = self.env.context.get('active_model', 'ir.ui.menu')
        data['form'] = self.read(['date_from', 'date_to', 'journal_ids', 'target_move', 'company_id'])[0]
        used_context = self._build_contexts(data)
        data['form']['used_context'] = dict(used_context, lang=self.env.context.get('lang', 'en_US'))
        # res = self.check_report()
        data['form'].update(self.read(['debit_credit', 'enable_filter', 'label_filter', 'account_report_id', 'date_from_cmp', 'date_to_cmp', 'journal_ids', 'filter_cmp', 'target_move', 'hierarchy_type', 'other_currency'])[0])
#        for field in ['account_report_id']:
#            if isinstance(data['form'][field], tuple):
#                data['form'][field] = data['form'][field][0]
        comparison_context = self._build_comparison_context(data)
        data['form']['comparison_context'] = comparison_context
        print ("###### comparison_context  >>>>>>>>>>>> ", comparison_context)
#        if data['form']['hierarchy_type'] == 'hierarchy':
#            report_lines = self.get_account_lines_hierarchy(data['form'])
#        else:
        report_lines = self.get_account_lines(data['form'])
        print ("###### report_lines  >>>>>>>>>>>> ", report_lines)

        ### Las lineas se identifican por el campo Level #####
        ### Si el Level es 1 indica que debe ir en la siguiente columna ####

        print ("###### data  >>>>>>>>>>>> ", data)
        
        date_to = self.date_to
        date_from = self.date_from

        # print ("\n\n\n\n\nreport_lines",report_lines)

        report_name = data['form']['account_report_id'][1]

        ################ Con XLSXWRITER ##############

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output)

        sheet = workbook.add_worksheet(report_name)

        ################ Logo a Archivo Temporal ##############

        company_logo = self.env.user.company_id.logo
        if company_logo:
            file_result_b64, logo_path_b64 = self.logo_b64_str_to_physical_file(company_logo, 'png', 'company_logo')
            image_module_path = logo_path_b64
        else:
            module_path = module.get_module_path('financial_reports_extend')
            image_module_path = module_path+'/static/img/logo.jpg'

        ############# ESTILOS ###########

        num_format = '$ #,##0.00'
        bg_gray = '#D8D8D8'

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

        f_wh_detail_save = workbook.add_format({'bold': True, 'valign':   'vcenter',  'align':   'center'})

        f_gray_detail_save = workbook.add_format({'bold': True, 'valign':   'vcenter',})
        f_gray_detail_save.set_border(2)
        f_gray_detail_save.set_bg_color(bg_gray)


        f_gray_detail_save_center = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center'})
        f_gray_detail_save_center.set_border(2)
        f_gray_detail_save_center.set_bg_color(bg_gray)

        f_blue_detail_save_center = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center'})
        f_blue_detail_save_center.set_border(2)
        f_blue_detail_save_center.set_font_color('white')
        f_blue_detail_save_center.set_bg_color("#3465a4")

        format_bold_border2 = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center'})
        format_bold_border2.set_border(2)

        format_bold_border_bg_yllw = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center'})
        format_bold_border_bg_yllw.set_border(2)
        format_bold_border_bg_yllw.set_bg_color("#F0FF5B")

        format_bold_border_bg_gray = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center'})
        format_bold_border_bg_gray.set_border(2)
        format_bold_border_bg_gray.set_bg_color(bg_gray)

        format_header_border_bg_gray = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center', 'text_wrap': True})
        format_header_border_bg_gray.set_border(4)
        format_header_border_bg_gray.set_bg_color(bg_gray)
        format_header_border_bg_gray.set_font_size(12)

        ############ Estilos para el Detalle ##########

        format_header_border_bg = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center', 'text_wrap': True})
        format_header_border_bg.set_border(4)
        format_header_border_bg.set_font_size(12)

        format_header_border_bg_left = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'left', 'text_wrap': True})
        format_header_border_bg_left.set_border(4)
        format_header_border_bg_left.set_font_size(12)

        format_header_border_bg_right = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'right', 'text_wrap': True, 'num_format': '$ #,##0.00'})
        format_header_border_bg_right.set_border(4)
        format_header_border_bg_right.set_font_size(12)

        format_header_border_bg_left_yll = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'left', 'text_wrap': True, 'num_format': '$ #,##0.00'})
        format_header_border_bg_left_yll.set_border(4)
        format_header_border_bg_left_yll.set_bg_color("#f7f4be")
        format_header_border_bg_left_yll.set_font_size(12)

        format_header_border_bg_right_yll = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'right', 'text_wrap': True, 'num_format': '$ #,##0.00'})
        format_header_border_bg_right_yll.set_border(4)
        format_header_border_bg_right_yll.set_bg_color("#f7f4be")
        format_header_border_bg_right_yll.set_font_size(12)

        format_header_border_bg_right_gray = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'right', 'text_wrap': True})
        format_header_border_bg_right_gray.set_border(4)
        format_header_border_bg_right_gray.set_bg_color(bg_gray)
        format_header_border_bg_right_gray.set_font_size(12)

        format_header_border_bg_left_gray = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'left', 'text_wrap': True, 'num_format': '$ #,##0.00'})
        format_header_border_bg_left_gray.set_border(4)
        format_header_border_bg_left_gray.set_bg_color(bg_gray)
        format_header_border_bg_left_gray.set_font_size(12)

        ############ Fin del Detalle #################

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

        format_bold_border_center = workbook.add_format({'bold': True})
        format_bold_border_center.set_border(2)
        format_bold_border_center.set_align('center_across')

        ################### FIN DE ESTILOS #####################

        ################### CABECERA DEL REPORTE ###############
        sheet.insert_image('A1', image_module_path, {'x_scale': 0.1, 'y_scale': 0.1})
        sheet.write('F3','COMPAÑIA:',format_bold_border_tit_gray)
        sheet.write('F4','RFC:',format_bold_border_tit_gray)
        sheet.write('F5','FECHA:',format_bold_border_tit_gray)

        sheet.write('F6','MOVS. DESTINO:',format_bold_border_tit_gray)

        sheet.merge_range('G3:H3',self.env.user.company_id.name,format_bold_border2)
        sheet.merge_range('G4:H4',self.env.user.company_id.vat,format_bold_border2)
        if self.date_to and self.date_from:
            sheet.write('F7','PERIODO:',format_bold_border_tit_gray)
            date_from_c = convert_date_to_MX(str(date_from))
            date_to_c = convert_date_to_MX(str(date_to))
            sheet.merge_range('G7:H7', date_from_c+' A '+date_to_c, format_bold_border2)
        # # sheet.merge_range('G4:H4',format_bold_border2)

        fecha_creacion = convert_date_to_MX(str(fields.Date.today()))
        sheet.merge_range('G5:H5',fecha_creacion,format_bold_border2)
        target_move = ""
        if data['form']['target_move'] == 'all':
            target_move =  "Todas las Entradas"
        if data['form']['target_move'] == 'posted':
            target_move =  "Todas las Entradas Asentadas"

        sheet.merge_range('G6:H6',target_move,format_bold_border2)

        # sheet.merge_range('G6:H6',str(fecha_creacion),format_bold_border_bg_gray)
        # sheet.merge_range('G7:H7', date_from+' al ' +date_to,format_bold_border_bg_yllw)
        # sheet.merge_range('G8:H8', department.name,format_bold_border2)
        sheet.set_column('A:A', 40)
        sheet.set_column('B:B', 20)

        sheet.set_column('D:D', 40)
        sheet.set_column('E:E', 20)

        sheet.set_column('F:F', 20)
        sheet.set_column('G:H', 17)

        sheet.merge_range('B4:D6', report_name.upper(), format_period_title)
        i = 10
        desc_i = 'A'
        periodo_i = 'B'
        percent_01_i = 'C'
        acum_i = 'D'
        percent_02_i = 'E'
        ######################### COMIENZA CON EL PROCESO ############################

        sheet.write(desc_i+str(i), 'Descripción', format_header_border_bg_gray)

        sheet.write(periodo_i+str(i), 'Periodo', format_header_border_bg_gray)

        sheet.write(percent_01_i+str(i), '%', format_header_border_bg_gray)

        sheet.write(acum_i+str(i), 'Acumulado', format_header_border_bg_gray)

        sheet.write(percent_02_i+str(i), '%', format_header_border_bg_gray)
        ######################### GRABAMOS LAS LINEAS ############################

        ################## ---- INICIO DE LA PRIMER PRUEBA ---- #################
        #### Saltos de Columna ########
        i+=1
        initial_index =  i
        next_column = False
        count_level_11 = 0
        #### Conteo Letra ####
        sumatory_report  =  0.0
        total_left, total_left_cmp, total_right, total_right_cmp = 0.00, 0.00, 0.00, 0.00
        total_general = 0.0
        #### Fin de Saltos de C.#######
        for each in report_lines:
            #### Saltos de Columna ######## 
            each_level = each['level']
            #### Fin de Saltos de C.#######
            print ("#### 01 LEVEL >>>>>>>> ", each['level'])
            print ("######### each >>>>>>> ", each)
            if each['level'] != 0:
                name = ""
                gap = " "
                name = each['name']
                print ("#---- NAME >>>> ", name)

                line_type = each['type']
                if each['report_side'] != 'right':
                    #### Saltos de Columna ########
                    print ("### != RIGHT >>>>>>>>>> ")
                    if line_type == 'report':
                        sheet.write(desc_i+str(i), name, format_header_border_bg_left_yll)
                        sheet.write(acum_i+str(i), each['balance'], format_header_border_bg_right_yll)
                    else:
                        sheet.write(desc_i+str(i), name, format_header_border_bg_left)
                        sheet.write(acum_i+str(i), each['balance'], format_header_border_bg_right)
                    #### Fin de Saltos de C.#######
#                        sheet1.write(row, 2, self.env.user.company_id.currency_id.symbol, left)
                    if each['level'] == 1:
                        total_left += each['balance']
                elif each['report_side'] == 'right':
                    print ("### RIGHT >>>>>>>>>> ")
                    if line_type == 'report':
                        sheet.write(desc_i+str(i), name, format_header_border_bg_left_yll)
                        sheet.write(acum_i+str(i), each['balance'], format_header_border_bg_right_yll)
                    else:
                        sheet.write(desc_i+str(i), name, format_header_border_bg_left)
                        sheet.write(acum_i+str(i), each['balance'], format_header_border_bg_right)
                    
                #### Saltos de Columna ######## 

                i+=1
                #### Fin de Saltos de C.#######
            else:
                if each_level == 0:
                    total_general = each['balance']

        # #### Saltos de Columna ######## 
        i+=1
        sheet.write('A'+str(i), 'Total', format_header_border_bg_left_gray)
        sheet.write('B'+str(i), total_general, format_header_border_bg_right)

        #### Fin de Saltos de C.#######

        ################## ---- FIN DE LA PRIMER PRUEBA ---- #################

        ######################### GUARDAMOS EL RESULTADO FINAL ############################


        workbook.close()
        xlsx_data = output.getvalue()

        datas_fname = report_name+".xlsx" # Nombre del Archivo

        self.env.cr.execute(""" DELETE FROM accounting_report_output""")

        attach_id = self.env['accounting.report.output'].create({'name': datas_fname, 'output': base64.encodestring(xlsx_data)})
        return {
                'name': _('Notification'),
                'view_type': 'form',
                'view_mode': 'form',
                'res_model': 'accounting.report.output',
                'res_id': attach_id.id,
                'type': 'ir.actions.act_window',
                'target': 'new'
                }

