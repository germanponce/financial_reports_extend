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
        
        fecha_creacion = convert_date_to_MX(str(fields.Date.today()))
        if self.date_from:
            sheet.write('F7','PERIODO:',format_bold_border_tit_gray)
            date_from_c = convert_date_to_MX(str(date_from))
            if not self.date_to:
                date_to_c = fecha_creacion
            else:
                date_to_c = convert_date_to_MX(str(date_to))
            sheet.merge_range('G7:H7', date_from_c+' A '+date_to_c, format_bold_border2)
        # # sheet.merge_range('G4:H4',format_bold_border2)

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


    def get_account_lines(self, data):
        print ("########### get_account_lines >>>>>>>>>>>>> ")
        lines = []
        date_to = self.date_to
        date_from = self.date_from
        print (":::::: date_to >>>>>>> ", date_to)
        print (":::::: date_from >>>>>>> ", date_from)
        # print ("\n\n\n======data['account_report_id'][0]",data['account_report_id'][0])
        account_report = self.env['account.financial.report'].search([('id', '=', data['account_report_id'][0])])
        child_reports = account_report._get_children_by_order()
        print ("### data.get('used_context') >>>>> ",data.get('used_context'))
        res = self.with_context(data.get('used_context'))._compute_report_balance(child_reports)
        print ("### RES >>>>>>>>>> ", res)
        if data['enable_filter']:
            comparison_res = self.with_context(data.get('comparison_context'))._compute_report_balance(child_reports)
            for report_id, value in comparison_res.items():
                res[report_id]['comp_bal'] = value['balance']
                report_acc = res[report_id].get('account')
                if report_acc:
                    for account_id, val in comparison_res[report_id].get('account').items():
                        report_acc[account_id]['comp_bal'] = val['balance']
        for report in child_reports:
            print (":::::: report >>>>>>> ", report)
            print (":::::: name >>>>>>> ", report.name)
            print (":::::: type >>>>>>> ", report.type)
            print ("###### res[report.id].get('account') >>>>>> ", res[report.id].get('account'))
            # print ('\n\n\n===========reprot====',report)
            # print (":::::: res[report.id].get('account') >>>>>>> ", res[report.id].get('account'))
            
            parent_initial_balance = 0.0
            accounts_initital_balance = {}
            parent_balance = 0.0
            if date_from:
                date_context_update = data.get('comparison_context')
                date1 = datetime.strptime(str(date_from),'%Y-%m-%d')
                prev_day_from = date1 - timedelta(days = 1)
                prev_day_from_str = str(prev_day_from)[0:10]
                print ("### prev_day_from >>>>>>>> ", str(prev_day_from))
                date_context_update.update({
                    'date_from': '2019-01-01',
                    'date_to': prev_day_from_str,
                    })
                report_initital_balance = self.with_context(date_context_update)._compute_report_balance(child_reports)
                # print (":::::: CALCULANDO MANUALMENTE >>>> ",report_initital_balance)
                parent_initial_balance = report_initital_balance[report.id]['balance'] * report.sign
                print ("********* parent_initial_balance >>>>>>> ", parent_initial_balance)
                ####### Sacando los balances de las lineas #########
                if report_initital_balance[report.id].get('account'):
                    for account_id, value in report_initital_balance[report.id]['account'].items():
                        flag = False
                        account = self.env['account.account'].browse(account_id)
                        if data['debit_credit']:
                            if not account.company_id.currency_id.is_zero(value['debit']) or not account.company_id.currency_id.is_zero(value['credit']):
                                flag = True
                        vals_balance = value['balance'] * report.sign or 0.0
                        if not account.company_id.currency_id.is_zero(vals_balance):
                            flag = True
                        if data['enable_filter']:
                            vals_balance_cmp = value['comp_bal'] * report.sign
                            if not account.company_id.currency_id.is_zero(vals_balance_cmp):
                                flag = True
                        if flag:
                            accounts_initital_balance.update({
                               account_id: vals_balance,
                            })
            print ("####### accounts_initital_balance >>>>>>>>> ", accounts_initital_balance)
            parent_balance = res[report.id]['balance'] * report.sign
            vals = {
                'name': report.name,
                'balance': parent_balance,
                'type': 'report',
                'level': bool(report.style_overwrite) and report.style_overwrite or report.level,
                'account_type': report.type or False, #used to underline the financial report balances
                'report_side': report.report_side,
                'account_id': False,
                'account_ids': False,
                'initial_balance': parent_initial_balance,
                'parent_initial_balance': parent_initial_balance,
                'parent_balance': parent_balance,
            }
            if report.report_side and report.report_side == 'right':
                data['right'] = True
            # print ("\n\n\n\n===========right=====================",data.get('right'))
            if data['debit_credit']:
                vals['debit'] = res[report.id]['debit']
                vals['credit'] = res[report.id]['credit']

            if data['enable_filter']:
                vals['balance_cmp'] = res[report.id]['comp_bal'] * report.sign
            
            #### Este apartado me funciona cuando calculo manualmente el Saldo Inicial ####
            # ####### Sacando las Cuentas del Nivel 1 #########
            # account_ids = []
            # if res[report.id].get('account'):
            #     for account_id, value in res[report.id]['account'].items():
            #         flag = False
            #         account = self.env['account.account'].browse(account_id)
            #         if data['debit_credit']:
            #             if not account.company_id.currency_id.is_zero(value['debit']) or not account.company_id.currency_id.is_zero(value['credit']):
            #                 flag = True
            #         vals_balance = value['balance'] * report.sign or 0.0
            #         if not account.company_id.currency_id.is_zero(vals_balance):
            #             flag = True
            #         if data['enable_filter']:
            #             vals_balance_cmp = value['comp_bal'] * report.sign
            #             if not account.company_id.currency_id.is_zero(vals_balance_cmp):
            #                 flag = True
            #         if flag:
            #             account_ids.append(account_id)
            # if account_ids:
            #     vals.update({
            #         'account_ids': account_ids,
            #         })

            lines.append(vals)
            if report.display_detail == 'no_detail':
                #the rest of the loop is used to display the details of the financial report, so it's not needed here.
                continue
            
            if res[report.id].get('account'):
                sub_lines = []
                for account_id, value in res[report.id]['account'].items():
                    #if there are accounts to display, we add them to the lines with a level equals to their level in
                    #the COA + 1 (to avoid having them with a too low level that would conflicts with the level of data
                    #financial reports for Assets, liabilities...)
                    flag = False
                    account = self.env['account.account'].browse(account_id)
                    initial_balance = 0.0
                    vals = {
                        'name': account.code + ' ' + account.name,
                        'balance': value['balance'] * report.sign or 0.0,
                        'type': 'account',
                        'level': report.display_detail == 'detail_with_hierarchy' and 4,
                        'account_type': account.internal_type,
                        'report_side': report.report_side,
                        'account_id': account_id,
                        'account_ids': False,
                    }
                    if accounts_initital_balance:
                        if account_id in accounts_initital_balance:
                            initial_balance = accounts_initital_balance[account_id]
                    print ("### initial_balance >>>>> ", initial_balance)
                    vals['initial_balance'] = initial_balance
                    vals['parent_initial_balance'] = parent_initial_balance
                    vals['parent_balance'] = parent_balance

                    if data['debit_credit']:
                        vals['debit'] = value['debit']
                        vals['credit'] = value['credit']
                        if not account.company_id.currency_id.is_zero(vals['debit']) or not account.company_id.currency_id.is_zero(vals['credit']):
                            flag = True
                    if not account.company_id.currency_id.is_zero(vals['balance']):
                        flag = True
                    if data['enable_filter']:
                        vals['balance_cmp'] = value['comp_bal'] * report.sign
                        if not account.company_id.currency_id.is_zero(vals['balance_cmp']):
                            flag = True
                    if flag:
                        sub_lines.append(vals)
                lines += sorted(sub_lines, key=lambda sub_line: sub_line['name'])
        return lines


    def _compute_account_initial_balance(self, accounts, date_from, grouped_by_account=False):
        print ("#### _compute_account_initial_balance >>>>>>> " )
        print (":::::::: accounts >>>>>>> ", accounts )
        print (":::::::: date_from >>>>>>> ", date_from )
        """ compute the balance, debit and credit for the provided accounts
        """
        mapping = {
            'balance': "COALESCE(SUM(debit),0) - COALESCE(SUM(credit), 0) as balance",
            'debit': "COALESCE(SUM(debit), 0) as debit",
            'credit': "COALESCE(SUM(credit), 0) as credit",
        }
        initial_balance = 0.0
        res = {}
        if accounts and date_from:
            tables, where_clause, where_params = self.env['account.move.line']._query_get()
            tables = tables.replace('"', '') if tables else "account_move_line"
            wheres = [""]
            if where_clause.strip():
                wheres.append(where_clause.strip())
            filters = " AND ".join(wheres)
            if date_from:
                filters = filters + " AND date < '%s' " % date_from
            if grouped_by_account:
                request = "SELECT account_id as id, " + ', '.join(mapping.values()) + \
                       " FROM " + tables + \
                       " WHERE account_id IN %s " \
                            + filters + \
                            " GROUP BY account_id"
            else:
                request = "SELECT " + ', '.join(mapping.values()) + \
                       " FROM " + tables + \
                       " WHERE account_id IN %s " \
                            + filters 
            params = (tuple(accounts),) + tuple(where_params)
            # print ("\n\n\n\n\n\n=========params", request % tuple(params))
            # print (":::::::: request >>>>>>> ", request )
            # print (":::::::: params >>>>>>> ", params )
            self.env.cr.execute(request, params)
            res = self.env.cr.dictfetchall()
            if res:
                return res[0]
            print ("#### res >>>>>>> ", res)
        return res



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
        percent_format = '0.00%'
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

        format_header_border_bg_center_percent = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center', 'text_wrap': True, 'num_format': percent_format})
        format_header_border_bg_center_percent.set_border(4)
        format_header_border_bg_center_percent.set_font_size(12)

        format_header_border_bg_left_yll = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'left', 'text_wrap': True, 'num_format': '$ #,##0.00'})
        format_header_border_bg_left_yll.set_border(4)
        format_header_border_bg_left_yll.set_bg_color("#f7f4be")
        format_header_border_bg_left_yll.set_font_size(12)

        format_header_border_bg_right_yll = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'right', 'text_wrap': True, 'num_format': '$ #,##0.00'})
        format_header_border_bg_right_yll.set_border(4)
        format_header_border_bg_right_yll.set_bg_color("#f7f4be")
        format_header_border_bg_right_yll.set_font_size(12)

        format_header_border_bg_center_yll_percent = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'center', 'text_wrap': True, 'num_format': percent_format})
        format_header_border_bg_center_yll_percent.set_border(4)
        format_header_border_bg_center_yll_percent.set_bg_color("#f7f4be")
        format_header_border_bg_center_yll_percent.set_font_size(12)

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
        fecha_creacion = convert_date_to_MX(str(fields.Date.today()))
        if self.date_from :
            sheet.write('F7','PERIODO:',format_bold_border_tit_gray)
            date_from_c = convert_date_to_MX(str(date_from))
            if not self.date_to:
                date_to_c = fecha_creacion
            else:
                date_to_c = convert_date_to_MX(str(date_to))
            sheet.merge_range('G7:H7', date_from_c+' A '+date_to_c, format_bold_border2)
        # # sheet.merge_range('G4:H4',format_bold_border2)

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
        sheet.set_column('C:C', 20)

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

                ### PRUEBA DEL BALANCE - CALCULANDOLO MANUALMENTE ##
                # account_ids = []
                # if each_level == 1:
                #     account_ids = each.get('account_ids', [])
                #     print ("::: account_ids >>>> ", account_ids)
                # else:
                #     account_id = each.get('account_id', False)
                #     if account_id:
                #         account_ids = [account_id]

                # #### Calculando el Balance Inicial de Periodos Anteriores ###
                # compute_initial_balance = self._compute_account_initial_balance(account_ids, date_from)
                # print ("### compute_initial_balance >>>>> ", compute_initial_balance)
                # initial_balance = 0.0
                # if compute_initial_balance:
                #     initial_balance = compute_initial_balance['balance']

                ######## SEGUNDA PRUEBA CON LA RECURSIVIDAD OBTENIDA DEL CALCULO DEL REPORTE ####
                initial_balance = each.get('initial_balance', 0.0)
                parent_initial_balance = each.get('parent_initial_balance', 0.0)
                parent_balance = each.get('parent_balance', 0.0)
                print (":::::::::::: initial_balance >>>>>>>>>>> ", initial_balance)
                print (":::::::::::: parent_initial_balance >>>>>>>>>>> ", parent_initial_balance)
                print (":::::::::::: parent_balance >>>>>>>>>>> ", parent_balance)
                line_balance = each.get('balance', 0.0)
                
                ### Sacando los Porcentajes ###
                ### Periodo #####
                percentage_period = 0.0
                if initial_balance == parent_initial_balance:
                    if initial_balance == 0.0:
                        percentage_period = 0.0
                    else:
                        percentage_period = 1.0
                elif initial_balance == 0.0 and parent_initial_balance == 0.0:
                    percentage_period = 0.0
                else:
                    percentage_period =  initial_balance / parent_initial_balance
                print (":::::::::::: percentage_period >>>>>>>>>>> ", percentage_period)
                ### Acumulado #####
                percentage_acum = 0.0
                if line_balance == parent_balance:
                    percentage_acum = 1.0
                elif line_balance == 0.0 and parent_balance == 0.0:
                    percentage_acum = 0.0
                else:
                    percentage_acum =  line_balance / parent_balance
                print (":::::::::::: percentage_acum >>>>>>>>>>> ", percentage_acum)

                line_type = each['type']
                if each['report_side'] != 'right':
                    #### Saltos de Columna ########
                    print ("### != RIGHT >>>>>>>>>> ")
                    if line_type == 'report':
                        sheet.write(desc_i+str(i), name, format_header_border_bg_left_yll)
                        sheet.write(periodo_i+str(i), initial_balance, format_header_border_bg_right_yll)
                        sheet.write(percent_01_i+str(i), percentage_period, format_header_border_bg_center_yll_percent)
                        sheet.write(acum_i+str(i), line_balance, format_header_border_bg_right_yll)
                        sheet.write(percent_02_i+str(i), percentage_acum, format_header_border_bg_center_yll_percent)
                    else:
                        sheet.write(desc_i+str(i), name, format_header_border_bg_left)
                        sheet.write(periodo_i+str(i), initial_balance, format_header_border_bg_right)
                        sheet.write(percent_01_i+str(i), percentage_period, format_header_border_bg_center_percent)
                        sheet.write(acum_i+str(i), line_balance, format_header_border_bg_right)
                        sheet.write(percent_02_i+str(i), percentage_acum, format_header_border_bg_center_percent)

                    #### Fin de Saltos de C.#######
#                        sheet1.write(row, 2, self.env.user.company_id.currency_id.symbol, left)
                    if each['level'] == 1:
                        total_left += each['balance']
                elif each['report_side'] == 'right':
                    print ("### RIGHT >>>>>>>>>> ")
                    if line_type == 'report':
                        sheet.write(desc_i+str(i), name, format_header_border_bg_left_yll)
                        sheet.write(periodo_i+str(i), initial_balance, format_header_border_bg_right_yll)
                        sheet.write(percent_01_i+str(i), percentage_period, format_header_border_bg_center_yll_percent)
                        sheet.write(acum_i+str(i), line_balance, format_header_border_bg_right_yll)
                        sheet.write(percent_02_i+str(i), percentage_acum, format_header_border_bg_center_yll_percent)
                    else:
                        sheet.write(desc_i+str(i), name, format_header_border_bg_left)
                        sheet.write(periodo_i+str(i), initial_balance, format_header_border_bg_right)
                        sheet.write(percent_01_i+str(i), percentage_period, format_header_border_bg_center_percent)
                        sheet.write(acum_i+str(i), line_balance, format_header_border_bg_right)
                        sheet.write(percent_02_i+str(i), percentage_acum, format_header_border_bg_center_percent)

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

