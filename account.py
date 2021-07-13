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
        self.ensure_one()
        
        # general_balance_2_columns
        # estado_resultados_extended
        tables, where_clause, where_params = self.env['account.move.line']._query_get()
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
#        if data['form']['hierarchy_type'] == 'hierarchy':
#            report_lines = self.get_account_lines_hierarchy(data['form'])
#        else:
        report_lines = self.get_account_lines(data['form'])

        ### Las lineas se identifican por el campo Level #####
        ### Si el Level es 1 indica que debe ir en la siguiente columna ####

        
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
        format_period_title.set_border(1)
        # format_period_title.set_bottom(2)
        # format_period_title.set_top(2)

        format_bold_border = workbook.add_format({'bold': True, 'valign':   'vcenter'})
        format_bold_border.set_border(2)

        format_bold_border_tit_gray = workbook.add_format({'bold': True, 'valign':   'vcenter'})
        format_bold_border_tit_gray.set_border(1)
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
        format_bold_border2.set_border(1)

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

        format_header_border_bg_left = workbook.add_format({'valign':   'vcenter', 'align':   'left', 'text_wrap': True})
        format_header_border_bg_left.set_border(4)
        format_header_border_bg_left.set_font_size(12)

        format_header_border_bg_right = workbook.add_format({ 'valign':   'vcenter', 'align':   'right', 'text_wrap': True, 'num_format': '$ #,##0.00'})
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
        sheet.write('D2','COMPAÑIA:',format_bold_border_tit_gray)
        sheet.write('D3','RFC:',format_bold_border_tit_gray)
        sheet.write('D4','FECHA:',format_bold_border_tit_gray)

        sheet.write('D5','MOVS. DESTINO:',format_bold_border_tit_gray)

        sheet.merge_range('E2:F2',self.env.user.company_id.name,format_bold_border2)
        sheet.merge_range('E3:F3',self.env.user.company_id.vat,format_bold_border2)
        
        fecha_creacion = convert_date_to_MX(str(fields.Date.today()))
        if self.date_from:
            sheet.write('D6','PERIODO:',format_bold_border_tit_gray)
            date_from_c = convert_date_to_MX(str(date_from))
            if not self.date_to:
                date_to_c = fecha_creacion
            else:
                date_to_c = convert_date_to_MX(str(date_to))
            sheet.write('E6:F6', date_from_c+' A '+date_to_c, format_bold_border2)
        sheet.merge_range('E4:F4',fecha_creacion,format_bold_border2)
        # # sheet.merge_range('G4:H4',format_bold_border2)

        
        target_move = ""
        if data['form']['target_move'] == 'all':
            target_move =  "Todas las Entradas"
        if data['form']['target_move'] == 'posted':
            target_move =  "Todas las Entradas Asentadas"

        sheet.merge_range('E5:F5',target_move,format_bold_border2)

        # sheet.merge_range('G6:H6',str(fecha_creacion),format_bold_border_bg_gray)
        # sheet.merge_range('G7:H7', date_from+' al ' +date_to,format_bold_border_bg_yllw)
        # sheet.merge_range('G8:H8', department.name,format_bold_border2)
        sheet.set_column('A:A', 45)
        sheet.set_column('B:B', 20)

        sheet.set_column('D:D', 45)
        sheet.set_column('E:E', 20)

        sheet.set_column('F:F', 20)
        sheet.set_column('G:H', 17)

        sheet.merge_range('A7:F8', report_name.upper(), format_period_title)
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
            ###### Estilos Dinamicos ######
            format_header_border_bg_lft_yll_dyn = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'left', 'text_wrap': True, 'num_format': '$ #,##0.00'})
            format_header_border_bg_lft_yll_dyn.set_border(4)
            format_header_border_bg_lft_yll_dyn.set_bg_color("#f7f4be")
            format_header_border_bg_lft_yll_dyn.set_font_size(12)

            format_header_border_bg_left_dyn = workbook.add_format({ 'valign':   'vcenter', 'align':   'left', 'text_wrap': True})
            format_header_border_bg_left_dyn.set_border(4)
            format_header_border_bg_left_dyn.set_font_size(12)

            #### Saltos de Columna ######## 
            each_level = each['level']
            #### Fin de Saltos de C.#######
            if each['level'] != 0:
                # #### Saltos de Columna ######## 
                if each_level == 1:
                    format_header_border_bg_lft_yll_dyn.set_indent(0)
                    format_header_border_bg_left_dyn.set_indent(0)
                    prev_indexsum = 1
                    if count_level_11 > 0:
                        i = initial_index
                        letra_i += 3

                else:
                    prev_indexsum += 1
                    ### Indentation ####                    
                    format_header_border_bg_lft_yll_dyn.set_indent(each_level)
                    format_header_border_bg_left_dyn.set_indent(each_level)

                # #### Fin de Saltos de C.#######

                name = ""
                gap = " "
                name = each['name']

                letra_c = indice_to_column_string(letra_i)
                letra_c_02 = indice_to_column_string(letra_i+1)
                line_type = each['type']
                if each['report_side'] != 'right':
                    #### Saltos de Columna ########
                    if line_type == 'report':
                        sheet.write(letra_c+str(i), name, format_header_border_bg_lft_yll_dyn)
                        sheet.write(letra_c_02+str(i), each['balance'], format_header_border_bg_right_yll)
                    else:
                        sheet.write(letra_c+str(i), name, format_header_border_bg_left_dyn)
                        sheet.write(letra_c_02+str(i), each['balance'], format_header_border_bg_right)
                    #### Fin de Saltos de C.#######
#                        sheet1.write(row, 2, self.env.user.company_id.currency_id.symbol, left)
                    if each['level'] == 1:
                        total_left += each['balance']
                elif each['report_side'] == 'right':
                    if line_type == 'report':
                        sheet.write(letra_c+str(i), name, format_header_border_bg_lft_yll_dyn)
                        sheet.write(letra_c_02+str(i), each['balance'], format_header_border_bg_right_yll)
                    else:
                        sheet.write(letra_c+str(i), name, format_header_border_bg_left_dyn)
                        sheet.write(letra_c_02+str(i), each['balance'], format_header_border_bg_right)
                    
                #### Saltos de Columna ######## 
                if each_level == 1:
                    count_level_11 += 1 
                if prev_indexsum > last_sum_index:
                    last_sum_index = prev_indexsum
                i+=1
                #### Fin de Saltos de C.#######
            else:
                if each_level == 0:
                    total_general = each['balance']

        # #### Saltos de Columna ######## 
        last_sum_index += 1
        total_i = initial_index + last_sum_index
        ### Por el momento no sacamos el Total ####
        # sheet.write('A'+str(total_i), 'Total', format_header_border_bg_left_gray)
        # sheet.write('B'+str(total_i), total_general, format_header_border_bg_right)

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
        lines = []
        date_to = self.date_to
        date_from = self.date_from
        # print ("\n\n\n======data['account_report_id'][0]",data['account_report_id'][0])
        account_report = self.env['account.financial.report'].search([('id', '=', data['account_report_id'][0])])
        child_reports = account_report._get_children_by_order()
        res = self.with_context(data.get('used_context'))._compute_report_balance(child_reports)
        if data['enable_filter']:
            comparison_res = self.with_context(data.get('comparison_context'))._compute_report_balance(child_reports)
            for report_id, value in comparison_res.items():
                res[report_id]['comp_bal'] = value['balance']
                report_acc = res[report_id].get('account')
                if report_acc:
                    for account_id, val in comparison_res[report_id].get('account').items():
                        report_acc[account_id]['comp_bal'] = val['balance']
        for report in child_reports:
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
                date_context_update.update({
                    'date_from': '2019-01-01',
                    'date_to': prev_day_from_str,
                    })
                report_initital_balance = self.with_context(date_context_update)._compute_report_balance(child_reports)
                # print (":::::: CALCULANDO MANUALMENTE >>>> ",report_initital_balance)
                parent_initial_balance = report_initital_balance[report.id]['balance'] * report.sign
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
#        if data['form']['hierarchy_type'] == 'hierarchy':
#            report_lines = self.get_account_lines_hierarchy(data['form'])
#        else:
        report_lines = self.get_account_lines(data['form'])

        ### Las lineas se identifican por el campo Level #####
        ### Si el Level es 1 indica que debe ir en la siguiente columna ####
        
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
        format_period_title.set_border(1)
        # format_period_title.set_bottom(2)
        # format_period_title.set_top(2)

        format_bold_border = workbook.add_format({'bold': True, 'valign':   'vcenter'})
        format_bold_border.set_border(2)

        format_bold_border_tit_gray = workbook.add_format({'bold': True, 'valign':   'vcenter'})
        format_bold_border_tit_gray.set_border(1)
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
        format_bold_border2.set_border(1)

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

        format_header_border_bg_left = workbook.add_format({'valign':   'vcenter', 'align':   'left', 'text_wrap': True})
        format_header_border_bg_left.set_border(4)
        format_header_border_bg_left.set_font_size(12)

        format_header_border_bg_right = workbook.add_format({'valign':   'vcenter', 'align':   'right', 'text_wrap': True, 'num_format': '$ #,##0.00'})
        format_header_border_bg_right.set_border(4)
        format_header_border_bg_right.set_font_size(12)

        format_header_border_bg_center_percent = workbook.add_format({'valign':   'vcenter', 'align':   'center', 'text_wrap': True, 'num_format': percent_format})
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
        sheet.write('D2','COMPAÑIA:',format_bold_border_tit_gray)
        sheet.write('D3','RFC:',format_bold_border_tit_gray)
        sheet.write('D4','FECHA:',format_bold_border_tit_gray)

        sheet.write('D5','MOVS. DESTINO:',format_bold_border_tit_gray)

        sheet.write('E2',self.env.user.company_id.name,format_bold_border2)
        sheet.write('E3',self.env.user.company_id.vat,format_bold_border2)
        fecha_creacion = convert_date_to_MX(str(fields.Date.today()))
        if self.date_from :
            sheet.write('D6','PERIODO:',format_bold_border_tit_gray)
            date_from_c = convert_date_to_MX(str(date_from))
            if not self.date_to:
                date_to_c = fecha_creacion
            else:
                date_to_c = convert_date_to_MX(str(date_to))
            sheet.write('E6', date_from_c+' A '+date_to_c, format_bold_border2)

        sheet.write('E4',fecha_creacion,format_bold_border2)
        # # sheet.merge_range('G4:H4',format_bold_border2)

        
        target_move = ""
        if data['form']['target_move'] == 'all':
            target_move =  "Todas las Entradas"
        if data['form']['target_move'] == 'posted':
            target_move =  "Todas las Entradas Asentadas"

        sheet.write('E5',target_move,format_bold_border2)

        # sheet.merge_range('G6:H6',str(fecha_creacion),format_bold_border_bg_gray)
        # sheet.merge_range('G7:H7', date_from+' al ' +date_to,format_bold_border_bg_yllw)
        # sheet.merge_range('G8:H8', department.name,format_bold_border2)
        sheet.set_column('A:A', 50)
        sheet.set_column('B:B', 20)
        sheet.set_column('C:C', 20)

        sheet.set_column('D:D', 40)
        sheet.set_column('E:E', 40)

        sheet.set_column('F:F', 20)
        sheet.set_column('G:H', 17)

        sheet.merge_range('A8:E9', report_name.upper(), format_period_title)
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
            ###### Estilos Dinamicos ######
            format_header_border_bg_lft_yll_dyn = workbook.add_format({'bold': True, 'valign':   'vcenter', 'align':   'left', 'text_wrap': True, 'num_format': '$ #,##0.00'})
            format_header_border_bg_lft_yll_dyn.set_border(4)
            format_header_border_bg_lft_yll_dyn.set_bg_color("#f7f4be")
            format_header_border_bg_lft_yll_dyn.set_font_size(12)

            format_header_border_bg_left_dyn = workbook.add_format({'valign':   'vcenter', 'align':   'left', 'text_wrap': True})
            format_header_border_bg_left_dyn.set_border(4)
            format_header_border_bg_left_dyn.set_font_size(12)
            
            #### Fin de Saltos de C.#######
            if each['level'] != 0:
                name = ""
                gap = " "
                name = each['name']
                if each_level == 1:
                    format_header_border_bg_lft_yll_dyn.set_indent(0)
                    format_header_border_bg_left_dyn.set_indent(0) 
                else:
                    format_header_border_bg_lft_yll_dyn.set_indent(each_level)
                    format_header_border_bg_left_dyn.set_indent(each_level)

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
                ### Acumulado #####
                percentage_acum = 0.0
                if line_balance == parent_balance:
                    percentage_acum = 1.0
                elif line_balance == 0.0 and parent_balance == 0.0:
                    percentage_acum = 0.0
                else:
                    percentage_acum =  line_balance / parent_balance

                line_type = each['type']
                if each['report_side'] != 'right':
                    #### Saltos de Columna ########
                    if line_type == 'report':
                        sheet.write(desc_i+str(i), name, format_header_border_bg_lft_yll_dyn)
                        sheet.write(acum_i+str(i), initial_balance, format_header_border_bg_right_yll)
                        sheet.write(percent_02_i+str(i), percentage_period, format_header_border_bg_center_yll_percent)
                        sheet.write(periodo_i+str(i), line_balance, format_header_border_bg_right_yll)
                        sheet.write(percent_01_i+str(i), percentage_acum, format_header_border_bg_center_yll_percent)

                        # sheet.write(periodo_i+str(i), initial_balance, format_header_border_bg_right_yll)
                        # sheet.write(periodo_i+str(i), initial_balance, format_header_border_bg_right_yll)
                        # sheet.write(percent_01_i+str(i), percentage_period, format_header_border_bg_center_yll_percent)
                        # sheet.write(acum_i+str(i), line_balance, format_header_border_bg_right_yll)
                        # sheet.write(percent_02_i+str(i), percentage_acum, format_header_border_bg_center_yll_percent)
                    else:
                        sheet.write(desc_i+str(i), name, format_header_border_bg_left_dyn)
                        sheet.write(acum_i+str(i), initial_balance, format_header_border_bg_right)
                        sheet.write(percent_02_i+str(i), percentage_period, format_header_border_bg_center_percent)
                        sheet.write(periodo_i+str(i), line_balance, format_header_border_bg_right)
                        sheet.write(percent_01_i+str(i), percentage_acum, format_header_border_bg_center_percent)

                        # sheet.write(periodo_i+str(i), initial_balance, format_header_border_bg_right)
                        # sheet.write(percent_01_i+str(i), percentage_period, format_header_border_bg_center_percent)
                        # sheet.write(acum_i+str(i), line_balance, format_header_border_bg_right)
                        # sheet.write(percent_02_i+str(i), percentage_acum, format_header_border_bg_center_percent)


                    #### Fin de Saltos de C.#######
#                        sheet1.write(row, 2, self.env.user.company_id.currency_id.symbol, left)
                    if each['level'] == 1:
                        total_left += each['balance']
                elif each['report_side'] == 'right':
                    if line_type == 'report':
                        sheet.write(desc_i+str(i), name, format_header_border_bg_lft_yll_dyn)
                        sheet.write(acum_i+str(i), initial_balance, format_header_border_bg_right_yll)
                        sheet.write(percent_02_i+str(i), percentage_period, format_header_border_bg_center_yll_percent)
                        sheet.write(periodo_i+str(i), line_balance, format_header_border_bg_right_yll)
                        sheet.write(percent_01_i+str(i), percentage_acum, format_header_border_bg_center_yll_percent)

                        # sheet.write(periodo_i+str(i), initial_balance, format_header_border_bg_right_yll)
                        # sheet.write(percent_01_i+str(i), percentage_period, format_header_border_bg_center_yll_percent)
                        # sheet.write(acum_i+str(i), line_balance, format_header_border_bg_right_yll)
                        # sheet.write(percent_02_i+str(i), percentage_acum, format_header_border_bg_center_yll_percent)

                    else:
                        sheet.write(desc_i+str(i), name, format_header_border_bg_left_dyn)
                        sheet.write(acum_i+str(i), initial_balance, format_header_border_bg_right)
                        sheet.write(percent_02_i+str(i), percentage_period, format_header_border_bg_center_percent)
                        sheet.write(periodo_i+str(i), line_balance, format_header_border_bg_right)
                        sheet.write(percent_01_i+str(i), percentage_acum, format_header_border_bg_center_percent)

                        # sheet.write(periodo_i+str(i), initial_balance, format_header_border_bg_right)
                        # sheet.write(percent_01_i+str(i), percentage_period, format_header_border_bg_center_percent)
                        # sheet.write(acum_i+str(i), line_balance, format_header_border_bg_right)
                        # sheet.write(percent_02_i+str(i), percentage_acum, format_header_border_bg_center_percent)

                #### Saltos de Columna ######## 

                i+=1
                #### Fin de Saltos de C.#######
            else:
                if each_level == 0:
                    total_general = each['balance']

        # #### Saltos de Columna ######## 
        ### De momento sin Total ####
        # i+=1
        # sheet.write('A'+str(i), 'Total', format_header_border_bg_left_gray)
        # sheet.write('B'+str(i), total_general, format_header_border_bg_right)

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



################## REPORTES DE LA OCA ####################
######################### INICIO #########################


class GeneralLedgerReportWizard(models.TransientModel):
    _inherit = "general.ledger.report.wizard"

    hide_partners = fields.Boolean('Sin desgloce de Empresas', 
                                    help='Oculta el detalle de Empresas y agregar las columnas Saldo Inicial y Final', 
                                    default=True)

    # add_initial_balance = fields.Boolean('Agregar Balance Inicial')

    def _compute_account_initial_balance(self, accounts, date_from, grouped_by_account=False):
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
        return res


    @api.multi
    def button_export_xlsx(self):
        self.ensure_one()
        report_type = 'xlsx'
        date_from = self.date_from
        date_to = self.date_to

        date_context_update = dict(self._context)
        date1 = datetime.strptime(str(date_from),'%Y-%m-%d')
        prev_day_from = date1 - timedelta(days = 1)
        prev_day_from_str = str(prev_day_from)[0:10]
        date_context_update.update({
            'initial_balance_date_from': '2019-01-01',
            'initial_balance_date_to': prev_day_from_str,
            'hide_partners': self.hide_partners,
            })
        return self.with_context(date_context_update)._export(report_type)

    def _export(self, report_type):
        """Default export is PDF."""
        model = self.env['report_general_ledger']
        report = model.create(self._prepare_report_general_ledger())
        report.compute_data_for_report()
        context = self._context

        return report.with_context(context).print_report(report_type)


class GeneralLedgerReportCompute(models.TransientModel):
    _inherit = 'report_general_ledger'


    @api.multi
    def print_report(self, report_type):
        context = dict(self._context)
        self.ensure_one()
        if report_type == 'xlsx':
            
            hide_partners = context.get('hide_partners', False)
            if hide_partners:
                report_name = 'a_f_r.report_general_ledger_xlsx_grouped'
            else:
                report_name = 'a_f_r.report_general_ledger_xlsx'
        else:
            report_name = 'account_financial_report.' \
                          'report_general_ledger_qweb'
        report_result = self.env['ir.actions.report'].search(
            [('report_name', '=', report_name),
             ('report_type', '=', report_type)], limit=1)
        return report_result.report_action(self)

    @api.multi
    def compute_data_for_report(self,
                                with_line_details=True,
                                with_partners=True):
        context = self._context
        params = context.get('params', {})  
        initial_balance_date_from = context.get('initial_balance_date_from', False)
        initial_balance_date_to = context.get('initial_balance_date_to', False)
        hide_partners = context.get('hide_partners', False)
        res = super().compute_data_for_report(with_line_details=with_line_details,
                                              with_partners=with_partners)
        return res

class GeneralLedgerXslxGrouped(models.AbstractModel):
    _name = 'report.a_f_r.report_general_ledger_xlsx_grouped'
    _inherit = 'report.account_financial_report.abstract_report_xlsx'

    def _get_report_name(self, report):
        report_name = 'Libro Mayor'
        return self._get_report_complete_name(report, report_name)

    def _get_report_columns(self, report):
        context = dict(self._context)

        res = {
            0: {'header': _('Date'), 'field': 'date', 'width': 11},
            1: {'header': _('Entry'), 'field': 'entry', 'width': 18},
            2: {'header': _('Journal'), 'field': 'journal', 'width': 8},
            3: {'header': _('Account'), 'field': 'account', 'width': 9},
            4: {'header': _('Taxes'),
                'field': 'taxes_description',
                'width': 15},
            5: {'header': _('Partner'), 'field': 'partner', 'width': 25},
            6: {'header': _('Ref - Label'), 'field': 'label', 'width': 40},
            7: {'header': _('Cost center'),
                'field': 'cost_center',
                'width': 15},
            8: {'header': _('Tags'),
                'field': 'tags',
                'width': 10},
            9: {'header': _('Rec.'), 'field': 'matching_number', 'width': 5},
            10: {'header': _('Saldo Inicial'),
                 'field': 'initial_balance',
                 'field_initial_balance': 'initial_balance',
                 'field_final_balance': 'final_balance',
                 'type': 'amount',
                 'width': 14},
            11: {'header': _('Debit'),
                 'field': 'debit',
                 'field_initial_balance': 'initial_debit',
                 'field_final_balance': 'final_debit',
                 'type': 'amount',
                 'width': 14},
            12: {'header': _('Credit'),
                 'field': 'credit',
                 'field_initial_balance': 'initial_credit',
                 'field_final_balance': 'final_credit',
                 'type': 'amount',
                 'width': 14},
            13: {'header': _('Cumul. Bal.'),
                 'field': 'cumul_balance',
                 'field_initial_balance': 'initial_balance',
                 'field_final_balance': 'final_balance',
                 'type': 'amount',
                 'width': 14},
        }
        if report.foreign_currency:
            foreign_currency = {
                13: {'header': _('Cur.'),
                     'field': 'currency_id',
                     'field_currency_balance': 'currency_id',
                     'type': 'many2one', 'width': 7},
                14: {'header': _('Amount cur.'),
                     'field': 'amount_currency',
                     'field_initial_balance':
                         'initial_balance_foreign_currency',
                     'field_final_balance':
                         'final_balance_foreign_currency',
                     'type': 'amount_currency',
                     'width': 14},
            }
            res = {**res, **foreign_currency}
        return res

    def _get_report_filters(self, report):
        return [
            [
                _('Date range filter'),
                _('From: %s To: %s') % (report.date_from, report.date_to),
            ],
            [
                _('Target moves filter'),
                _('All posted entries') if report.only_posted_moves
                else _('All entries'),
            ],
            [
                _('Account balance at 0 filter'),
                _('Hide') if report.hide_account_at_0 else _('Show'),
            ],
            [
                _('Centralize filter'),
                _('Yes') if report.centralize else _('No'),
            ],
            [
                _('Show analytic tags'),
                _('Yes') if report.show_analytic_tags else _('No'),
            ],
            [
                _('Show foreign currency'),
                _('Yes') if report.foreign_currency else _('No')
            ],
        ]

    def _get_col_count_filter_name(self):
        return 2

    def _get_col_count_filter_value(self):
        return 2

    def _get_col_pos_initial_balance_label(self):
        return 5

    def _get_col_count_final_balance_name(self):
        return 5

    def _get_col_pos_final_balance_label(self):
        return 5

    def _generate_report_content(self, workbook, report):
        context =  self._context
        # hide_partners
        # print ("###### _generate_report_content >>>>>>>> ")
        # print ("####################### context >>>>>>>> ",context)
        # For each account
        account_move_line = self.env['account.move.line']
        for account in report.account_ids:
            # Write account title
            self.write_array_title(account.code + ' - ' + account.name)
            # print ("##### if not account.partner_ids >>>>>>>> ", account.partner_ids)
            ending_balance_vals_summatory = {
                                             10: 0.0,# initial_balance,
                                             11: 0.0, # debit',
                                             12: 0.0, # credit',
                                             13: 0.0, # cumul_balance',
                                            }
            account_read = account.read()[0]
            sum_initial_balance = account_read.get('field_initial_balance')
            sum_debit = 0.0
            sum_credit = 0.0
            sum_cumul_balance = account_read.get('cumul_balance')
            if not account.partner_ids:
                # Display array header for move lines
                self.write_array_header()

                # Display initial balance line for account
                self.write_initial_balance_special(account)
                # Display account move lines
                grouped_lines = {}
                for line in account.move_line_ids:
                    line_read = line.read()[0]
                    line_journal = line_read['journal']
                    line_account = line_read.get('report_account_id','')

                    line_debit = line_read.get('debit', 0.0)
                    line_credit = line_read.get('credit', 0.0)
                    line_cumul_balance = line_read.get('cumul_balance', 0.0)
                    ### Sumatorias ###
                    sum_debit += float(line_debit)
                    sum_credit += float(line_credit)

                    vals = {line:{
                        0: line_read['date'],
                        1: line_read['entry'],
                        2: line_read['journal'],
                        3: line_read['account'],
                        4: line_read['taxes_description'],
                        5: line_read['partner'],
                        6: line_read['label'],
                        7: line_read['cost_center'],
                        8: line_read['tags'],
                        9: line_read['matching_number'],
                        11: line_debit,
                        12: line_credit,
                        13: line_cumul_balance,
                    }
                    }
                    grouped_lines.update(vals)
                    
                if grouped_lines:
                    self.write_lines_grouped(grouped_lines)
            else:
                # For each partner
                self.write_array_header()

                # Display initial balance line for account
                self.write_initial_balance_special(account)
                grouped_lines = {}
                for partner in account.partner_ids:
                    # Write partner title
                    #self.write_array_title(partner.name)

                    # Display array header for move lines
                    #self.write_array_header()

                    # Display initial balance line for partner
                    #self.write_initial_balance_special(partner)

                    # Display account move lines
                    for line in partner.move_line_ids:
                        line_read = line.read()[0]
                        line_journal = line_read['journal']
                        line_account = line_read.get('report_account_id','')
                        if not line_account:
                            move_line_id = line_read.get('move_line_id','')
                            acmv_line = account_move_line.browse(move_line_id[0])
                            line_account = acmv_line.account_id.name_get()[0][1]
                        line_debit = line_read.get('debit', 0.0)
                        line_credit = line_read.get('credit', 0.0)                       
                        line_cumul_balance = line_read.get('cumul_balance', 0.0)
                        ### Sumatorias ###
                        sum_debit += float(line_debit)
                        sum_credit += float(line_credit)

                        
                        vals = {line:{
                            0: line_read['date'],
                            1: line_read['entry'],
                            2: line_read['journal'],
                            3: line_read['account'],
                            4: line_read['taxes_description'],
                            5: line_read['partner'],
                            6: line_read['label'],
                            7: line_read['cost_center'],
                            8: line_read['tags'],
                            9: line_read['matching_number'],
                            11: line_debit,
                            12: line_credit,
                            13: line_cumul_balance,
                        }
                        }
                        grouped_lines.update(vals)
                if grouped_lines:
                    self.write_lines_grouped(grouped_lines)
                # print ("#### line_read >>>>>> ", line_read)
                # self.write_line_special(line)

                # Display ending balance line for partner
                # self.write_ending_balance_special(partner)

                # Line break
                self.row_pos += 1

            # Display ending balance line for account
            ending_balance_vals_summatory = [{
                                             10: sum_initial_balance,# initial_balance,
                                             11: sum_debit, # debit',
                                             12: sum_credit, # credit',
                                             13: sum_cumul_balance, # cumul_balance',
                                            }]

            if not report.filter_partner_ids:
                self.write_ending_balance_special(account, ending_balance_vals_summatory)

            # 2 lines break
            self.row_pos += 2

    def write_initial_balance_special(self, my_object):
        """Specific function to write initial balance for General Ledger"""
        if 'partner' in my_object._name:
            label = _('Partner Initial balance')
            my_object.currency_id = my_object.report_account_id.currency_id
        elif 'account' in my_object._name:
            label = _('Initial balance')
        super(GeneralLedgerXslxGrouped, self).write_initial_balance_special(
            my_object, label
        )

    def write_ending_balance_special(self, my_object, list_summ):
        """Specific function to write ending balance for General Ledger"""
        if 'partner' in my_object._name:
            name = my_object.name
            label = _('Partner ending balance')
        elif 'account' in my_object._name:
            name = my_object.code + ' - ' + my_object.name
            label = _('Ending balance')
        super(GeneralLedgerXslxGrouped, self).write_ending_balance_special(
            my_object, name, label, list_summ
        )

class AbstractReportXslx(models.AbstractModel):
    _inherit = 'report.account_financial_report.abstract_report_xlsx'

    def write_lines_grouped(self, lines_dict):

        for line in lines_dict:
            vals = lines_dict.get(line)
            columns = vals.keys()
            # # Fecha #
            # self.sheet.write_string(self.row_pos, 0, '')
            # # Diario #
            # self.sheet.write_string(self.row_pos, 2, journal)

            for col_pos in columns:
                value = vals[col_pos]
                if value:
                    if col_pos in (0,1,2,3,4,5,6,7,8,9):
                        self.sheet.write_string(self.row_pos, col_pos, value)
                    else:
                        cell_format = self.format_amount
                        self.sheet.write_number(
                            self.row_pos, col_pos, float(value), cell_format
                        )
                else:
                    if col_pos in (11,12,13):
                        cell_format = self.format_amount
                        self.sheet.write_number(
                            self.row_pos, col_pos, float(0.0), cell_format
                        )

            self.row_pos += 1

    def write_line_special(self, line_object):
        # print ("######### write_line_special >>>>>>>>>>>>>>>> ")
        # print ("######### line_object >>>>>>>>>>>>>>>> ", line_object)
        """Write a line on current line using all defined columns field name.
        Columns are defined with `_get_report_columns` method.
        """
        for col_pos, column in self.columns.items():
            # print (":::::::: col_pos >>> ",col_pos)
            # print (":::::::: column >>> ",column)
            if col_pos == 10:
                self.sheet.write_string(self.row_pos, col_pos, '',
                                                self.format_bold)
            else:
                value = getattr(line_object, column['field'])
                cell_type = column.get('type', 'string')
                # print (":::::::: cell_type >>> ",cell_type)
                # print (":::::::: value >>> ",value)
                if cell_type == 'many2one':
                    self.sheet.write_string(
                        self.row_pos, col_pos, value.name or '', self.format_right)
                elif cell_type == 'string':
                    if hasattr(line_object, 'account_group_id') and \
                            line_object.account_group_id:
                        self.sheet.write_string(self.row_pos, col_pos, value or '',
                                                self.format_bold)
                    else:
                        self.sheet.write_string(self.row_pos, col_pos, value or '')
                elif cell_type == 'amount':
                    if hasattr(line_object, 'account_group_id') and \
                            line_object.account_group_id:
                        cell_format = self.format_amount_bold
                    else:
                        cell_format = self.format_amount
                    self.sheet.write_number(
                        self.row_pos, col_pos, float(value), cell_format
                    )
                elif cell_type == 'amount_currency':
                    if line_object.currency_id:
                        format_amt = self._get_currency_amt_format(line_object)
                        self.sheet.write_number(
                            self.row_pos, col_pos, float(value), format_amt
                        )
        self.row_pos += 1


    def write_initial_balance_special(self, my_object, label):
        # print ("### write_initial_balance_special >>>>>>>>> ")
        # print ("::: my_object >>>>>>>>> ", my_object)
        vals_read = my_object.read()[0]
        # print ("::: label >>>>>>>>> ", label)
        """Write a specific initial balance line on current line
        using defined columns field_initial_balance name.

        Columns are defined with `_get_report_columns` method.
        """
        account_id = False
        account_br = False
        if 'account_id' in vals_read:
            account_id = vals_read['account_id']
        else:
            if 'account_id' in vals_read:
                account_id = vals_read['report_account_id']
        account_sign = ""
        if account_id:
            account_br = self.env['account.account'].browse(account_id[0])
            if 'cuenta_tipo' in account_br._fields:
                cuenta_tipo = account_br.cuenta_tipo
                if cuenta_tipo == 'D':
                    account_sign = 1
                elif cuenta_tipo == 'A':
                    account_sign = -1
            else:
                if 'sign' in account_br._fields:
                    account_sign = account_br.sign if account_br else False
            # cuenta_tipo = account_br.cuenta_tipo
            # if cuenta_tipo == 'D':
            #     account_sign = 1
            # elif cuenta_tipo == 'A':
            #     account_sign = -1
            #account_sign = account_br.sign if account_br else False

        # print ("### account_sign >>>>>>>>> ", account_sign)
        # print ("### account_br >>>>>>>>> ", account_br)

        col_pos_label = self._get_col_pos_initial_balance_label()
        self.sheet.write(self.row_pos, col_pos_label, label, self.format_right)
        for col_pos, column in self.columns.items():
            # print ("######### col_pos >>>>>>>>>>> ",col_pos)
            # print ("######### column >>>>>>>>>>> ",column)
            if column.get('field_initial_balance'):
                value = getattr(my_object, column['field_initial_balance'])
                # print ("############# VALUE >>>>>>>>> ", value)
                cell_type = column.get('type', 'string')
                if cell_type == 'string':
                    self.sheet.write_string(self.row_pos, col_pos, value or '')
                elif cell_type == 'amount':
                    if col_pos == 11:
                        self.sheet.write_number(
                            self.row_pos, col_pos, 0.0, self.format_amount
                        )
                    else:
                        if col_pos == 12:
                                self.sheet.write_number(
                                    self.row_pos, col_pos, 0.0, self.format_amount
                                )
                        else:
                            self.sheet.write_number(
                                self.row_pos, col_pos, float(value), self.format_amount
                            )
                    # if account_sign == 1 and col_pos == 11:
                    #     self.sheet.write_number(
                    #         self.row_pos, col_pos, 0.0, self.format_amount
                    #     )
                    # else:
                    #     if account_sign == -1 and col_pos == 12:
                    #         self.sheet.write_number(
                    #             self.row_pos, col_pos, 0.0, self.format_amount
                    #         )
                    #     else:
                    #         if col_pos == 12:
                    #             self.sheet.write_number(
                    #                 self.row_pos, col_pos, 0.0, self.format_amount
                    #             )
                    #         else:
                    #             self.sheet.write_number(
                    #                 self.row_pos, col_pos, float(value), self.format_amount
                    #             )
                    #         # self.sheet.write_number(
                    #         #     self.row_pos, col_pos, float(value), self.format_amount
                    #         # )
                elif cell_type == 'amount_currency':
                    if my_object.currency_id:
                        format_amt = self._get_currency_amt_format(
                            my_object)
                        self.sheet.write_number(
                            self.row_pos, col_pos,
                            float(value), format_amt
                        )
            elif column.get('field_currency_balance'):
                value = getattr(my_object, column['field_currency_balance'])
                cell_type = column.get('type', 'string')
                if cell_type == 'many2one':
                    if my_object.currency_id:
                        self.sheet.write_string(
                            self.row_pos, col_pos,
                            value.name or '',
                            self.format_right
                        )
        self.row_pos += 1

    def write_ending_balance_special(self, my_object, name, label, list_summ):

        """Write a specific ending balance line on current line
        using defined columns field_final_balance name.

        Columns are defined with `_get_report_columns` method.
        """
        ending_balance_vals_summatory = list_summ[0]
        vals_read = my_object.read()[0]

        account_id = False
        if 'account_id' in vals_read:
            account_id = vals_read['account_id']
        else:
            if 'account_id' in vals_read:
                account_id = vals_read['report_account_id']
        
        account_sign = ""
        if account_id:
            account_br = self.env['account.account'].browse(account_id[0])
            if 'cuenta_tipo' in account_br._fields:
                cuenta_tipo = account_br.cuenta_tipo
                if cuenta_tipo == 'D':
                    account_sign = 1
                elif cuenta_tipo == 'A':
                    account_sign = -1
            else:
                if 'sign' in account_br._fields:
                    account_sign = account_br.sign if account_br else False

        initial_balance = vals_read.get('initial_balance',0.0)
        sum_debit = ending_balance_vals_summatory.get(11,0.0)
        sum_credit = ending_balance_vals_summatory.get(12,0.0)
        sum_cumul_balance = ending_balance_vals_summatory.get(13,0.0)

        for i in range(0, len(self.columns)):
            self.sheet.write(self.row_pos, i, '', self.format_header_right)
        row_count_name = self._get_col_count_final_balance_name()
        col_pos_label = self._get_col_pos_final_balance_label()
        self.sheet.merge_range(
            self.row_pos, 0, self.row_pos, row_count_name - 1, name,
            self.format_header_left
        )
        self.sheet.write(self.row_pos, col_pos_label, label,
                         self.format_header_right)
        for col_pos, column in self.columns.items():
            if column.get('field_final_balance'):
                value = getattr(my_object, column['field_final_balance'])
                
                cell_type = column.get('type', 'string')
                if cell_type == 'string':
                    self.sheet.write_string(self.row_pos, col_pos, value or '',
                                            self.format_header_right)
                elif cell_type == 'amount':
                    # if col_pos == 11:
                    #     self.sheet.write_number(
                    #         self.row_pos, col_pos, float(value) - initial_balance,
                    #         self.format_header_amount
                    #     )
                    if col_pos == 10:
                        initial_value = getattr(my_object, column['field_initial_balance'])
                        self.sheet.write_number(
                            self.row_pos, col_pos, initial_value if initial_value else 0.0,
                            self.format_header_amount
                        )
                    elif col_pos == 11:
                        self.sheet.write_number(
                            self.row_pos, col_pos, sum_debit if sum_debit else 0.0,
                            self.format_header_amount
                        )
                    elif col_pos == 12:
                        self.sheet.write_number(
                            self.row_pos, col_pos, sum_credit if sum_credit else 0.0,
                            self.format_header_amount
                        )
                    # if col_pos == 13:
                    #     self.sheet.write_number(
                    #         self.row_pos, col_pos, sum_cumul_balance if sum_cumul_balance else 0.0,
                    #         self.format_header_amount
                    #     )
                    else:
                        self.sheet.write_number(
                            self.row_pos, col_pos, float(value),
                            self.format_header_amount
                        )
                elif cell_type == 'amount_currency':
                    if my_object.currency_id:
                        format_amt = self._get_currency_amt_header_format(
                            my_object)
                        self.sheet.write_number(
                            self.row_pos, col_pos, float(value),
                            format_amt
                        )
            elif column.get('field_currency_balance'):
                value = getattr(my_object, column['field_currency_balance'])
                cell_type = column.get('type', 'string')
                if cell_type == 'many2one':
                    if my_object.currency_id:
                        self.sheet.write_string(
                            self.row_pos, col_pos,
                            value.name or '',
                            self.format_header_right
                        )
        self.row_pos += 1

# class AccountReportGeneralBalanceGrouped(models.Model):
#     _name = 'account.report.general.balance.grouped'
#     _description = 'Ref. Información Agrupada para el Reporte'
    
#     name = fields.Char('Diario', size=128)
#     account_id = fields.Many2one('account.account', 'Cuenta')

#     