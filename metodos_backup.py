

    @api.multi
    def print_excel_report(self):
        
        ######################### CON XLWT ############################

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
        # print ("\n\n\n\n\nreport_lines",report_lines)

        report_name = data['form']['account_report_id'][1]
        total_left, total_left_cmp, total_right, total_right_cmp = 0.00, 0.00, 0.00, 0.00

        Style = ExcelStyles()
        wbk = xlwt.Workbook()
        sheet1 = wbk.add_sheet(report_name)
        sheet1.set_panes_frozen(True)
        sheet1.set_horz_split_pos(6)
        sheet1.show_grid = True
        sheet1.col(0).width = 11000
        sheet1.col(1).width = 5000
        sheet1.col(2).width = 5000
        sheet1.col(3).width = 5000
        sheet1.col(4).width = 1500
        sheet1.col(5).width = 4000
        sheet1.col(6).width = 4000
        sheet1.col(7).width = 4000
        sheet1.col(8).width = 4000
        sheet1.col(9).width = 1000
        sheet1.col(10).width = 4000
        sheet1.col(11).width = 1000
        sheet1.col(12).width = 4000
        sheet1.col(13).width = 4000
        sheet1.col(14).width = 4000
        sheet1.col(15).width = 4000
        r1 = 10
        r2 = 11
        r3 = 12
        r4 = 13
        r5 = 14
        sheet1.row(r1).height = 600
        sheet1.row(r2).height = 600
        sheet1.row(r3).height = 350
        sheet1.row(r4).height = 350
        sheet1.row(r5).height = 256

        title = report_name
        row = r5
        right_row = r5


        if data['form']['debit_credit'] == True:

            sheet1.write(r3, 0, "Target Move", Style.subTitle())
            if data['form']['target_move'] == 'all':
                sheet1.write(r4, 0, "All Entries", Style.subTitle())
            if data['form']['target_move'] == 'posted':
                sheet1.write(r4, 0, "All Posted Entries", Style.subTitle())
            date_from = date_to = False
            if data['form']['date_from']:
    #            date_from = datetime.strptime(data['form']['date_from'], "%Y-%m-%d")
    #             date_from = datetime.strftime(data['form']['date_from'], "%d-%m-%Y")
                sheet1.write(r3, 1, "Date From", Style.subTitle())
                sheet1.write(r4, 1, data['form']['date_from'], Style.normal_date_alone())
            else:
                sheet1.write(r3, 1, "", Style.subTitle())
                sheet1.write(r4, 1, "", Style.subTitle())
            if data['form']['date_to']:
    #            date_to = datetime.strptime(data['form']['date_to'], "%Y-%m-%d")
    #             date_to = datetime.strftime(data['form']['date_to'], "%d-%m-%Y")
                sheet1.write(r3, 2, "Date To", Style.subTitle())
                sheet1.write(r4, 2, data['form']['date_to'], Style.normal_date_alone())
            else:
                sheet1.write(r3, 2, "", Style.subTitle())
                sheet1.write(r4, 2, "", Style.subTitle())
            if data['form'].get('right') == True:
                # print ("inside right==========================================")
                sheet1.write(r3, 4, "Target Move", Style.subTitle())
                if data['form']['target_move'] == 'all':
                    sheet1.write(r4, 4, "All Entries", Style.subTitle())
                if data['form']['target_move'] == 'posted':
                    sheet1.write(r4, 4, "All Posted Entries", Style.subTitle())
    #            date_from = date_to = False
                if data['form']['date_from']:
    #                date_from = datetime.strptime(data['form']['date_from'], "%Y-%m-%d")
    #                date_from = datetime.strftime(date_from, "%d-%m-%Y")
                    sheet1.write(r3, 5, "Date From", Style.subTitle())
                    sheet1.write(r4, 5, data['form']['date_from'], Style.normal_date_alone())
                else:
                    sheet1.write(r3, 5, "", Style.subTitle())
                    sheet1.write(r4, 5, "", Style.subTitle())
                if data['form']['date_to']:
    #                date_to = datetime.strptime(data['form']['date_to'], "%Y-%m-%d")
    #                 date_to = datetime.strftime(data['form']['date_to'], "%d-%m-%Y")
                    sheet1.write(r3, 6, "Date To", Style.subTitle())
                    sheet1.write(r4, 6, data['form']['date_to'], Style.normal_date_alone())
                else:
                    sheet1.write(r3, 6, "", Style.subTitle())
                    sheet1.write(r4, 6, "", Style.subTitle())


            sheet1.write_merge(r1, r1, 0, 3, self.env.user.company_id.name, Style.title_color())
            sheet1.write_merge(r2, r2, 0, 3, title, Style.sub_title_color())
            sheet1.write(r3, 3, "", Style.subTitle())
            sheet1.write(r4, 3, "", Style.subTitle())
            row = row + 1
            right_row +=1
            sheet1.row(row).height = 256 * 3
            sheet1.write(row, 0, "Account", Style.subTitle_color())
            sheet1.write(row, 1, "Debit", Style.subTitle_color())
            sheet1.write(row, 2, "Credit", Style.subTitle_color())
            sheet1.write(row, 3, "Balance", Style.subTitle_color())
#            sheet1.write(row, 4, "", Style.subTitle_color())
            if data['form'].get('right'):
                sheet1.write_merge(r1, r1, 4, 7, self.env.user.company_id.name, Style.title_color())
                sheet1.write_merge(r2, r2, 4, 7, title, Style.sub_title_color())
                sheet1.write(r3, 7, "", Style.subTitle())
                sheet1.write(r4, 7, "", Style.subTitle())
                sheet1.write(row, 4, "Account", Style.subTitle_color())
                sheet1.write(row, 5, "Debit", Style.subTitle_color())
                sheet1.write(row, 6, "Credit", Style.subTitle_color())
                sheet1.write(row, 7, "Balance", Style.subTitle_color())
            for each in report_lines:
                print ("#### LEVEL >>>>>>>> ", each['level'])
                print ("#### LEVEL 0 ES CABECERA PRINCIPAL >>>>>>>> ", each['level'])
                print ("#### LEVEL 1 ES AGRUPACIÓN DE LINEAS - ACTIVO - PASIVO - CAPITAL >>>>>>>> ", each['level'])
                print ("#### row >>>>>>> ",row)
                print ("#### right_row >>>>>>> ",right_row)
                if each['level'] != 0:
                    name = ""
                    gap = " "
                    name = each['name']
                    left = Style.normal_left()
                    right = Style.normal_num_right_3separator()
                    if each.get('level') > 3:
                        gap = " " * (each['level'] * 5)
                        left = Style.normal_left()
                        right = Style.normal_num_right_3separator()
                    if not each.get('level') > 3:
                        gap = " " * (each['level'] * 5)
                        left = Style.subTitle_sub_color_left()
                        right = Style.subTitle_float_sub_color()
                    if each.get('level') == 1:
                        gap = " " * each['level']
                    if each.get('account_type') == 'view':
                        left = Style.subTitle_sub_color_left()
                        right = Style.subTitle_float_sub_color()
                    if each['report_side'] != 'right':
                        row = row + 1
                        sheet1.row(row).height = 400
                        name = gap + name
                        sheet1.write(row, 0, name, left)
                        sheet1.write(row, 1, each['debit'], right)
                        sheet1.write(row, 2, each['credit'], right)
                        sheet1.write(row, 3, each['balance'], right)
#                        sheet1.write(row, 4, self.env.user.company_id.currency_id.symbol, left)
                        if each['level'] == 1:
                            total_left += each['balance']
                    elif each['report_side'] == 'right':
                        sheet1.col(4).width = 11000
                        sheet1.col(5).width = 5000
                        sheet1.col(6).width = 5000
                        sheet1.col(7).width = 5000
                        right_row = right_row + 1
                        sheet1.row(right_row).height = 400
                        name = gap + name
                        sheet1.write(right_row, 4, name, left)
                        sheet1.write(right_row, 5, each['debit'], right)
                        sheet1.write(right_row, 6, each['credit'], right)
                        sheet1.write(right_row, 7, each['balance'], right)
                        if each['level'] == 1:
                            total_right += each['balance']
            if data['form'].get('right'):
                if right_row > row:
                    sheet1.write(right_row+1, 0,  'Total', Style.groupByTotalNocolor())
                    sheet1.write_merge(right_row+1, right_row+1, 1, 3, total_left, Style.groupByTotalNocolor())
                    sheet1.write(right_row+1, 4, 'Total', Style.groupByTotalNocolor())
                    sheet1.write_merge(right_row+1, right_row+1, 5, 7, total_right, Style.groupByTotalNocolor())
                else:
                    sheet1.write(row+1, 0, 'Total', Style.groupByTotalNocolor())
                    sheet1.write_merge(row+1, row+1, 1, 3, total_left, Style.groupByTotalNocolor())
                    sheet1.write(row+1, 4, 'Total', Style.groupByTotalNocolor())
                    sheet1.write_merge(row+1, row+1, 5, 7, total_right, Style.groupByTotalNocolor())
            else:
                sheet1.write(row+1, 0, 'Total', Style.groupByTotalNocolor())
                sheet1.write_merge(row+1, row+1, 1, 3, total_left, Style.groupByTotalNocolor())

        if not data['form']['enable_filter'] and not data['form']['debit_credit']:
            sheet1.write(r3, 0, "Target Move", Style.subTitle())
            if data['form']['target_move'] == 'all':
                sheet1.write(r4, 0, "All Entries", Style.subTitle())
            if data['form']['target_move'] == 'posted':
                sheet1.write(r4, 0, "All Posted Entries", Style.subTitle())
            date_from = date_to = False
            if data['form']['date_from']:
    #            date_from = datetime.strptime(data['form']['date_from'], "%Y-%m-%d")
    #             date_from = datetime.strftime(data['form']['date_from'], "%d-%m-%Y")
                sheet1.write(r3, 1, ("From" + " - "+ data['form']['date_from']), Style.subTitle())
#                sheet1.write(r4, 1, date_from, Style.normal_date_alone())
            else:
                sheet1.write(r3, 1, "", Style.subTitle())
#                sheet1.write(r4, 1, "", Style.subTitle())
            if data['form']['date_to']:
    #            date_to = datetime.strptime(data['form']['date_to'], "%Y-%m-%d")
    #             date_to = datetime.strftime(data['form']['date_to'], "%d-%m-%Y")
                sheet1.write(r4, 1, "To" + " - "+ data['form']['date_to'], Style.subTitle())
#                sheet1.write(r4, 2, date_to, Style.normal_date_alone())
            else:
#                sheet1.write(r3, 2, "", Style.subTitle())
                sheet1.write(r4, 1, "", Style.subTitle())
            if data['form'].get('right') == True:
                # print ("inside right==========================================")
                sheet1.write(r3, 2, "Target Move", Style.subTitle())
                if data['form']['target_move'] == 'all':
                    sheet1.write(r4, 2, "All Entries", Style.subTitle())
                if data['form']['target_move'] == 'posted':
                    sheet1.write(r4, 2, "All Posted Entries", Style.subTitle())
    #            date_from = date_to = False
                if data['form']['date_from']:
    #                date_from = datetime.strptime(data['form']['date_from'], "%Y-%m-%d")
    #                date_from = datetime.strftime(date_from, "%d-%m-%Y")
#                    sheet1.write(r3, 4, "Date From", Style.subTitle())
                    sheet1.write(r3, 3, ("From" + " - "+ data['form']['date_from']), Style.subTitle())
#                    sheet1.write(r4, 4, date_from, Style.normal_date_alone())
                else:
                    sheet1.write(r3, 3, "", Style.subTitle())
#                    sheet1.write(r4, 4, "", Style.subTitle())
                if data['form']['date_to']:
    #                date_to = datetime.strptime(data['form']['date_to'], "%Y-%m-%d")
    #                 date_to = datetime.strftime(data['form']['date_to'], "%d-%m-%Y")
#                    sheet1.write(r3, 5, "Date To", Style.subTitle())
                    sheet1.write(r4, 3, ("To" + " - "+ data['form']['date_to']), Style.subTitle())
#                    sheet1.write(r4, 5, date_to, Style.normal_date_alone())
                else:
#                    sheet1.write(r3, 5, "", Style.subTitle())
                    sheet1.write(r4, 3, "", Style.subTitle())
            sheet1.write_merge(r1, r1, 0, 1, self.env.user.company_id.name, Style.title_color())
            sheet1.write_merge(r2, r2, 0, 1, title, Style.sub_title_color())
            if data['form'].get('right'):
                sheet1.write_merge(r1, r1, 2, 3, self.env.user.company_id.name, Style.title_color())
                sheet1.write_merge(r2, r2, 2, 3, title, Style.sub_title_color())
            row = row + 1
            right_row = right_row + 1

            sheet1.row(row).height = 256 * 3
            sheet1.write(row, 0, "Name", Style.subTitle_color())
            sheet1.write(row, 1, "Balance", Style.subTitle_color())
#            sheet1.write(row, 2, "", Style.subTitle_color())
            if data['form'].get('right'):
                sheet1.write(row, 2, "Name", Style.subTitle_color())
                sheet1.write(row, 3, "Balance", Style.subTitle_color())
#                sheet1.write(row, 5, "", Style.subTitle_color())
            print ("###  0 es A  >>>>")
            print ("###  1 es B y ASI SUCESIVAMENTE  >>>>")

            ################## ---- INICIO DE LA PRIMER PRUEBA ---- #################
            #### Saltos de Columna ########
            name_1 = 0
            balance_1 = 1
            name_2 = 2
            balance_2 = 3
            count_init_line = 0
            save_row = row
            save_right_row = right_row
            cum_row = row
            cum_right_row = right_row
            last_row = 0
            last_right_row = 0
            sum_row = 0
            sum_right_row = 0
            print ("###### INIT row >>>>>>>>>>> ", row)
            print ("###### INIT right_row >>>>>>>>>>> ", right_row)
            #### Fin de Saltos de C.#######
            for each in report_lines:
                #### Saltos de Columna ######## 
                each_level = each['level']
                #### Fin de Saltos de C.#######
                print ("#### 01 LEVEL >>>>>>>> ", each['level'])
                # print ("#### 01 LEVEL 0 ES CABECERA PRINCIPAL >>>>>>>> ", each['level'])
                # print ("#### 01 LEVEL 1 ES AGRUPACIÓN DE LINEAS - ACTIVO - PASIVO - CAPITAL >>>>>>>> ", each['level'])
                # print ("#### 01 row >>>>>>> ",row)
                # print ("#### 01 right_row >>>>>>> ",right_row)
                # print ("#### 01 each >>>>>>> ",each)

                if each['level'] != 0:
                    # #### Saltos de Columna ######## 
                    if each_level == 1:
                        print ("### NIVEL 1 >>>> ")
                        print ("### count_init_line >>>> ", count_init_line)
                        if count_init_line > 0:
                            name_1 += 3
                            balance_1 += 3
                            name_2 += 3
                            balance_2 += 3
                            
                    # #### Fin de Saltos de C.#######

                    name = ""
                    gap = " "
                    name = each['name']
                    print ("#---- NAME >>>> ", name)
                    left = Style.normal_left()
                    right = Style.normal_num_right_3separator()
                    if each.get('level') > 3:
                        gap = " " * (each['level'] * 5)
                        left = Style.normal_left()
                        right = Style.normal_num_right_3separator()
                    if not each.get('level') > 3:
                        gap = " " * (each['level'] * 5)
                        left = Style.subTitle_sub_color_left()
                        right = Style.subTitle_float_sub_color()
                    if each.get('level') == 1:
                        gap = " " * each['level']
                    if each.get('account_type') == 'view':
                        left = Style.subTitle_sub_color_left()
                        right = Style.subTitle_float_sub_color()
                    if each['report_side'] != 'right':
                        #### Saltos de Columna ########
                        row = row + 1
                        cum_row += 1
                        #### Fin de Saltos de C.#######
                        sheet1.row(row).height = 400
                        name = gap + name
                        #### Saltos de Columna ######## 
                        print ("### name_1 >>> ",name_1)
                        print ("### balance_1 >>> ",balance_1)
                        if each_level == 1:
                            if count_init_line > 0:
                                sheet1.write(save_row+1, name_1, name, left)
                                sheet1.write(save_right_row+1, balance_1, each['balance'], right)
                            else:
                                sheet1.write(row, name_1, name, left)
                                sheet1.write(row, balance_1, each['balance'], right)
                        else:
                            sheet1.write(row, name_1, name, left)
                            sheet1.write(row, balance_1, each['balance'], right)
                        #### Fin de Saltos de C.#######
#                        sheet1.write(row, 2, self.env.user.company_id.currency_id.symbol, left)
                        if each['level'] == 1:
                            total_left += each['balance']
                    elif each['report_side'] == 'right':
                        sheet1.col(2).width = 11000
                        sheet1.col(3).width = 5000
#                        sheet1.col(5).width = 5000
                        #### Saltos de Columna ########
                        right_row = right_row + 1
                        cum_right_row += 1
                        #### Fin de Saltos de C.#######
                        sheet1.row(right_row).height = 400
                        name = gap + name
                        #### Saltos de Columna ######## 
                        sheet1.write(right_row, name_2, name, left)
                        sheet1.write(right_row, balance_2, each['balance'], right)
                        #### Fin de Saltos de C.#######
#                        sheet1.write(right_row, 5, self.env.user.company_id.currency_id.symbol, left)
                        if each['level'] == 1:
                            total_right += each['balance']
                    #### Saltos de Columna ######## 
                    if each_level == 1:
                        count_init_line += 1
                        print ("#:::: Before row >>>>>>>> ", row)
                        print ("#:::: Before right_row >>>>>>>> ", right_row)
                        if count_init_line > 1:
                            # name_1 += 3
                            # balance_1 += 3
                            # name_2 += 3
                            # balance_2 += 3
                            print ("#:::: row >>>>>>>> ", row)
                            print ("#:::: right_row >>>>>>>> ", right_row)
                            if row > last_row:
                                print ("### SI >>>>>> ")
                                sum_row = row+1
                                sum_right_row = last_row+1
                            else:
                                sum_row = last_row+1
                                sum_right_row = last_right_row+1
                            row = save_row+1
                            right_row = save_right_row+1
                        else:
                            last_row = row
                            last_right_row = right_row
                            sum_row = last_row
                            sum_right_row = last_right_row
                        #     save_row = row
                        #     save_right_row = right_row
                    #### Fin de Saltos de C.#######
            #### Saltos de Columna ######## 
            print ("### sum_row >>>>>> ",sum_row)
            print ("### sum_right_row >>>>>> ",sum_right_row)
            print ("### cum_row >>>>>> ",cum_row)
            print ("### cum_right_row >>>>>> ",cum_right_row)
            if data['form'].get('right'):
                if sum_right_row > sum_row:
                    sheet1.write(sum_right_row+1, 0,  'Total', Style.groupByTotalNocolor())
                    sheet1.write_merge(sum_right_row+1, sum_right_row+1, 1, 1, total_left, Style.groupByTotalNocolor())
                    sheet1.write(sum_right_row+1, 2, 'Total', Style.groupByTotalNocolor())
                    sheet1.write_merge(sum_right_row+1, sum_right_row+1, 3, 3, total_right, Style.groupByTotalNocolor())
                else:
                    sheet1.write(sum_row+1, 0, 'Total', Style.groupByTotalNocolor())
                    sheet1.write_merge(sum_row+1, sum_row+1, 1, 1, total_left, Style.groupByTotalNocolor())
                    sheet1.write(sum_row+1, 2, 'Total', Style.groupByTotalNocolor())
                    sheet1.write_merge(sum_row+1, sum_row+1, 3, 3, total_right, Style.groupByTotalNocolor())
            else:
                sheet1.write(sum_row+1, 0, 'Total', Style.groupByTotalNocolor())
                sheet1.write_merge(sum_row+1, sum_row+1, 1, 1, total_left, Style.groupByTotalNocolor())
            #### Fin de Saltos de C.#######

        ################## ---- FIN DE LA PRIMER PRUEBA ---- #################

        if data['form']['enable_filter'] and not data['form']['debit_credit']:
            sheet1.write(r3, 0, "Target Move", Style.subTitle())
            if data['form']['target_move'] == 'all':
                sheet1.write(r4, 0, "All Entries", Style.subTitle())
            if data['form']['target_move'] == 'posted':
                sheet1.write(r4, 0, "All Posted Entries", Style.subTitle())
            date_from = date_to = False
            if data['form']['date_from']:
    #            date_from = datetime.strptime(data['form']['date_from'], "%Y-%m-%d")
    #             date_from = datetime.strftime(data['form']['date_from'], "%d-%m-%Y")
                sheet1.write(r3, 1, "Date From", Style.subTitle())
                sheet1.write(r4, 1, data['form']['date_from'], Style.normal_date_alone())
            else:
                sheet1.write(r3, 1, "", Style.subTitle())
                sheet1.write(r4, 1, "", Style.subTitle())
            if data['form']['date_to']:
    #            date_to = datetime.strptime(data['form']['date_to'], "%Y-%m-%d")
    #             date_to = datetime.strftime(data['form']['date_to'], "%d-%m-%Y")
                sheet1.write(r3, 2, "Date To", Style.subTitle())
                sheet1.write(r4, 2, data['form']['date_to'], Style.normal_date_alone())
            else:
                sheet1.write(r3, 2, "", Style.subTitle())
                sheet1.write(r4, 2, "", Style.subTitle())
            if data['form'].get('right') == True:
                # print ("inside right==========================================")
                sheet1.write(r3, 3, "Target Move", Style.subTitle())
                if data['form']['target_move'] == 'all':
                    sheet1.write(r4, 3, "All Entries", Style.subTitle())
                if data['form']['target_move'] == 'posted':
                    sheet1.write(r4, 3, "All Posted Entries", Style.subTitle())
    #            date_from = date_to = False
                if data['form']['date_from']:
    #                date_from = datetime.strptime(data['form']['date_from'], "%Y-%m-%d")
    #                date_from = datetime.strftime(date_from, "%d-%m-%Y")
                    sheet1.write(r3, 4, "Date From", Style.subTitle())
                    sheet1.write(r4, 4, data['form']['date_from'], Style.normal_date_alone())
                else:
                    sheet1.write(r3, 4, "", Style.subTitle())
                    sheet1.write(r4, 4, "", Style.subTitle())
                if data['form']['date_to']:
    #                date_to = datetime.strptime(data['form']['date_to'], "%Y-%m-%d")
    #                 date_to = datetime.strftime(data['form']['date_to'], "%d-%m-%Y")
                    sheet1.write(r3, 5, "Date To", Style.subTitle())
                    sheet1.write(r4, 5, data['form']['date_to'], Style.normal_date_alone())
                else:
                    sheet1.write(r3, 5, "", Style.subTitle())
                    sheet1.write(r4, 5, "", Style.subTitle())


            sheet1.write_merge(r1, r1, 0, 2, self.env.user.company_id.name, Style.title_color())
            sheet1.write_merge(r2, r2, 0, 2, title, Style.sub_title_color())
            if data['form'].get('right') == True:
                sheet1.write_merge(r1, r1, 3, 5, self.env.user.company_id.name, Style.title_color())
                sheet1.write_merge(r2, r2, 3, 5, title, Style.sub_title_color())
#            sheet1.write(r3, 3, "", Style.subTitle())
#            sheet1.write(r4, 3, "", Style.subTitle())
            row = row + 1
            right_row += 1
            sheet1.row(row).height = 256 * 3
            sheet1.write(row, 0, "Name", Style.subTitle_color())
            sheet1.write(row, 1, "Balance", Style.subTitle_color())
#            sheet1.write(row, 2, "", Style.subTitle_color())
            sheet1.write(row, 2, data['form']['label_filter'], Style.subTitle_color())
            if data['form'].get('right'):
                sheet1.col(3).width = 11000
                sheet1.col(4).width = 5000
                sheet1.col(5).width = 5000
                sheet1.write(row, 3, "Name", Style.subTitle_color())
                sheet1.write(row, 4, "Balance", Style.subTitle_color())
                sheet1.write(row, 5, data['form']['label_filter'], Style.subTitle_color())
            for each in report_lines:
                print ("#### 02 LEVEL >>>>>>>> ", each['level'])
                print ("#### 02 LEVEL 0 ES CABECERA PRINCIPAL >>>>>>>> ", each['level'])
                print ("#### 02 LEVEL 1 ES AGRUPACIÓN DE LINEAS - ACTIVO - PASIVO - CAPITAL >>>>>>>> ", each['level'])
                print ("#### 02 row >>>>>>> ",row)
                print ("#### 02 right_row >>>>>>> ",right_row)
                if each['level'] != 0:
                    name = ""
                    gap = " "
                    name = each['name']
                    left = Style.normal_left()
                    right = Style.normal_num_right_3separator()
                    if each.get('level') > 3:
                        gap = " " * (each['level'] * 5)
                        left = Style.normal_left()
                        right = Style.normal_num_right_3separator()
                    if not each.get('level') > 3:
                        gap = " " * (each['level'] * 5)
                        left = Style.subTitle_sub_color_left()
                        right = Style.subTitle_float_sub_color()
                    if each.get('level') == 1:
                        gap = " " * each['level']
                    if each.get('account_type') == 'view':
                        left = Style.subTitle_sub_color_left()
                        right = Style.subTitle_float_sub_color()
                    if each['report_side'] != 'right':
                        row = row + 1
                        sheet1.row(row).height = 400
                        name = gap + name
                        sheet1.write(row, 0, name, left)
                        sheet1.write(row, 1, each['balance'], right)
    #                    sheet1.write(row, 2, self.env.user.company_id.currency_id.symbol, left)
                        sheet1.write(row, 2, each['balance_cmp'], right)
                        if each['level'] == 1:
                            total_left += each['balance']
                            total_left_cmp += each['balance_cmp']
                    elif each['report_side'] == 'right':
                        right_row = right_row + 1
                        sheet1.row(right_row).height = 400
                        name = gap + name
                        sheet1.write(right_row, 3, name, left)
                        sheet1.write(right_row, 4, each['balance'], right)
    #                    sheet1.write(row, 2, self.env.user.company_id.currency_id.symbol, left)
                        sheet1.write(right_row, 5, each['balance_cmp'], right)
                        if each['level'] == 1:
                            total_right += each['balance']
                            total_right_cmp += each['balance_cmp']
            if data['form'].get('right'):
                if right_row > row:
                    sheet1.write(right_row+1, 0,  'Total', Style.groupByTotalNocolor())
                    sheet1.write_merge(right_row+1, right_row+1, 1, 1, total_left, Style.groupByTotalNocolor())
                    sheet1.write_merge(right_row+1, right_row+1, 2, 2, total_left_cmp, Style.groupByTotalNocolor())
                    sheet1.write(right_row+1, 3, 'Total', Style.groupByTotalNocolor())
                    sheet1.write_merge(right_row+1, right_row+1, 4, 4, total_right, Style.groupByTotalNocolor())
                    sheet1.write_merge(right_row+1, right_row+1, 5, 5, total_right_cmp, Style.groupByTotalNocolor())
                else:
                    sheet1.write(row+1, 0, 'Total', Style.groupByTotalNocolor())
                    sheet1.write_merge(row+1, row+1, 1, 1, total_left, Style.groupByTotalNocolor())
                    sheet1.write_merge(row+1, row+1, 2, 2, total_left_cmp, Style.groupByTotalNocolor())
                    sheet1.write(row+1, 3, 'Total', Style.groupByTotalNocolor())
                    sheet1.write_merge(row+1, row+1, 4, 4, total_right, Style.groupByTotalNocolor())
                    sheet1.write_merge(row+1, row+1, 5, 5, total_right_cmp, Style.groupByTotalNocolor())
            else:
                sheet1.write(row+1, 0, 'Total', Style.groupByTotalNocolor())
                sheet1.write_merge(row+1, row+1, 1, 1, total_left, Style.groupByTotalNocolor())
                sheet1.write_merge(row+1, row+1, 2, 2, total_left_cmp, Style.groupByTotalNocolor())

        stream = io.BytesIO()
        wbk.save(stream)
        self.env.cr.execute(""" DELETE FROM accounting_report_output""")
#        self.write({'name': report_name + '.xls', 'output': base64.encodestring(stream.getvalue())})
#        return {
#                'name': _('Notification'),
#                'view_type': 'form',
#                'view_mode': 'form',
#                'res_model': 'accounting.report',
#                'res_id': self.id,
#                'type': 'ir.actions.act_window',
#                'target': 'new'
#                }
        attach_id = self.env['accounting.report.output'].create({'name': report_name + '.xls', 'output': base64.encodestring(stream.getvalue())})
        return {
                'name': _('Notification'),
                'view_type': 'form',
                'view_mode': 'form',
                'res_model': 'accounting.report.output',
                'res_id': attach_id.id,
                'type': 'ir.actions.act_window',
                'target': 'new'
                }
