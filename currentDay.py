import gspread
import xlwings as xw


class CurrentDate(object):
    def __init__(self, Today_google, Today_excel):
        self.Today_google = Today_google
        self.Today_excel = Today_excel
    def exec_Today(self):
        #这一段是google的API
        gc = gspread.service_account(filename='client_secret.json')
        gc.open_by_key('1zcItciXhgc85O6FK2x5b7HZ7UvmcbrN_QpEvdTb2vNY')
        worksheet_google = gc.open('BOI Production Report').get_worksheet(1)

        #这里就是用xlwings打开文件
        app = xw.App(visible= False, add_book= False)
        workbook_excel = app.books.open('D:\BOI Production Data\BOI_DAILY_REPORT_'+self.Today_excel+'.xlsx')
        worksheet_excel=workbook_excel.sheets['Sheet1']#sheets后面必须带有[]才能读取

        Galvo_build = worksheet_excel.range('C1').value
        Galvo_ship = worksheet_excel.range('C2').value

        F3OM_build = worksheet_excel.range('C3').value
        F3OM_ship = worksheet_excel.range('C4').value

        F3LOM_build = worksheet_excel.range('C5').value
        F3LOM_ship = worksheet_excel.range('C6').value

        F3Tank_build = worksheet_excel.range('C7').value
        F3Tank_ship = worksheet_excel.range('C8').value

        F3LTank_build = worksheet_excel.range('C9').value
        F3LTank_ship = worksheet_excel.range('C10').value

        F3CF_build = worksheet_excel.range('C11').value
        F3CF_ship = worksheet_excel.range('C12').value

        F3LCF_build = worksheet_excel.range('C13').value
        F3LCF_ship = worksheet_excel.range('C14').value

        wash_build = worksheet_excel.range('C15').value
        wash_ship = worksheet_excel.range('C16').value

        cure_build = worksheet_excel.range('C17').value
        cure_ship = worksheet_excel.range('C18').value

        washL_build = worksheet_excel.range('C19').value
        washL_ship = worksheet_excel.range('C20').value

        cureL_build = worksheet_excel.range('C21').value #获取excel中需要读取的值
        cureL_ship = worksheet_excel.range('C22').value #获取excel中需要读取的值

        cell_list=worksheet_google.findall(self.Today_google)

        for cell in cell_list:

            worksheet_google.update_cell(4,cell.col,Galvo_build)
            worksheet_google.update_cell(5,cell.col,Galvo_ship)

            worksheet_google.update_cell(7,cell.col,F3OM_build)
            worksheet_google.update_cell(8,cell.col,F3OM_ship)

            worksheet_google.update_cell(16,cell.col,F3LOM_build)
            worksheet_google.update_cell(17,cell.col,F3LOM_ship)

            worksheet_google.update_cell(19,cell.col,F3Tank_build)
            worksheet_google.update_cell(35,cell.col,F3Tank_ship)

            worksheet_google.update_cell(37,cell.col,F3LTank_build)
            worksheet_google.update_cell(38,cell.col,F3LTank_ship)

            worksheet_google.update_cell(40,cell.col,F3CF_build)
            worksheet_google.update_cell(41,cell.col,F3CF_ship)

            worksheet_google.update_cell(43,cell.col,F3LCF_build)
            worksheet_google.update_cell(44,cell.col,F3LCF_ship)

            worksheet_google.update_cell(46,cell.col,wash_build)
            worksheet_google.update_cell(47,cell.col,wash_ship)

            worksheet_google.update_cell(49,cell.col,cure_build)
            worksheet_google.update_cell(50,cell.col,cure_ship)

            worksheet_google.update_cell(52,cell.col,washL_build)
            worksheet_google.update_cell(53,cell.col,washL_ship)

            worksheet_google.update_cell(55,cell.col,cureL_build)#将excel所得到的数据传入，cell.col是这一列的列
            worksheet_google.update_cell(56,cell.col,cureL_ship)#将excel所得到的数据传入，cell.col是这一列的列


        workbook_excel.close()
        app.quit()
