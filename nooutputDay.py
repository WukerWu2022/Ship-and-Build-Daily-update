import gspread

class NoOutputDate(object):
    def __init__(self, GoogleDate):
        self.GoogleDate = GoogleDate
    def exec_Noouputday(self):
        # 这一段是google的API
        gc = gspread.service_account(filename='client_secret.json')
        gc.open_by_key('1zcItciXhgc85O6FK2x5b7HZ7UvmcbrN_QpEvdTb2vNY')
        worksheet_google = gc.open('BOI Production Report').get_worksheet(1)

        cell_list = worksheet_google.findall(self.GoogleDate)
        for cell in cell_list:
            worksheet_google.update_cell(4, cell.col, 0)
            worksheet_google.update_cell(5, cell.col, 0)

            worksheet_google.update_cell(7, cell.col, 0)
            worksheet_google.update_cell(8, cell.col, 0)

            worksheet_google.update_cell(16, cell.col, 0)
            worksheet_google.update_cell(17, cell.col, 0)

            worksheet_google.update_cell(19, cell.col, 0)
            worksheet_google.update_cell(35, cell.col, 0)

            worksheet_google.update_cell(37, cell.col, 0)
            worksheet_google.update_cell(38, cell.col, 0)

            worksheet_google.update_cell(40, cell.col, 0)
            worksheet_google.update_cell(41, cell.col, 0)

            worksheet_google.update_cell(43, cell.col, 0)
            worksheet_google.update_cell(44, cell.col, 0)

            worksheet_google.update_cell(46, cell.col, 0)
            worksheet_google.update_cell(47, cell.col, 0)

            worksheet_google.update_cell(49, cell.col, 0)
            worksheet_google.update_cell(50, cell.col, 0)

            worksheet_google.update_cell(52, cell.col, 0)
            worksheet_google.update_cell(53, cell.col, 0)

            worksheet_google.update_cell(55, cell.col, 0)
            worksheet_google.update_cell(56, cell.col, 0)

