from anotherDay import *
from currentDay import *
from nooutputDay import *

from datetime import datetime


class SBdailyManager(object):
    def run(self):
        while True:
            self.show_menu()
            menu_num = int(input('输入你需要的功能序号：'))
            if menu_num == 1:
                self.current_date()
            elif menu_num == 2:
                self.another_date()
            elif menu_num == 3:
                self.nooutput_date()
            elif menu_num == 4:
                break

    @staticmethod
    def show_menu():
        print('请选择以下功能：')
        print('1: 日期为当天')
        print('2: 日期不是当天')
        print('3: 某天没生产')
        print('4: 结束按 4 退出系统')

    def current_date(self):
        today = datetime.now()
        Today_google = today.strftime('%Y-%m-%d')#匹配google日期如2021-10-10的格式
        Today_excel = today.strftime('%Y%m%d')#匹配excel日期如20211010的格式
        current_day = CurrentDate(Today_google, Today_excel)
        current_day.exec_Today()

    def another_date(self):
        ExcelDate = input('填写年月日，格式为YYYYMMDD:')
        GoogleDate = input('填写年月日，格式为YYYY-MM-DD:')
        current_day = AnotherDate(ExcelDate, GoogleDate)
        current_day.exec_Anotherday()

    def nooutput_date(self):
        GoogleDate = input('填写年月日，格式为YYYY-MM-DD:')
        nooutputdate = NoOutputDate(GoogleDate)
        nooutputdate.exec_Noouputday()
