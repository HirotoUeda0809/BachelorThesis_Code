#########################################################################
#作成日: 2019/9/27 作成者: Hiroto Ueda(hrt.ueda0809@gmail.com)
#メモ: これはシフト希望を入力してもらうために配るシートで，
    #全員のシフト希望を集約したら分析にかけるデータファイルとなる．
#########################################################################


#ライブラリ読み取りパート
import openpyxl as px
import RAP_OJT as rp
import DE_OJT as de
import datetime
path = rp.path()


#パラメーターの読み込みパート
#スタッフ系の情報の読み込み
A = rp.AllMembersList()
ALLNAMELIST = rp.AllMembersNameList()
AContractW = rp.StaffContract_Week()
#日付系の情報の読み取り
MONTH = rp.Month()
DAYS = rp.Days()
SATURDAYS = rp.St_Sn_Hdays("St")
SUNDAYS = rp.St_Sn_Hdays("Sn")
HOLIDAYS = rp.St_Sn_Hdays("H")


#ファイルの作成と保存パート
wb = px.Workbook()
ShiftSheet = wb.create_sheet("HopeShiftSheet", 0)

de.draw_line(ShiftSheet, A, DAYS, "M")
de.write_and_color(ShiftSheet, A, ALLNAMELIST, DAYS, MONTH, SATURDAYS, HOLIDAYS,"M")
de.addjust_width(ShiftSheet)

wb.save("{:}\\{:}月希望.xlsx".format(path, MONTH))