from math import sqrt
import pandas as pd
from scipy import stats

#đọc file
honda = pd.read_excel(r"C:\Users\ACER\Documents\Phân tích dữ liệu với Python\honda_sample2.xlsx")
#Các số lượng các loại xe
xe = honda["Condition"].value_counts()
print("Các loại xe",xe)
#Danh sách xe cũ đã bán
# honda_xecu = honda[honda["Condition"] == "Used"]
# print("danh sách xe cũ", honda_xecu)
# xe_used = honda[honda["Condition"] == "Used"].count()
# print("số lượng xe cũ",xe_used)
tong_xe = len(honda)
print("Tổng số xe: ",tong_xe,"xe")

xeused = (honda["Condition"] =="Used").sum()
print("Tổng số xe cũ đã bán: ", xeused,"xe")

def uoc_tinh_ti_le(mauxe,tong,confidence_level=0.95 ):
    #tính tỉ lệ ước lượng
    ti_le_ul = xeused / tong_xe
    print("Tỉ lệ ước lượng ",ti_le_ul)
    #tính sai số tiêu chuẩn
    sai_so_tieu_chuan = sqrt(ti_le_ul * (1 - ti_le_ul) / tong_xe)
    print("sai số tiêu chuẩn",sai_so_tieu_chuan)
    #tính giá trị giới hạn tương  ứng với đôj tin cậy
    z_value = abs(stats.norm.ppf((1 + confidence_level) / 2))
    #khoảng tin cậy
    khoang_tin_cay=(ti_le_ul - z_value * sai_so_tieu_chuan, ti_le_ul + z_value * sai_so_tieu_chuan)
    return ti_le_ul, khoang_tin_cay

#ước lượng tỉ lệ và khoảng tin cậy
tileuocluong, khoangtincay = uoc_tinh_ti_le(xeused, tong_xe)
#in kết quar
print("Tỉ lệ ước lượng: {:.2%}".format(tileuocluong))
print("Khoảng tin cậy 95%: ({:.2%}, {:.2%})".format(khoangtincay[0], khoangtincay[1]))