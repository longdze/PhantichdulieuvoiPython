import pandas as pd
import openpyxl
from openpyxl import load_workbook
#đọc file
doc = pd.read_excel(r"C:\Users\ACER\Downloads\Temperature.xlsx")
print(doc)

# df = pd.DataFrame(doc)

#tải file excel
workbook = load_workbook(filename =r"C:\Users\ACER\Downloads\Temperature.xlsx")
#chọn sheet mặc định
sheet = workbook.active
#lấy giá trị của ô A1
print("giá trị của ô A1")
value = sheet['A1'].value
print(value)
#lấy giá trị của oo bất kì
print("lấy giá trị của ô bất kì")
value2 = sheet.cell(row=3, column=2).value
print(value)

print("lấy các ô từ hàng 1 đến hàng 10, cột 1 đến cột 3")
#đọc nhiều ô cùng một lúc bằng cách duyệt qua nhiều ô
for row in sheet.iter_rows(min_row=1, min_col=1, max_row=10, max_col=3, values_only=True):
    for cell in row:
        print(cell)

#dùng iter_row để duyệt các ô hàng 1 đén 10, cột 1 đến 3
#value_only để lây giá trị của từng ô



#SẮP XẾP THEO THUỘC TÍNH
#sắp xếp theo tăng dần của cột Temperature
doc_sorted = doc.sort_values(by=['Temperature'])
print('dữ liệu đã sắp xếp')
print(doc_sorted)
#hoặc sắp xếp theo thứ tự giảm dần
doc_sorted2 = doc.sort_values(by=['Temperature'], ascending=False)
print("du lieu sau khi sắp xếp")
print(doc_sorted2)
#tham số ascending để sắp xếp theo thứ tự giảm dần

#SẮP XẾP NHIỀU THUỘC TÍNH: Temperature theo thứ tự giảm dần và Salinity theo thứ tự tăng dần
doc_sorte3 = doc.sort_values(by=['Temperature','Salinity'],ascending=[False,True])
print('Dữ liệu sau khi sắp xếp')
print(doc_sorte3)

#lấy danh sách cách giá trị trong cột Season
#chưa lọc giá trị, danh sách in ra sẽ bị trùng lặp
names = doc['Season'].tolist()  #tolist chuyển đổi 1 cột của object thành 1 list trong python
print("in ra danh sách các giá trị trong cột Season",names)

#lọc các giá trị và in ra danh sách các giá trị không bị trùng lặp
loc = doc['Season'].unique()
for value in loc:
    print(value)

#tính max, min, tổng, trung bình,...
max = doc['Temperature'].max()
min = doc['Temperature'].min()
print("giá trị lớn nhất:",max, "giá trị nhỏ nhất:",min)
tong = doc['Temperature'].sum()
print("Tổng là:",tong)
trung_binh = doc['Temperature'].mean()
print("giá trị trung bình của cột Temperature là: ",trung_binh)
