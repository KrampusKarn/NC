import netCDF4 
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

window = tk.Tk()
window.title('ดึงข้อมูลGCM BY KobEbaN WE 37')
# สร้างฟังก์ชันสำหรับเลือกไฟล์
window.geometry("500x300")
def select_folder_and_save_file():
    name_gcm = list(pd.read_excel('C:/Users/cattr/Desktop/cmip6/สรุป.xlsx')['ชื่อแบบบจำลอง'])
    ssp = ['historical','ssp126','ssp245','ssp370','ssp585']
    type_list = ['pr','tas','tasmax','tasmin']
    save_list = []
    for i in range(len(name_gcm)):
        save_list.append([])
        for ii in range(len(ssp)):
            save_list[i].append([])
            for iii in range(len(type_list)):
                save_list[i][ii].append([])
                if ii == 0:
                    for iiii in range((2015-2004) * 12):
                        save_list[i][ii][iii].append(0)
                else:
                    for iiii in range((2100-2014) * 12):
                        save_list[i][ii][iii].append(0)
    uiiuuu = 0
    folder_path = 'C:/Users/cattr/Desktop/cmip6'
    for root, dirs, files in os.walk(folder_path):
        for file_name in files:
            if file_name[-2:] == '.nc':
                data_load = netCDF4.Dataset(file_name)
                uiiuuu += 1
                print(uiiuuu)
                file_name_list = str(file_name).split('_')
                var_data = file_name_list[0]
                Gcm_name = file_name_list[2]
                ssp_name = file_name_list[3]
                y_s = int(file_name_list[6][:4])
                y_e = int(file_name_list[6][7:11])
                if y_e > 2003 and name_gcm.count(Gcm_name) != 0 and ssp.count(ssp_name) != 0 and type_list.count(var_data) != 0 and y_e < 2101:
                    data = data_load.variables[var_data][:]
                    lon = data_load.variables['lon'][:]
                    lat = data_load.variables['lat'][:]
                    for i in range(len(lon)):
                        if lon[0] > lon[1]:
                            if lon[i] > 98.9755:
                                lon_1 = i
                                lon_2 = i+1
                        else:
                            if lon[i] < 98.9755:
                                lon_1 = i
                                lon_2 = i +1
                    for i in range(len(lat)):
                        if lat[0] > lat[1]:
                            if lat[i] > 19.2686:
                                lat_1 = i
                                lat_2 = i+1
                        else:
                            if lat[i] < 19.2686:
                                lat_1 = i
                                lat_2 = i+1
                    dis_11 = 1 / ((lat[lat_1] - 19.286) ** 2 + (lon[lon_1] - 98.9755) ** 2)
                    dis_12 = 1 / ((lat[lat_1] - 19.286) ** 2 + (lon[lon_2] - 98.9755) ** 2)
                    dis_21 = 1 / ((lat[lat_2] - 19.286) ** 2 + (lon[lon_1] - 98.9755) ** 2)
                    dis_22 = 1 / ((lat[lat_2] - 19.286) ** 2 + (lon[lon_2] - 98.9755) ** 2)
                    for i in range(len(data)):
                        if ssp_name == ssp[0]:
                            gh = (y_s - 2004) * 12 + i
                        else:
                            gh = (y_s - 2015) * 12 + i
                        if gh > -1:
                            pr_real = ((data[i][lat_1][lon_1])*dis_11 + (data[i][lat_1][lon_2])*dis_12 + (data[i][lat_2][lon_1])*dis_21 +(data[i][lat_2][lon_2])*dis_22)/(dis_22+dis_11+dis_21+dis_12)
                            z = name_gcm.index(Gcm_name)
                            x = ssp.index(ssp_name)
                            c = type_list.index(var_data)
                            save_list[z][x][c][gh] = pr_real
    save = pd.DataFrame()
    writer = pd.ExcelWriter('C:/Users/cattr/Desktop/cmip6/GCMรอบสุดท้ายสุด.xlsx')
    for ii in range(len(save_list[0])):
        for iii in range(len(save_list[0][0])):
            save = pd.DataFrame()
            for i in range(len(save_list)):
                save[name_gcm[i]] = pd.Series(save_list[i][ii][iii])
            save.to_excel(writer,sheet_name=ssp[ii] + '_' + type_list[iii])
    writer._save()
text_kobe = tk.Label(text='โปรแกรมนี้พัตนาโดย KobEbaN WE 37 \n โดยทำการรวมภาพจากโดรน \n ขั้นตอนที่1: เลือกโฟร์เดอร์ไฟล์ \n ขั้นตอนที่2: ตั้งชื่อไฟล์ที่บันทึก')
text_kobe.pack()
label = tk.Label(window, text="")
label.pack()


# สร้างปุ่มสำหรับเลือกไฟล์
button_select_and_save = tk.Button(window, text="เลือกโฟลเดอร์และบันทึกไฟล์", command=select_folder_and_save_file)
button_select_and_save.pack()
status_label = tk.Label(window, text="")  # ป้ายกำกับสำหรับแสดงสถานะ
status_label.pack()
# แสดงหน้าต่าง
window.mainloop()
