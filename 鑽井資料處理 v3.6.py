import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import csv
import random

location = pd.read_excel("location.xlsx")
df = pd.read_excel("code.xlsx")

Location_Pid_plus_Hid = []
code_Pid_plus_Hid = []

###---位置資訊處裡---###

Location_missing_data = []
Location_total_data_Err = []
Location_str_list = []
for i in range(0, len(location)):
    Location_str_list.append('%s %s'%(location['source_x'][i], location['source_y'][i]))
    if str(location['elevation'][i]) == 'nan' or str(location['source_x'][i]) == 'nan' or str(location['source_y'][i]) == 'nan':  #位置資料缺漏搜尋
        Location_missing_data.append(i+2)
        Location_total_data_Err.append(i+2)

Location_unique_value = np.unique(Location_str_list)
Location_repeat_value = []
for i in range(0, len(Location_unique_value)):
    if Location_str_list.count(Location_unique_value[i]) != 1:
        Location_repeat_value.append(Location_unique_value[i])

Location_repeat_list = []
for i in range(0, len(Location_str_list)):
    if Location_str_list[i] in Location_repeat_value:
        Location_repeat_list.append(i+2)

print('位置資訊共', len(location), '筆')
print('位置資訊缺失共', len(Location_missing_data) ,'筆, 檢索位置請輸入 Location_missing_data')
print('位置資料重複共', len(Location_repeat_list), '筆, 檢索位置請輸入 Location_repeat_list , 製表請輸入 Location_repeat_output()')
print('')

###---材料資訊處理---###

Total_code_NaN = 0
Total_discontinuous = 0
code_total_Num = 1
code_total_first_Err = 0
code_total_logic_Err = 0
code = []
code_dic = {}
code_NaN_locade = []
discontinuous = []
code_total_data_Err = []
code_first_Err_list = []
code_logic_Err_list = []
code_loss_index = []
code_loss_dic = {}

for i in range(0 , len(df)-1):
    
    B = df['bottom_depth'][i]
    T0 = df['top_depth'][i]
    T = df['top_depth'][i+1]
    N0 = df['hole_no'][i]
    N1 = df['hole_no'][i+1]    
    
    if str(df['code'][i]) != 'nan' :   #code資料缺漏搜尋
        if df['code'][i] not in code:
            code.append(df['code'][i])
            code_dic[df['code'][i]] = abs(df['bottom_depth'][i] - df['top_depth'][i])
        else:
            code_dic[df['code'][i]] = code_dic[df['code'][i]] + abs(df['bottom_depth'][i] - df['top_depth'][i])
    else:
        Total_code_NaN += 1
        code_NaN_locade.append(i+2)
        if df['hole_no'][i] not in code_loss_index:   #缺漏資料建立
            code_loss_index.append(df['hole_no'][i])
            code_loss_dic[df['hole_no'][i]] = abs(B - T0)
        else:
            code_loss_dic[df['hole_no'][i]] = code_loss_dic[df['hole_no'][i]] + abs(B - T0)
        if df['hole_no'][i] not in code_total_data_Err:
            code_total_data_Err.append(df['hole_no'][i])
    
    
    if N0 == N1 :
        if B != T:  #連續性錯誤搜尋
            Total_discontinuous += 1
            discontinuous.append([i+2, i+3])            
            if df['hole_no'][i] not in code_loss_index:   #缺漏資料建立
                code_loss_index.append(df['hole_no'][i])
                code_loss_dic[df['hole_no'][i]] = abs(B - T)
            else:
                code_loss_dic[df['hole_no'][i]] = code_loss_dic[df['hole_no'][i]] + abs(B - T)
            if df['hole_no'][i] not in code_total_data_Err:
                code_total_data_Err.append(df['hole_no'][i])
    else:
        code_total_Num += 1
        if T != 0:  #首位錯誤搜尋
            code_total_first_Err +=1
            code_first_Err_list.append(i+3)
            if df['hole_no'][i+1] not in code_total_data_Err:
                code_total_data_Err.append(df['hole_no'][i+1])


    if T0 >= B :   #邏輯錯誤搜尋
        code_total_logic_Err += 1
        code_logic_Err_list.append(i+2)
        if df['hole_no'][i] not in code_total_data_Err:
            code_total_data_Err.append(df['hole_no'][i])

###---尾端處理---###

if str(df['code'][len(df)-1]) != 'nan' :   #code資料缺漏搜尋
    if df['code'][len(df)-1] not in code:
        code.append(df['code'][len(df)-1])
        code_dic[df['code'][len(df)-1]] = abs(df['bottom_depth'][len(df)-1] - df['top_depth'][len(df)-1])
    else:
        code_dic[df['code'][len(df)-1]] = code_dic[df['code'][len(df)-1]] + abs(df['bottom_depth'][len(df)-1] - df['top_depth'][len(df)-1])
else:
    Total_code_NaN += 1
    code_NaN_locade.append(len(df)+1)
    
    if df['hole_no'][len(df)-1] not in code_loss_index:   #缺漏資料建立
        code_loss_index.append(df['hole_no'][len(df)-1])
        code_loss_dic[df['hole_no'][len(df)-1]] = df['bottom_depth'][len(df)-1] - df['top_depth'][len(df)-1]
    else:
        code_loss_dic[df['hole_no'][len(df)-1]] = code_loss_dic[df['hole_no'][len(df)-1]] + df['bottom_depth'][len(df)-1] - df['top_depth'][len(df)-1]
    
    if df['hole_no'][len(df)-1] not in code_total_data_Err:
        code_total_data_Err.append(df['hole_no'][len(df)-1])


if df['top_depth'][len(df)-1] >= df['bottom_depth'][len(df)-1]:   #邏輯錯誤搜尋
    code_total_logic_Err += 1
    code_logic_Err_list.append(len(df)+1)
    if df['hole_no'][len(df)-1] not in code_total_data_Err:
        code_total_data_Err.append(df['hole_no'][len(df)-1])

###---後續統計---###

def code_percentage_print():   #材料資訊出圖
    print('請輸入圖名:', end = '')
    title = str(input())
    code_value = []
    code_print_value = []
    code_print_labels = []
    code_value_else = 0
    for i in range(0,len(code)):
        code_value.append(code_dic[code[i]])
    
    for i in range(0,len(code)):
        if code_value[i]/sum(code_value) < 0.02:
            code_value_else += code_value[i]
        else:
            code_print_value.append(code_value[i])
            code_print_labels.append(code[i])
    code_print_labels.append('else')
    code_print_value.append(code_value_else)
    
    plt.pie(code_print_value,               # 數值
            labels = code_print_labels,     # 標籤
            autopct = "%1.1f%%",            # 將數值百分比並留到小數點一位
            pctdistance = 0.9,              # 數字距圓心的距離
            textprops = {"fontsize" : 10},  # 文字大小
            shadow=False)                   # 設定陰影
    plt.title(title)
    plt.axis('equal')
    plt.savefig('%s.png'%title, dpi=300)
    plt.show()

def mix():   #案號與孔號合併
    for i in range(0, len(location)):
        Location_Pid_plus_Hid.append(['%s %s' %(location['proj_no'][i], location['hole_no'][i])])
    for i in range(0, len(df)):
        code_Pid_plus_Hid.append(['%s %s' %(df['proj_no'][i], df['hole_no'][i])])
    
    with open('Location_mix.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(Location_Pid_plus_Hid)

    with open('code_mix.csv', 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerows(code_Pid_plus_Hid)
            
    print('mix_done')

def check_repeat_hole_number():   #檢測重複項目
    U = np.unique(location['hole_no'])
    LO = location['hole_no'].tolist()
    location_repeat_Num = 0
    code_repeat_Num = 0
    print('檢查location資料重複項:')
    for i in range(0, len(U)):
        if LO.count(U[i]) != 1:
            print(U[i])
            location_repeat_Num += 1
    
    print('')
    print('檢查code資料重複項(基於首位為0):')   
    CO = []
    for i in range(0, len(df['hole_no'])):
        if df['top_depth'][i] == 0 :
            CO.append(df['hole_no'][i])
    u = np.unique(df['hole_no'])
    for i in range(0, len(u)):
        if CO.count(u[i]) != 1:
            print(u[i])
            code_repeat_Num += 1
            
    if location_repeat_Num != 0:
        print('location資料重複項數量:', location_repeat_Num)

    if code_repeat_Num != 0:
        print('code資料重複項數量:', code_repeat_Num)


def Location_repeat_output():   #重複位置資料製表
    Location_repeat_output_list = [['proj_no', 'hole_no', 'elevation', 'coord_system', 'source_x', 'source_y']]
    for i in range(0, len(Location_repeat_list)):
        Location_repeat_output_list.append([location['proj_no'][Location_repeat_list[i]-2], location['hole_no'][Location_repeat_list[i]-2], location['elevation'][Location_repeat_list[i]-2], location['coord_system'][Location_repeat_list[i]-2], location['source_x'][Location_repeat_list[i]-2], location['source_y'][Location_repeat_list[i]-2]])
    with open('Location_repeat.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(Location_repeat_output_list)
    print('輸出成功')

def check_lost_data():   #缺漏資料檢索
    ISO_code = np.unique(df['hole_no']).tolist()
    ISO_location = np.unique(location['hole_no']).tolist()
    Location_loss = np.setdiff1d(ISO_code, ISO_location)
    Code_loss = np.setdiff1d(ISO_location, ISO_code)
    print('缺漏位置:', Location_loss)
    print('缺漏鑽井: ', Code_loss)
    
    if len(Location_loss) != 0:
        print('缺漏位置數量:', len(Location_loss))
    if len(Code_loss) != 0:
        print('缺漏鑽井數量:', len(Code_loss)) 

def output():   #輸出可用於GMS之檔案

    ask = input('輸出與細切程序無矯錯功能，請確認資料無任何錯誤再行輸出，繼續執行? (Y/N)')
    if ask != 'Y':
        print('程序結束')
        return
    
    dic_locat = {}
    for i in range(0,len(location)):
        locat_data = []
        locat_data.append(location['elevation'][i])
        locat_data.append(location['source_x'][i])
        locat_data.append(location['source_y'][i])
        dic_locat[location['hole_no'][i]] = locat_data
    
    hole_id = 1
    New_Data = [['id', 'name', 'x', 'y', 'z', 'material']]
    for i in range(0,len(df)-1):
        ID = df['hole_no'][i]
        New_Data.append([str(hole_id), df['hole_no'][i], dic_locat[ID][1], dic_locat[ID][2], dic_locat[ID][0] - df['top_depth'][i], df['code'][i]])
        if df['top_depth'][i+1] == 0:
            New_Data.append([str(hole_id), df['hole_no'][i], dic_locat[ID][1], dic_locat[ID][2], dic_locat[ID][0] - df['bottom_depth'][i], df['code'][i]])
            hole_id += 1
    New_Data.append([str(hole_id), df['hole_no'][len(df)-1], dic_locat[ID][1], dic_locat[ID][2], dic_locat[ID][0] - df['top_depth'][len(df)-1], df['code'][len(df)-1]])
    New_Data.append([str(hole_id), df['hole_no'][len(df)-1], dic_locat[ID][1], dic_locat[ID][2], dic_locat[ID][0] - df['bottom_depth'][len(df)-1], df['code'][len(df)-1]])
    
    with open('New_Data.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(New_Data)
    print('輸出成功')

def output_cut():   #輸出細切檔案
    dic_locat = {}
    cut_size = 0.1   #切割大小，小於0.1的話下面的浮點數整除判斷要改為 X100 甚至 X1000 ，總之要乘到變成整數位
    for i in range(0,len(location)):
        locat_data = []
        locat_data.append(location['elevation'][i])
        locat_data.append(location['source_x'][i])
        locat_data.append(location['source_y'][i])
        dic_locat[location['hole_no'][i]] = locat_data
    
    process_step = 10   #計數器起始值
    hole_id = 1
    New_Data = [['id', 'name', 'x', 'y', 'z', 'material']]
    for i in range(0,len(df)-1):
        process = 100*(i / len(df))
        if process > process_step:
            print('完成度 %s'%(process_step),'%')
            process_step += 10   #計數器跳動值
        ID = df['hole_no'][i]
        New_Data.append([str(hole_id), df['hole_no'][i], dic_locat[ID][1], dic_locat[ID][2], dic_locat[ID][0] - df['top_depth'][i], df['code'][i]])
        
        cut_times = int(((df['bottom_depth'][i] *10) - (df['top_depth'][i] *10)) / (cut_size *10))   #浮點數問題，進位
        if ((df['bottom_depth'][i] - df['top_depth'][i]) * 10) % (cut_size * 10) == 0:   #這邊用了一個很爛的方法解決浮點數整除問題，不知道會不會有衍伸錯誤產生，未來有更好的方法建議修正
            cut_times -= 1
        if cut_times > 0:
            for j in range(1, cut_times + 1):
                New_Data.append([str(hole_id), df['hole_no'][i], dic_locat[ID][1], dic_locat[ID][2], dic_locat[ID][0] - df['top_depth'][i] - (j * cut_size), df['code'][i]])
        if df['top_depth'][i+1] == 0:
            New_Data.append([str(hole_id), df['hole_no'][i], dic_locat[ID][1], dic_locat[ID][2], dic_locat[ID][0] - df['bottom_depth'][i], df['code'][i]])
            hole_id += 1

    New_Data.append([str(hole_id), df['hole_no'][len(df)-1], dic_locat[ID][1], dic_locat[ID][2], dic_locat[ID][0] - df['top_depth'][len(df)-1], df['code'][len(df)-1]])
    cut_times = int(((df['bottom_depth'][len(df)-1] *10) - (df['top_depth'][len(df)-1] *10)) / (cut_size *10))   #浮點數問題，進位
    if ((df['bottom_depth'][len(df)-1] - df['top_depth'][len(df)-1]) * 10) % (cut_size * 10) == 0:   #這邊用了一個很爛的方法解決浮點數整除問題，不知道會不會有衍伸錯誤產生，未來有更好的方法建議修正
        cut_times -= 1
    if cut_times > 0:
        for j in range(1, cut_times + 1):
            New_Data.append([str(hole_id), df['hole_no'][len(df)-1], dic_locat[ID][1], dic_locat[ID][2], dic_locat[ID][0] - df['top_depth'][len(df)-1] - (j * cut_size), df['code'][len(df)-1]])    
    New_Data.append([str(hole_id), df['hole_no'][len(df)-1], dic_locat[ID][1], dic_locat[ID][2], dic_locat[ID][0] - df['bottom_depth'][len(df)-1], df['code'][len(df)-1]])
    
    with open('New_Data_cutted.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(New_Data)
    print('輸出成功')

def New_cut(): #新的細切程序
    location = pd.read_csv("New_Data.csv")
    id = 1
    New_Data = [['id', 'name', 'x', 'y', 'z', 'material']]
    
    print('正在初步切割', end='')
    timer_sep = int(len(location)/10)
    timer = int(len(location)/10)
    for i in range(0, len(location)-1):
        New_Data.append([str(id), location['name'][i], location['x'][i], location['y'][i], round(location['z'][i], 3), location['material'][i]])
        if location['name'][i] == location['name'][i+1]:
            cut_time = int(((location['z'][i] - location['z'][i+1])/0.1))
            for j in range(1, cut_time+1):
                New_Data.append([str(id), location['name'][i], location['x'][i], location['y'][i], round(location['z'][i] - (0.1*j), 3), location['material'][i]])
        else:
            id += 1
        
        if i > timer:
            print('.', end = '')
            timer += timer_sep
    New_Data.append([str(id), location['name'][len(location)-1], location['x'][len(location)-1], location['y'][len(location)-1], round(location['z'][len(location)-1], 3), location['material'][len(location)-1]])
    
    print('done')
    print('正在除錯', end = '')
    timer_sep = int(len(New_Data)/10)
    timer = int(len(New_Data)/10)
    
    Out_list = [['id', 'name', 'x', 'y', 'z', 'material']]
    for i in range(1, len(New_Data)-1):
        if New_Data[i][4] == New_Data[i+1][4]:
            if New_Data[i][1] == New_Data[i+1][1]:
                continue
        Out_list.append(New_Data[i])
        
        if i > timer:
            print('.', end = '')
            timer += timer_sep
    
    Out_list.append(New_Data[len(New_Data)-1])
    
    print('done')
    print('正在輸出', end = '')
    
    with open('New_Cut.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(Out_list)
    print('..........done')
    print('完成, 程序結束')



def draw():
    def event():
        result = random.randint(1,1000)
        if result <= 25:
            result = 3
        if result >= 26 and result <= 205:
            result = 2
        if result >= 206:
            result = 1
        return result
    
    def last():
        result = random.randint(1,1000)
        if result <= 25:
            result = 3
        else:
            result = 2
        return result
    
    OUT = []
    for _ in range(9):
        OUT.append(event())
    OUT.append(last())
    return OUT

def Handsome():
    print('陳建語')


print('材料資訊共計', code_total_Num,'孔鑽井,')
print('總計', len(code), '種材料,檢索材料代碼請輸入 code,')
print('查詢各別材料所占總厚度請輸入 code_dic, 製圖請輸入code_percentage_print()')
print('')
print('總計', len(code_NaN_locade), '筆缺失資料,檢索位置請輸入 code_NaN_locade')
print('總計', code_total_first_Err,'筆首位錯誤, 檢索位置請輸入 code_first_Err_list')
print('總計', Total_discontinuous, '筆資料不連續, 檢索位置請輸入 discontinuous')
print('總計', code_total_logic_Err, '筆邏輯錯誤, 檢索位置請輸入code_logic_Err_list')
print('各孔總缺漏厚度檢索請輸入 code_loss_dic , 檢視索引請輸入code_loss_index')
print('')
print('輸出csv請輸入 output()')
print('輸出細切檔案請輸入 output_cut()')
print('注1:目前為測試版, 可能有意料外的error, 請確定資料品質穩定再行輸出')
print('注2:細切過程中遇到浮點數整除問題，暫時以進位方法判斷，希望未來可以修正為更直接的判斷方法')
print('先行使用: mix() 測試: check_lost_data() check_repeat_hole_number()')