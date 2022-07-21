# 因為excel欄位有字數限制，WKTFormat太長的線段無法寫入，直接改寫入csv
# 寫入的csv不要使用excel開啟，需要用 Brackets 或 Notepad++

import requests
from bs4 import BeautifulSoup
import csv

read_route = open("GGIT Gas Pipeline WKT too long.csv")  # 讀取csv檔
reader = csv.reader(read_route)

write_route = open("GGIT Gas Pipeline WKT too long result.csv", 'w')  # 寫入csv檔
wr = csv.writer(write_route)

for column in reader:
    URL = column[1]  # 批次讀取csv第二攔
    # print(URL)
    response = requests.get(URL)  # 爬蟲
    soup = BeautifulSoup(response.text, "html.parser")
    div_web = soup.find_all("div", class_="mapdata") and soup.find_all("div", id="map_google3_1")
    div_str = "".join(map(str, div_web))
    # print(div_str)

    if ',"pos":[{' in div_str and div_str.count(',"pos":[{') == 1:  # Linestring格式
        L_startsplit = div_str.split(',"pos":')  # 去除非坐標訊息
        for L_1 in L_startsplit:
            if L_1.startswith('[{"lat":'):
                L_2 = L_1.split(',"polygons":')
                for L_3 in L_2:
                    if L_3.startswith('[{"lat":') and L_3.endswith('}]'):
                        L_str = L_3.replace('[', '').replace('{"lat":', '') \
                            .replace(',"lon":', ',').replace('},', '; ').replace('}]}]', '')
                        Lturn_1 = L_str.split('; ')
                        my_Linestring = ''
                        for Lturn_2 in Lturn_1:  # 經緯度交換
                            Lturn_3 = Lturn_2.split(',')
                            Lturn_4 = Lturn_3[1] + ' ' + Lturn_3[0]
                            my_Linestring += str(Lturn_4 + ", ")  # 整合WKT格式
                        result_1 = ("LINESTRING (" + my_Linestring[:-2] + ")")
                        print(result_1)
                        wr.writerow([result_1])  # 寫入csv

    elif ',"pos":[{' in div_str and div_str.count(',"pos":[{') > 1:  # Multilinestring格式
        M_step1 = div_str.split(',"pos":')  # 去除前段非坐標訊息
        Mmid_final = ''  # 整合最後一段以外的Linestring預留
        Mend_list = ''  # 為了加入最後一段的Linestring預留
        for M_step2 in M_step1:
            if M_step2.startswith('[{"lat":'):
                if M_step2.endswith(',"strokeWeight":"2"'):  # 去除中段非坐標訊息
                    M_step3_mid = M_step2.split(',{"text":')
                    for M_step4_mid in M_step3_mid:
                        if M_step4_mid.startswith('[{"lat":'):
                            # print(M_step4_mid)
                            M_str_mid = M_step4_mid.replace('[', '').replace('{"lat":', '') \
                                .replace(',"lon":', ',').replace('},', '; ').replace('}]}', '')
                            M_turn1_mid = M_str_mid.split('; ')
                            my_Mmid = ''
                            for M_turn2_mid in M_turn1_mid:  # 經緯度交換
                                M_turn3_mid = M_turn2_mid.split(',')
                                M_turn4_mid = M_turn3_mid[1] + ' ' + M_turn3_mid[0]
                                my_Mmid += str(M_turn4_mid + ", ")
                            Mmid_list = "(" + my_Mmid[:-2] + "), "  # 整合WKT格式
                            # print(Mmid_list)
                            for concat in Mmid_list:  # 最後一段以外的Linestring整併
                                Mmid_final += concat

                elif M_step2.endswith('</div>'):  # 去除後段非坐標訊息
                    M_step3_end = M_step2.split('],"polygons":')
                    for M_step4_end in M_step3_end:
                        if M_step4_end.startswith('[{"lat":'):
                            # print(M_step4_end)
                            M_str_end = M_step4_end.replace('[', '').replace('{"lat":', '') \
                                .replace(',"lon":', ',').replace('},', '; ').replace('}]}', '')
                            M_turn1_end = M_str_end.split('; ')
                            my_Mend = ''
                            for M_turn2_end in M_turn1_end:  # 經緯度交換
                                M_turn3_end = M_turn2_end.split(',')
                                M_turn4_end = M_turn3_end[1] + ' ' + M_turn3_end[0]
                                my_Mend += str(M_turn4_end + ", ")
                            Mend_list = ("(" + my_Mend[:-2] + "))")  # 整合WKT格式
                            # print(Mend_list)

        result_2 = ("MULTILINESTRING (" + Mmid_final + Mend_list)  # 加最後一段Linestring
        print(result_2)
        wr.writerow([result_2])  # 寫入csv

    else:
        print("Something wrong!")  # 錯誤提示

print("Finish!")
