import openpyxl
from openpyxl import load_workbook
import time
start_time = time.time()
# your code


# data_only=True로 해줘야 수식이 아닌 값으로 받아온다.
filepath="./신계약명세.xlsx"
load_wb = load_workbook(filepath, data_only=True)


# # 셀 주소로 값 출력
# print(load_ws['B2'].value)
elapsed_time = time.time() - start_time
print("Done : file open (",end="")
print(round(elapsed_time,2),end=")")


start_time = time.time()
sheetname='신계약명세(전월)'
# 시트 이름으로 불러오기
load_sheet = load_wb[sheetname]
elapsed_time = time.time() - start_time


def getval(cells):
    result = []
    Nonecheck = False
    for i in cells:
        if i.value is not None:
            Nonecheck = True
        if Nonecheck:
            result.append(i.value)
        else:
            pass
    if result == []:
        return None
    return result


namespace = load_sheet['1']
cd = {}
idx = 0
for i in getval(namespace):
    cd[i] = idx
    idx += 1
print(cd)
get_cells = load_sheet.rows

alltable = []
for row in get_cells:
    if getval(row) is not None:
        alltable.append(getval(row))

# 일단 대상보험사 다불러와

goods = {"KDB": ["오행복", "버팀목"],
         "DGB": ["마이솔", "그랑", "마음든든"],
         "하나": ["하나로"],
         "푸르": ["100세", "달러평생", "달러 평생", "함께"],
         "삼성": ["신성장"],
         "라이나": ["종신"],
         "메트": ["모두의", "백만인"],
         "미래": ["선택하는"],
         "DB": ["알차고", "암종신"]
         }

resulttable = []

for row in alltable:
    for i in goods:
        for k in goods[i]:
            if i in row[cd["보험사명"]] and k in row[cd["상품명"]]:
                resulttable.append(row)

idx = 0

for row in resulttable:
    if int(row[cd["납입기간"]]) > 99999999999999:
        idx += 1
    else:
        del resulttable[idx]

for row in resulttable:
    print(row)

##
##and조건 => 순차적으로
##or조건 => 중복을 세지않고 더하기
# 나중에 모든 column 길이 맞는지 체크해야됨

# 조건 : 포함 and 나 or / 제외