import pandas as pd
from openpyxl import load_workbook

csv_path = '정책.csv'
df = pd.read_csv(csv_path, skiprows=1, encoding='cp949')
df['변경_룰셋'] = ''
df['변경_원본 룰셋'] = ''
df['변경여부'] = 'N'

old_input = input("구장비 입력(값들은 쉼표로 구분):")
new_input = input("하남 신장비 입력(값들은 쉼표로 구분):")

old_eqp = [i.strip() for i in old_input.split(',')]
new_eqp = [i.strip() for i in new_input.split(',')]

if len(old_eqp) != len(new_eqp):
    print("장비 개수가 안맞습니다")
    exit()

replace = list(zip(old_eqp, new_eqp))

for row in range(len(df)):
    for col in [8,9]:
        cell_data = str(df.iat[row,col])
        for old,new in replace:
            if old in cell_data:
                df.iat[row,col+27] = df.iat[row,col]
                df.iat[row,col+27] = df.iat[row,col+27].replace(old,new)
                df.iat[row,37] = 'Y'

subset = df.iloc[:, [37,2,8,9,35,36]]
result_path = 'result.xlsx'
subset.to_excel(result_path, index=False, engine='openpyxl')