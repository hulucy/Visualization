# %%
import pandas as pd
import numpy as np
import seaborn as sns
import datetime
# %% 

data_path = '/Users/wuhao/Documents/work/龙虎榜'
type_list = ['不含新股第一天', '新股第一天']
checkDate = '2021-05-27'

# %%

def serial_date_to_string(srl_no):
    new_date = datetime.datetime(1900,1,1) + datetime.timedelta(srl_no - 1)
    return new_date.strftime("%Y-%m-%d")



# %%
def recolor_longhu(file_name, sheet_name):
    
    df = pd.read_excel(file_name, sheet_name)
    df.head()

    change_columns = list(df.columns[1:len(df.columns)-3])
    change_columns = [x.replace('2021-', '') for x in change_columns]
    df.columns = [type + '-' + sheet_name] + change_columns + ['总次数', '总买入金额（万）', '总利润（万）']


    df.set_index(type + '-' + sheet_name, inplace=True)

    cm = sns.light_palette('red', as_cmap=True)

    s = df.style.format('{:.0f}', na_rep = '')\
            .background_gradient(cmap=cm)\
            .set_properties(**{'max-width': '80px', 'font-size': '13pt'})

    return(s)


# %%
# type = type_list[0]
# file_name = f'{data_path}/{date_range}各券商在新股上的上榜次数-{type}-按周统计.xlsx'
sheet_name_list = ['深圳+上海', '深圳', '上海']

# i = 0
# sheet_name = sheet_name_list[i]
# sheet_name

# %%
i = 0
s = []
for type in type_list:
    file_name = f'{data_path}/{checkDate}各券商在新股上的上榜次数-{type}-按周统计.xlsx'
    for sheet_name in sheet_name_list:
        s.append(recolor_longhu(file_name = file_name, sheet_name = sheet_name))
        i += 1
        print(i, type, sheet_name)




# %%
with pd.ExcelWriter(f'{data_path}/{checkDate}.xlsx') as writer:
    s[0].to_excel(writer, sheet_name = type_list[0] + '-' + sheet_name_list[0])
    s[1].to_excel(writer, sheet_name = type_list[0] + '-' + sheet_name_list[1])
    s[2].to_excel(writer, sheet_name = type_list[0] + '-' + sheet_name_list[2])

    s[3].to_excel(writer, sheet_name = type_list[1] + '-' + sheet_name_list[0])
    s[4].to_excel(writer, sheet_name = type_list[1] + '-' + sheet_name_list[1])
    s[5].to_excel(writer, sheet_name = type_list[1] + '-' + sheet_name_list[2])
# %%
df = pd.read_excel(file_name, sheet_name)
df.head()


# %%
change_columns = list(df.columns[1:len(df.columns)-3])
change_columns = [x.replace('2021-', '') for x in change_columns]
df.columns = [sheet_name] + change_columns + ['总次数', '总买入金额（万）', '总利润（万）']


# %%

df.set_index(sheet_name, inplace=True)

cm = sns.light_palette('red', as_cmap=True)

s = df.style.format('{:.0f}', na_rep = '').background_gradient(cmap=cm)