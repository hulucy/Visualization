# %%
import pandas as pd
import numpy as np
import seaborn as sns
import datetime
# %% 

checkDate = '2021-05-27'
data_path = '/Users/wuhao/Documents/work/龙虎榜'
type_list = ['不含新股第一天', '新股第一天']


# %%

def serial_date_to_string(srl_no):
    new_date = datetime.datetime(1900,1,1) + datetime.timedelta(srl_no - 1)
    return new_date.strftime("%Y-%m-%d")


# %% 解决NaN黑色背景的问题
def color_zero_white(val):
    color = 'white' if val == 0 else 'black'
    return 'color: %s' % color

def color_zero_background(data, color='white'):
    attr = 'background-color: {}'.format(color)
    if data.ndim == 1:  # Series from .apply(axis=0) or axis=1
        is_max = data == 0
        return [attr if v else '' for v in is_max]
    else:  # from .apply(axis=None)
        is_max = data == 0
        return pd.DataFrame(np.where(is_max, attr, ''),
                            index=data.index, columns=data.columns)





# %%
def recolor_longhu(file_name, sheet_name):
    
    df = pd.read_excel(file_name, sheet_name)

    change_columns = list(df.columns[1:len(df.columns)-3])
    change_columns = [x.replace('2021-', '') for x in change_columns]
    df.columns = [type + '-' + sheet_name] + change_columns + ['总次数', '总买入金额（万）', '总利润（万）']


    df.set_index(type + '-' + sheet_name, inplace=True)
    df_col = list(df.columns[:(df.shape[1]-3)])
    df_sub = df[df_col]

    cm = sns.light_palette('red', as_cmap=True)
    # cm.set_under('white')

    dt = df.copy().fillna(0)

    s = dt.style.format("{:.0f}")\
        .set_properties(**{'max-width': '80px', 'font-size': '13pt'})\
        .background_gradient(cmap=cm, axis=None, subset = df_col)\
        .applymap(lambda x: color_zero_white(x))\
        .apply(color_zero_background, axis=None)

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
