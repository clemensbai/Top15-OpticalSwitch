import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

data = pd.ExcelFile('D:\Python-3.11\Optical_switch_modeling.xlsx', engine='openpyxl')
sheet_names = data.sheet_names
print(sheet_names)

categories = [] 
value1 = []
value2 = []
value3 = []
value4 = []
value5 = []
Availability = 0.999654619175739

for sheet in sheet_names:
    if sheet != 'MTBF':
        if sheet != 'Receiver_Module':
            if sheet == 'Amplifier':
                df = pd.read_excel('D:\Python-3.11\Optical_switch_modeling.xlsx', sheet_name=sheet).fillna('')
                Ava = df.iloc[14, 0]
                df_ = df.iloc[7:12, :]
                data1 = df_.iloc[0:5, 1:6]
                df1 = df_.iloc[:, 0]
                print(Ava)
                for row in df1:
                    categories.append(row)
                data2 = data1.iloc[:, 0]
                for row in data2:
                    value1.append(row * row * Ava)
                data2 = data1.iloc[:, 1]
                for row in data2:
                    value2.append(row * row * Ava)
                data2 = data1.iloc[:, 2]
                for row in data2:
                    value3.append(row * row * Ava)
                data2 = data1.iloc[:, 3]
                for row in data2:
                    value4.append(row * row * Ava)
                data2 = data1.iloc[:, 4]
                for row in data2:
                    value5.append(row * row * Ava)
            elif sheet =='Multiplexer_Demultiplexer':
                df = pd.read_excel('D:\Python-3.11\Optical_switch_modeling.xlsx', sheet_name=sheet).fillna('')
                Ava = df.iloc[14, 0]
                df_ = df.iloc[7:12, :] 
                data1 = df_.iloc[0:5, 1:6]
                df1 = df_.iloc[:, 0]
                print(Ava)
                for row in df1:
                    categories.append(row)
                data2 = data1.iloc[:, 0]
                for row in data2:
                    value1.append(row * row * Ava)
                data2 = data1.iloc[:, 1]
                for row in data2:
                    value2.append(row * row * Ava)
                data2 = data1.iloc[:, 2]
                for row in data2:
                    value3.append(row * row * Ava)
                data2 = data1.iloc[:, 3]
                for row in data2:
                    value4.append(row * row * Ava)
                data2 = data1.iloc[:, 4]
                for row in data2:
                    value5.append(row * row * Ava)
            elif sheet == 'Transceiver':
                df = pd.read_excel('D:\Python-3.11\Optical_switch_modeling.xlsx', sheet_name=sheet).fillna('')
                Ava = df.iloc[14, 0]
                df_ = df.iloc[7:12, :]
                data1 = df_.iloc[0:5, 1:6]
                df1 = df_.iloc[:, 0]
                print(Ava)
                for row in df1:
                    categories.append(row)
                data2 = data1.iloc[:, 0]
                for row in data2:
                    value1.append(row * row * Ava)
                data2 = data1.iloc[:, 1]
                for row in data2:
                    value2.append(row * row * Ava)
                data2 = data1.iloc[:, 2]
                for row in data2:
                    value3.append(row * row * Ava)
                data2 = data1.iloc[:, 3]
                for row in data2:
                    value4.append(row * row * Ava)
                data2 = data1.iloc[:, 4]
                for row in data2:
                    value5.append(row * row * Ava)
            elif sheet == 'Transponder':
                df = pd.read_excel('D:\Python-3.11\Optical_switch_modeling.xlsx', sheet_name=sheet).fillna('')
                Ava = df.iloc[14, 0]
                df_ = df.iloc[7:12, :]
                data1 = df_.iloc[0:5, 1:6]
                df1 = df_.iloc[:, 0]
                print(Ava)
                for row in df1:
                    categories.append(row)
                data2 = data1.iloc[:, 0]
                for row in data2:
                    value1.append(row * row * Ava)
                data2 = data1.iloc[:, 1]
                for row in data2:
                    value2.append(row * row * Ava)
                data2 = data1.iloc[:, 2]
                for row in data2:
                    value3.append(row * row * Ava)
                data2 = data1.iloc[:, 3]
                for row in data2:
                    value4.append(row * row * Ava)
                data2 = data1.iloc[:, 4]
                for row in data2:
                    value5.append(row * row * Ava)
            elif sheet == 'Wavelength_Coupler_Splitter':
                df = pd.read_excel('D:\Python-3.11\Optical_switch_modeling.xlsx', sheet_name=sheet).fillna('')
                Ava = df.iloc[14, 0]
                df_ = df.iloc[7:12, :]
                data1 = df_.iloc[0:5, 1:6]
                df1 = df_.iloc[:, 0]
                print(Ava)
                for row in df1:
                    categories.append(row)
                data2 = data1.iloc[:, 0]
                for row in data2:
                    value1.append(row * row * Ava)
                data2 = data1.iloc[:, 1]
                for row in data2:
                    value2.append(row * row * Ava)
                data2 = data1.iloc[:, 2]
                for row in data2:
                    value3.append(row * row * Ava)
                data2 = data1.iloc[:, 3]
                for row in data2:
                    value4.append(row * row * Ava)
                data2 = data1.iloc[:, 4]
                for row in data2:
                    value5.append(row * row * Ava)
            else:
                df=pd.read_excel('D:\Python-3.11\Optical_switch_modeling.xlsx', sheet_name=sheet).fillna('')
                Ava = df.iloc[14, 0]
                df_ = df.iloc[7:12, :]
                data1 = df_.iloc[0:5,1:6]
                df1 = df_.iloc[:,0]
                print(Ava)
                for row in df1:
                    categories.append(row)
                data2 = data1.iloc[:,0]
                for row in data2:
                    value1.append(row * Ava)
                data2 = data1.iloc[:,1]
                for row in data2:
                    value2.append(row * Ava)
                data2 = data1.iloc[:,2]
                for row in data2:
                    value3.append(row * Ava)
                data2 = data1.iloc[:,3]
                for row in data2:
                    value4.append(row * Ava)
                data2 = data1.iloc[:,4]
                for row in data2:
                    value5.append(row * Ava)

        else:
            df = pd.read_excel('D:\Python-3.11\Optical_switch_modeling.xlsx', sheet_name=sheet).fillna('')
            Ava = df.iloc[24, 0]
            df_ = df.iloc[27:32, :]
            data1 = df_.iloc[0:5, 1:6]
            df1 = df_.iloc[:, 0]
            print(Ava)
            for row in df1:
                categories.append(row)
            data2 = data1.iloc[:, 0]
            for row in data2:
                value1.append(row * Ava)
            data2 = data1.iloc[:, 1]
            for row in data2:
                value2.append(row * Ava)
            data2 = data1.iloc[:, 2]
            for row in data2:
                value3.append(row * Ava)
            data2 = data1.iloc[:, 3]
            for row in data2:
                value4.append(row * Ava)
            data2 = data1.iloc[:, 4]
            for row in data2:
                value5.append(row * Ava)
print(value5[0:10])
print(categories[0:10])
offset1 = [value1[i] - Availability for i in range(min(len(value1), len(value3)))]
offset2 = [value2[i] - Availability for i in range(min(len(value2), len(value3)))]
offset3 = [value4[i] - Availability for i in range(min(len(value4), len(value3)))]
offset4 = [value5[i] - Availability for i in range(min(len(value5), len(value3)))]
# abs_max_value = max(max(offset1, key=abs), max(offset2, key=abs), max(offset3, key=abs),
#                             max(offset4, key=abs)) * 1.05
min_value = min(min(value1, key=abs), min(value2, key=abs), min(value3, key=abs), min(value4, key=abs), min(value5, key=abs))
max_value = max(max(value1, key=abs), max(value2, key=abs), max(value3, key=abs), max(value4, key=abs), max(value5, key=abs))
min_offset = []
for i in range(len(offset1)):
    A_1 = max(abs(offset1[i]), abs(offset4[i]), abs(offset2[i]), abs(offset3[i]))
    min_offset.append(A_1)
n = len(min_offset)
for i in range(n):
    for j in range(0, n - i - 1):
        if min_offset[j] > min_offset[j + 1]:
            min_offset[j], min_offset[j + 1] = min_offset[j + 1], min_offset[j]
            categories[j], categories[j + 1] = categories[j + 1], categories[j]
            offset1[j], offset1[j + 1] = offset1[j + 1], offset1[j]
            offset2[j], offset2[j + 1] = offset2[j + 1], offset2[j]
            offset3[j], offset3[j + 1] = offset3[j + 1], offset3[j]
            offset4[j], offset4[j + 1] = offset4[j + 1], offset4[j]

categories_1 = categories[-15:]
offset_1 = offset1[-15:]
offset_2 = offset2[-15:]
offset_3 = offset3[-15:]
offset_4 = offset4[-15:]
abs_max_value = max(max(offset_1, key=abs), max(offset_2, key=abs), max(offset_3, key=abs),
                            max(offset_4, key=abs)) * 1.3
fig, ax = plt.subplots(1,1, figsize = (14,8))
ax.barh(categories_1, offset_4, align='center', color='red', label='+50%')
ax.barh(categories_1, offset_3, align='center', color='black', label='+25%')
ax.barh(categories_1, offset_1, align='center', color='blue', label='-50%')
ax.barh(categories_1, offset_2, align='center', color='green', label='-25%')

ax.axvline(0.0, color='black', linewidth=1)

plt.legend(fontsize=16)
plt.xlabel('Deviation in availability', fontsize=22)
plt.ylabel('Input Parameter', fontsize=22)
plt.yticks(rotation=0, fontsize=20)

ax.set_xlim((-1* abs_max_value, abs_max_value))
ax_xticks = np.linspace(-1* abs_max_value, abs_max_value, 11)
ax_xticks1 = [round(i, 8) for i in ax_xticks]
ax.set_xticks(ax_xticks1)
ax.xaxis.set_tick_params(labelsize=18, rotation=20)
ax.grid(True)

ax2 = ax.twiny()
ax2.set_xticks(ax.get_xticks())
ax2.set_xbound(ax.get_xbound())
ax2_xticks = np.linspace(min_value,max_value,11)
ax2_xticks[5] = value3[0]
ax2_xticks1 = [round(i,8) for i in ax2_xticks]
ax2.set_xticklabels(ax2_xticks1,rotation = 20, fontsize = 18)
ax.get_xticklabels()[5].set_weight("bold")
ax2.get_xticklabels()[5].set_weight("bold")

ax.invert_xaxis()
ax2.invert_xaxis()
plt.show()
