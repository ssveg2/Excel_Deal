import openpyxl, pprint                                                   #导入相关库文件
print(openpyxl.__version__)                                               #查看openpyxl库版本
print('Opening workbook...')
wb = openpyxl.load_workbook('G:/Python_Learning/TEC_Python/automate_online-materials/censuspopdata.xlsx') #加载Excel文件
sheet = wb['Population by Census Tract']                                  #获取指定名称工作表
countyData = {}                                                           #定义空字典
print('Reading rows...')
for row in range(2, sheet.max_row + 1):                                   #遍历行，返回[2,工作表最大行数]
    State = sheet['B' + str(row)].value                                   #将单元格'Bstr(row)'赋值给State
    county = sheet['C' + str(row)].value                                  #将单元格'Cstr(row)'赋值给county
    pop = sheet['D' + str(row)].value                                     #将单元格'Dstr(row)'赋值给pop
    countyData.setdefault(State, {})                                      #初始化State键的值；
    countyData[State].setdefault(county, {'tracts': 0, 'pop': 0})         #初始化键值的字典
    countyData[State][county]['tracts'] += 1                              #键值自加1
    countyData[State][county]['pop'] += int(pop)                          #人口叠加
print('Writing results')
resultFile = open('census2010.py', 'w')                                   #打开文件
resultFile.write('allData = ' + pprint.pformat(countyData))               #写入文件
resultFile.close()
print('Done.')

