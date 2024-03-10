import xlrd
import xlwt
import re

global set_weeks_one_cell
global dic_weeks_all_cols
def judge_part(x):
    """正则提取每个单元格里包含的周数
    value == x
        测试完毕，此函数无问题"""
    global set_weeks_one_cell
    set_weeks_one_cell = set()
    pattern3 = 'n\((\d{1,2}-\d{1,2})[单|双]'
    weeks_num3 = re.findall(pattern3, str(x))
    pattern2 = 'n\((\d{1,2}-\d{1,2})\s'
    weeks_num2 = re.findall(pattern2, str(x))
    pattern1 = 'n\((\d{1,2}\s)'
    weeks_num1 = re.findall(pattern1, str(x))
    if weeks_num1:
        weeks_num1.sort()
        for x in weeks_num1:
            set_weeks_one_cell.add(int(x))
    if weeks_num2:
        y = 0
        while y < len(weeks_num2):
            y += 1
            for i in weeks_num2:
                list_0 = []
                i = i.split('-')
                for z in i:
                    list_0.append(int(z))
                list_0.sort()
                for x in range(int(list_0[0]), int(list_0[1]) + 1):
                    set_weeks_one_cell.add(x)
    if weeks_num3:
        y = 0
        while y < len(weeks_num3):
            y += 1
            for i in weeks_num3:
                list_0 = []
                i = i.split('-')
                for z in i:
                    list_0.append(int(z))
                list_0.sort()
                for x in range(int(list_0[0]), int(list_0[1]) + 1, 2):
                    set_weeks_one_cell.add(x)

def write_at_once():
    """一次性写入全部数据，传入两个字典"""
    global dic_weeks_all_cols
    set_week20 = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20}
    wb = xlwt.Workbook(encoding='utf-8')
    list1 = ['周一', '周二', '周三', '周四', '周五']
    list2 = ['1-2', '3-4', '5-6', '7-8']
    for i in range(1, 21):
        sheet = wb.add_sheet(f'sheet{i}')  # 创建20页空表格
        sheet.write_merge(0, 0, 0, 5, f'第{i}周 无课表')  # 合并单元格，写入大表头
        y = 1
        for x in list1:
            sheet.write(1, y, x)  # 写入星期
            y += 1
        b = 2
        for a in list2:
            sheet.write(b, 0, a)  # 写入节次
            sheet.write_merge(b, b+3, 0, 0, a)
            b += 4
    filename_and_name_dic = {'yanshun.xls': '闫顺', 'yangguang.xls': '杨光', 'yangjian.xls': '杨建',
                             'ljz.xls': '陆金泽'}
    for filename, name in filename_and_name_dic.items():
        workbook = xlrd.open_workbook(filename, formatting_info=True)
        worksheet = workbook.sheet_by_name('sheet1')
        for num in range(1, 5):
            get_info_from_old_excel(worksheet)
            weeks = dic_weeks_all_cols
            for xq, in_dic in weeks.items():  # 总字典，所有的信息都在里面，{星期：{节次：周数集合}}
                # 提取星期xq和里面的字典
                for week_b, week_cell in in_dic.items():  # 提取出{节次+1：周数集合}中的key和value
                    week_cell_set = set(set_week20 - set(week_cell))  # 计算出无课的周次，保存为集合
                    week_cell_list = list(week_cell_set)
                    week_cell_list.sort()
                    for x in week_cell:
                        sheet = wb.get_sheet(f'sheet{x}')
                        sheet.write(int(week_b) + num -1, int(xq), f'{name}')
        wb.save('无课表.xls')

def get_info_from_old_excel(x):
    """需要传入一个worksheet给x,从excel单元格里获取信息，两个字典嵌套储存行和列"""
    global dic_weeks_all_cols
    dic_weeks_all_cols = {}  # 周一至周五课程的字典，key对应星期
    for col in range(1, 6):
        dic_weeks_one_col = {}  # 所有周次的某一天的课程，循环一天清空一次，key对应：行数/节次
        for row in range(4, 11, 2):
            value = x.cell(row, col)
            judge_part(value)
            dic_weeks_one_col[int(row / 2)] = set_weeks_one_cell
        dic_weeks_all_cols[col] = dic_weeks_one_col


if __name__ == '__main__':
    write_at_once()
