import openpyxl

# 创建一个新的Excel工作簿
workbook = openpyxl.Workbook()

# 选择默认活动工作表
worksheet = workbook.active

# 添加表格标题
worksheet['A1'] = '隐患排查表'

# 添加列名
worksheet['A2'] = '信息系统清单'
worksheet['B2'] = '建设单位'
worksheet['C2'] = '联系人'
worksheet['D2'] = '隐患排查发现个数及类型'
worksheet['E2'] = '整改完成情况'
worksheet['F2'] = '网络安全承诺签订情况'
worksheet['G2'] = '数据安全承诺签订情况'
worksheet['H2'] = '非政务云环境留存数据情况'

# 设置列宽
worksheet.column_dimensions['A'].width = 20
worksheet.column_dimensions['B'].width = 20
worksheet.column_dimensions['C'].width = 20
worksheet.column_dimensions['D'].width = 35
worksheet.column_dimensions['E'].width = 20
worksheet.column_dimensions['F'].width = 25
worksheet.column_dimensions['G'].width = 25
worksheet.column_dimensions['H'].width = 30

# 填写数据
data = [
    ['系统名称1', '单位1', '张三', '3个，包括1个风险较高', '已完成', '已签订', '已签订', '无'],
    ['系统名称2', '单位2', '李四', '2个，包括1个风险较高', '未完成', '未签订', '已签订', '有'],
    ['系统名称3', '单位3', '王五', '1个，风险较低', '已完成', '已签订', '未签订', '无']
]

for row in range(3, 6):
    for col in range(1, 9):
        worksheet.cell(row, col, data[row-3][col-1])

# 保存Excel文件
workbook.save('隐患排查表.xlsx')
