import xlrd
from qqwry import QQwry

q = QQwry()
# 加载纯真镜像 dat 文件
# https://github.com/WisdomFusion/qqwry.dat
q.load_file('F:\qqwry.dat')

# 打开存储 IP 的 excel 表格
wb = xlrd.open_workbook(filename="IP.xlsx")
# 对应到目标表
ip = wb.sheet_by_index(0)
# 从上到下，按列
rows = ip.nrows
# 将结果保存在 result.txt 中
ipResult = open("result.txt", 'w')
for i in range(rows):
    # 得到每一个 IP
    ip_value = ip.cell(i, 0).value
    # 查询
    result = q.lookup(ip_value)
    print(result)
    city = result[0]
    isp = result[1]
    ipResult.write(ip_value + '\t' + city + '\t' + isp + '\r\n')
ipResult.close()
