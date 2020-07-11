import requests
import json
import xlrd
import configparser
import mysql.connector

# 读取excel
def read_excel():
    #记录sql执行次数
    total_count = 0
    # 获取config.ini中保存的地图api调用key
    config = configparser.ConfigParser()
    config.read("config.ini",encoding="utf-8")
    KEY = config.get('geo_map_api_key',"KEY")
    #创建mysql链接
    mydb = mysql.connector.connect(
        host=config.get('formal_mysql_db','host'),  # 数据库主机地址
        user=config.get('formal_mysql_db','user'),  # 数据库用户名
        passwd=config.get('formal_mysql_db','passwd'),  # 数据库密码
        database=config.get('formal_mysql_db','database')  # database
    )
    mycursor = mydb.cursor()
    # 打开文件
    rb = xlrd.open_workbook(filename="1.xlsx")
    # 读取的工作簿
    read_Sheet = rb.sheet_by_index(0)
    # 获取行数
    rows = read_Sheet.nrows

    for i in range(1, 3):
        my_row_values = read_Sheet.row_values(i)
        address = my_row_values[9] + my_row_values[7] + my_row_values[5] + my_row_values[3]
        village_name = my_row_values[1].rstrip('委会')
        geoUrl = "https://restapi.amap.com/v3/geocode/geo?key=%s&address=%s" % (KEY, address+village_name)
        req = requests.get(geoUrl).text
        # 将 JSON 对象转换为 Python 字典
        json_data = json.loads(req)
        longitude_latitude = json_data["geocodes"][0]["location"]
        list_result = longitude_latitude.partition(",")
        result_longitude = list_result[0]
        result_latitude = list_result[2]
        # sql语句
        print("更新%s" % village_name)
        update_sql = "update village set longitude='%s' ,latitude = '%s' where state=1 and village_name='%s'" % (
            result_longitude, result_latitude, village_name)
        mycursor.execute(update_sql)
        mydb.commit()
        total_count += mycursor.rowcount
    print(total_count, " 条记录被修改")

if __name__ == '__main__':
    read_excel()
