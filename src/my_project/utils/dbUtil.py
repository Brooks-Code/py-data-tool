import mysql.connector

# 1. 建立连接
conn = mysql.connector.connect(
    host='localhost',
    user='root',
    password='1234',
    database='my_python_data'
)

# 2. 创建游标
cursor = conn.cursor()

# 3. 执行查询
cursor.execute("SHOW DATABASES")

# 4. 获取结果
results = cursor.fetchall()
for row in results:
    print(row)

# 5. 关闭游标和连接
cursor.close()
conn.close()