import chromadb

# 1. 创建一个本地客户端（会自动在当前目录生成一个文件夹保存数据）
client = chromadb.PersistentClient(path="./chroma_db")

# 2. 获取或创建一个集合（类似于传统数据库里的“表”）
collection = client.get_or_create_collection(name="chatbot_kb")

# 测试连接是否成功
print(f"成功连接到向量数据库！当前集合包含 {collection.count()} 条数据。")

# 查看前 10 条数据的预览
results = collection.peek()
print("=== 数据库预览 ===")
print(results)