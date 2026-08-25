# 1. 定义一个简单的回复字典
responses = {
    "你好": "你好呀！很高兴认识你！",
    "名字": "我是你亲手写的极简聊天机器人！",
    "天气": "我还没联网，不知道外面天气怎么样呢。",
    "再见": "拜拜！祝你今天开心！"
}

print("=== 极简聊天机器人已启动 (输入'退出'结束) ===")

# 2. 开启对话循环
while True:
    user_input = input("你: ").strip()

    # 退出机制
    if user_input in ["退出", "quit", "exit"]:
        print("机器人: 拜拜！")
        break

    # 3. 核心逻辑：关键词匹配
    matched = False
    for keyword, reply in responses.items():
        if keyword in user_input:
            print(f"机器人: {reply}")
            matched = True
            break

    # 如果没匹配到任何关键词
    if not matched:
        print("机器人: 哎呀，我还在学习中，听不懂这句话呢...")