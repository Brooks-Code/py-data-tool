import os # 系统操作

import requests # 响应爬取
from DrissionPage import ChromiumPage #谷歌
import re # 正则表达式
import json
import time

# 1. 【Windows 专属】动态获取当前脚本所在的目录，作为临时文件夹
# 这样无论你把代码放在哪个盘，它都会在代码旁边建一个 tmp 文件夹
current_dir = os.path.dirname(os.path.abspath(__file__))
local_tmp = os.path.join(current_dir, 'tmp')
os.makedirs(local_tmp, exist_ok=True)

# 2. 【Windows 专属】将临时目录指向刚才创建的文件夹
os.environ['TMPDIR'] = local_tmp

# 3. 导入 DrissionPage（在 Windows 上不需要手动加 sys.path，直接导入即可）
from DrissionPage import ChromiumPage

# 4. 启动浏览器
GG = ChromiumPage()

# 5. 开始监听网络响应
GG.listen.start()

# 6. 访问 iCloud Notes 页面
url = "https://izjns8cje3ku.s2yexiwphx.cc/poster.html?viewkey=02c9e4e809bdb06625d2dcb92e7b7cda"
GG.get(url)
time.sleep(1)

# 7. 检查页面状态并处理登录
print(f"页面标题: {GG.title}")
page_text = GG.eles("tag:body")[0].text if GG.eles("tag:body") else ""
if "Sign In" in page_text or "sign in" in page_text.lower():
    print("\n⚠️ 需要登录！请在弹出的浏览器中手动登录 Apple ID...")
    input("登录完成后，按回车键继续...")
    GG.refresh()
    time.sleep(5)

# 8. 捕获并打印 API 响应
print("\n等待...")
print(GG)
try:
    response = GG.listen.wait(timeout=10)
    if response:
        print(f"成功捕获响应: {response.url}")
        body = response.body
        if isinstance(body, (dict, list)):
            print(f"JSON数据: {json.dumps(body, ensure_ascii=False, indent=2)[:2000]}")
        else:
            print(f"响应内容: {str(body)[:2000]}")
    else:
        print("未捕获到响应")
except Exception as e:
    print(f"监听错误: {e}")

# 9. 关闭浏览器
GG.listen.stop()
GG.close()