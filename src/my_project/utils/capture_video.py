import requests # 响应爬取
from DrissionPage import ChromiumPage #谷歌
import re # 正则表达式
import os # 系统操作

#GG=ChromiumPage()             #自动打开网页
#GG.listen.start() # 监听
GG.get("https://www.douyin.com/user/self?from_tab_name=main&modal_id=7635249213525912548")
sjb=GG.listen.wait() # 等待监听
jason=sjb.response.body   #处理数据
print(jason)