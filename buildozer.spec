[app]
# 应用信息
title = 闲鱼竞价商品库
package.name = xianyu_price_monitor
package.domain = org.yourname
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json
# 版本
version = 1.0
# 需求
requirements = python3, requests, tkinter, json, re
# 方向设置
orientation = portrait
# 屏幕适配
fullscreen = 0
# 支持Android版本
android.api = 33
android.minapi = 21
android.ndk = 25.2.9519653
# 应用图标（如没有则用默认）
# icon.filename = icon.png
# 额外python模块
# p4a.branch = develop
p4a.whitelist = *
p4a.extra_packages = requests
# 入口
entrypoint = price_monitor_editor::main
