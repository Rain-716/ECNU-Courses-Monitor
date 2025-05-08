# 课程人数监控脚本

## 项目简介
该脚本基于 Selenium 和 pandas 自动登录华东师范大学教务系统，定时获取课程选课人数信息，并将课程详情与当前人数、限额人数及剩余席位导出到 Excel 文件。可用于实时监控课程余量，帮助学生及时抢课。

## 功能
- 自动打开浏览器并登录教务系统
- 抓取两份 JS 源数据：课程详情（`stdElectCourse!data.js`）和选课人数（`stdElectCourse!queryStdCount.js`）
- 解析并合并课程信息与选课人数数据
- 计算剩余可选席位
- 导出为格式化的 Excel（`lesson_overview_with_counts.xlsx`）
- 支持定时循环抓取（默认 60 秒间隔）

## 环境依赖
- Python 3.7+
- Chrome 浏览器及对应版本的 [ChromeDriver](https://chromedriver.chromium.org/)
- 以下 Python 库：
  - `selenium`
  - `pandas`
  - `openpyxl`

可通过以下命令一次性安装：
```bash
pip install selenium pandas openpyxl
```

## 配置

在脚本开头部分修改以下常量：

```python
USERNAME = '你的学号或用户名'
PASSWORD = '你的教务系统登录密码'
LOGIN_URL = 'https://applicationnewjw.ecnu.edu.cn/eams/stdElectCourse.action'
TARGET_URL = 'https://applicationnewjw.ecnu.edu.cn/eams/stdElectCourse!queryStdCount.action?profileId=6022'
JS_FILE_COUNTS = 'stdElectCourse!queryStdCount.js'
JS_FILE_LESSONS = 'stdElectCourse!data.js'
OUTPUT_EXCEL = 'lesson_overview_with_counts.xlsx'
INTERVAL_SECONDS = 60  # 循环抓取时间间隔（秒）
```

## 使用说明

1. **下载并配置 ChromeDriver**

   * 确保本地安装的 ChromeDriver 与 Chrome 浏览器版本对应，并将其路径加入环境变量。
2. **填写账号信息**

   * 在脚本顶部 `USERNAME`、`PASSWORD` 常量处填写你的教务系统账号密码。
3. **运行脚本**

   ```bash
   python monitor_courses.py
   ```
4. **输入验证码并继续**

   * 脚本会暂时停在命令行，等待你手动输入验证码并按回车。
5. **自动抓取并导出**

   * 程序每隔设定的 `INTERVAL_SECONDS` 秒会自动重复抓取并更新 JS 文件、解析数据、输出最新 Excel 文件。

## 脚本结构

* `load_js_json(file_path)`
  解析本地 JS 文件内容，将其中的 JSON 提取并返回字符串。
* `run_task()`

  * 利用 Selenium 登录并下载两份 JS 源文件
  * 调用 `load_js_json` 解析课程详情和人数数据
  * 使用 `pandas` 合并、计算并输出 Excel
* `__main__`
  进入循环，每隔 `INTERVAL_SECONDS` 秒调用一次 `run_task()`。

## 注意事项

* 请勿在高频率（<30 秒）下运行，以免对教务系统造成过大压力。
* 导出的 Excel 文件会每次覆盖，如需保留历史可自行重命名或修改脚本。