# 筑波大学manaba提交统计工具

这个Python脚本用于统计筑波大学manaba系统中学生的提交情况，包括提交次数、出勤次数和口头提交次数。

## 功能特点

- 统计指定学生的提交次数
- 统计不同日期的出勤次数
- 检查口头提交情况
- 导出Excel格式的统计结果
- 支持保存HTML内容用于调试

## 使用前准备

1. **安装必要的Python包**

```bash
pip install requests pandas beautifulsoup4
```

2. **导出Chrome浏览器的Cookies**
   - 登录manaba系统
   - 安装Chrome扩展：[EditThisCookie](https://chrome.google.com/webstore/detail/editthiscookie/fngmhnnpilhplaeedifhccceomclgfbg)
   - 点击扩展图标，选择"导出Cookies"
   - 将导出的内容保存为 `cookies.json` 文件，放在脚本同目录下

3. **修改学生名单**
   - 在 `process_submission_table` 函数中的 `student_list` 修改为你需要统计的学生名单：

```python
student_list = [
    "DD　XX", "AA　BC",
]
```

## 使用方法

```bash
python main.py
``` 


2. 脚本会自动：
   - 读取cookies.json
   - 访问manaba系统
   - 获取所有提交页面
   - 检查每个学生的提交情况
   - 生成统计结果

3. 输出文件：
   - `提交情况统计.xlsx`：包含所有统计结果的Excel文件
   - `page.html`：最后一次访问的页面HTML内容（用于调试）
   - `detail_page.html`：最后一次访问的详细页面HTML内容（用于调试）

## 统计结果说明

Excel文件包含以下列：
- 姓名：学生姓名
- 出勤次数：不同日期的提交次数
- 口头提交：口头提交的次数
- 提交次数：总提交次数
- 提交日期：所有提交日期的列表

## 注意事项

- 确保cookies.json文件是最新的
- 如果遇到连接问题，可能需要更新cookies
- 默认只处理前10个链接，可以修改代码中的限制进行完整统计
- 请遵守学校的使用政策和隐私规定

## 调试

如果遇到问题：
1. 检查生成的HTML文件内容
2. 确认cookies是否有效
3. 查看控制台输出的错误信息

## 贡献

欢迎提交Issue和Pull Request来改进这个工具。
