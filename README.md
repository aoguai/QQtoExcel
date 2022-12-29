# QQtoExcel

[![MIT License](https://img.shields.io/badge/License-MIT-green.svg)](https://choosealicense.com/licenses/mit/)

一个让PC QQ 导出TXT聊天记录转Excel表格的工具

A tool for PC QQ to export TXT chat records to excel tables

## 文档

详细说明、使用方法、参数介绍 请看wiki

[Wiki](https://github.com/aoguai/QQtoExcel/wiki)

## 部署与使用

### 部署

1. clone本项目到本地

2. 手动从QQ消息管理器中导出需要转换的消息，注意改为UTF-8-BOM

3. 运行 CIL 相关命令即可

示例：

```bash
  QQtoExcel 全部消息记录.txt
```

### GUI 与 可执行文件说明

**如果您是windows用户，没有浏览项目代码需求**
可以前往[下载页面](https://github.com/aoguai/QQtoExcel/releases)下载 解压后得到

- releases.exe ：**普通用户使用，不支持 CLI ，有简单 CMD 界面与引导**
- QQtoExcel.exe ：一般开发者使用，仅支持 CLI

同时你可以前往 [QQtoExcel_GUI](https://github.com/abyss-zues/QQtoExcel_GUI) 下载GUI版本

## 贡献

您可以直接在 [issues](https://github.com/aoguai/QQtoExcel/issues) 中提出您的问题或通过 [提交PR](https://github.com/aoguai/QQtoExcel/pulls)
贡献您的代码

## 相关项目

以下是一些相关项目

[QQtoExcel_GUI](https://github.com/abyss-zues/QQtoExcel_GUI) ： 一个QQtoExcel的GUI界面版本

## 演示

![1](https://github.com/aoguai/QQtoExcel/blob/main/images/1.png)
![2](https://github.com/aoguai/QQtoExcel/blob/main/images/2.png)
![3](https://github.com/aoguai/QQtoExcel/blob/main/images/3.png)
![4](https://github.com/aoguai/QQtoExcel/blob/main/images/4.png)

## 开发规划

### 规划

- [x] 支持好友/群聊/全部聊天记录 转换导出
- [x] 支持可选项 选择导出
- [x] 增加 消息分组 可选项，可按分组导出
- [x] 支持 聊天记录清洗，去除无效聊天记录
- [ ] 支持 多工作表 导出
- [ ] 支持 自定义导出文件名规则

### 更新日志

- **2022/12/29 QQtoExcelV1.6.0版本更新**
    - 新增 过滤无意义内容 可选项，可去除无效聊天记录
    - 修复 导出文件使用Excel打开错误 BUG
    - 修复 部分消息被错误分割、无法导出问题
    - 优化 代码结构，使代码更加规范
    - 更新 新匹配消息正则表达式，更精准的分割聊天记录
    - 更新README
- **2022/8/1 QQtoExcelV1.5.0版本更新**
    - 新增 消息分组 可选项，可按分组导出
    - 修复 打包程序在 windows7 不可用的情况
    - 更新后支持 CIL
    - 优化 代码结构
- **2022/7/31 更新README**
- **2022/7/19 QQtoExcelV1.1.0版本更新**
    - 新增 自定义可选项标题 功能
    - 新增 操作流程一些细节显示
    - 优化 匹配正则表达式
    - 修复 消息对象分割错误BUG

## 免责声明

此存储库遵循 MIT 开源协议，请务必理解。

我们严禁所有通过本程序违反任何国家法律的行为，请在法律范围内使用本程序。

默认情况下，使用此项目将被视为您同意我们的规则。请务必遵守道德和法律标准。

如果您不遵守，您将对后果负责，作者将不承担任何责任！