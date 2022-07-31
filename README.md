# QQtoExcel

##  这个项目是什么
一个让PC QQ 导出TXT聊天记录转Excel表格的工具

A tool for PC QQ to export TXT chat records to excel tables

## 前言

由于QQ未提供聊天记录导出成Excel表格的功能，同时，QQ自带的消息管理功能BUG频出无人修复。

导致无法合理利用个人的聊天记录数据实现一些有意思的功能，特此开发此项目。

## 功能
1. QQ聊天记录备份
2. 利用导出的QQ聊天记录进行数据分析与统计、构建简单用户画像、数据可视化处理
3. 利用导出的QQ聊天记录生成语料库进行NLP或者聊天机器人模型训练
4. 高效整理与搜索QQ聊天记录

### 特点
- 支持好友/群聊转换导出
- 支持选择导出内容
- 最完善的正则表达式匹配功能，避免非法字符等原因导致的导出崩溃或者数据不准确问题

## 使用
**如果您是windows用户，没有浏览项目代码需求**

可以前往[下载页面](https://github.com/aoguai/QQtoExcel/releases)下载可执行文件，或者前往 [QQtoExcel_GUI](https://github.com/abyss-zues/QQtoExcel_GUI) 下载GUI版本，即可直接运行。
### 流程
1. clone本项目到本地

2. 手动从QQ消息管理器中导出需要转换的消息，注意改为UTF-8-BOM

3. 运行 QQtoExcel.py --> 输入聊天记录txt路径 --> 输入导出路径 --> 导出

### 注意事项
|  规定名称   | 解释  |
|  ----  | ----  |
| 消息分组  | 您的QQ好友分组或QQ群聊分组名称 |
| 消息对象  | 您的QQ好友或QQ群组 |

当前版本下你可以选择导出项有以下：
|  可选择项   | 解释  |
|  ----  | ----  |
| 时间  | 每个消息对象中每条消息的对应时间，格式为：yyyy-mm-dd hh:mm:ss |
| 昵称  | 每个消息对象中每条消息的对应备注，若无备注着可能为空、QQ号、QQ昵称 |
| uid  | 每个消息对象中每条消息的联系方式，可能为QQ号或邮箱。该项在好友消息中可能为空 |
| 内容  | 每个消息对象中每条消息的内容 |

以上 可选项 将作为标题均可自定义

当前版本默认导出文件名为：
"分组_昵称.xls"

同时，由于QQ聊天记录中字符复杂，为了避免导出错误程序将对分组名、昵称、内容等涉及到导出Excel的数据进行 **本地** 预处理。

例如，如检测到您的分组或者昵称存在非法字符将会把非法字符替换为"()"，避免windows系统下文件名规定导致的保存失败。

## 效果
![1](https://github.com/aoguai/QQtoExcel/blob/main/images/1.png)
![2](https://github.com/aoguai/QQtoExcel/blob/main/images/2.png)
![3](https://github.com/aoguai/QQtoExcel/blob/main/images/3.png)

## 开发规划
### 规划
- [x] 支持好友/群聊/全部聊天记录 转换导出
- [x] 支持可选项 选择导出
- [ ] 增加 消息分组 可选项，可按分组导出
- [ ] 支持 多工作表 导出
- [ ] 支持 自定义导出文件名规则
- [ ] 支持 聊天记录清洗，去除无效聊天记录

### 更新日志
- **2022/7/31 更新README**
- **2022/7/19 QQtoExcelV1.1.0版本更新**
  - 新增 自定义可选项标题 功能
  - 新增 操作流程一些细节显示
  - 优化 匹配正则表达式
  - 修复 消息对象分割错误BUG

### 如何参与贡献
您可以直接在 [issues](https://github.com/aoguai/QQtoExcel/issues) 中提出您的问题或通过 [提交PR](https://github.com/aoguai/QQtoExcel/pulls) 贡献您的代码

### 相关项目
- [QQtoExcel_GUI](https://github.com/abyss-zues/QQtoExcel_GUI) ： 一个QQtoExcel的GUI界面版本

## 免责声明
此存储库遵循 MIT 开源协议，请务必理解。

我们严禁所有通过本程序违反任何国家法律的行为，请在法律范围内使用本程序。

默认情况下，使用此项目将被视为您同意我们的规则。请务必遵守道德和法律标准。

如果您不遵守，您将对后果负责，作者将不承担任何责任！
