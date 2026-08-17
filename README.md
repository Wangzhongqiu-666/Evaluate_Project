# B站教学视频评论情感分析

这是一个用于整理和分析 B站教学视频评论的 Python 课程项目。程序可以采集视频的顶层评论，也可以直接读取已有的 Excel 数据，完成文本清洗、中文分词、情感分类、关键词统计和分析报告生成。

## 实现内容

- 根据 BV/av 视频地址采集顶层评论，并保存昵称、性别和评论内容；
- 合并一个或多个 Excel 文件中的评论数据；
- 清理评论文本，使用自定义词典和停用词表进行分词；
- 使用针对教学评价训练的 SnowNLP 模型计算情感分数；
- 按情感分数将评论分为正面、中性和负面；
- 统计高频关键词，生成词云、情绪分布图和情感分数分布图；
- 生成包含评论统计、典型评论和关键词统计的 Word 报告，并在 Windows 上尝试导出 PDF。

情感分类使用以下区间：

- 分数大于或等于 `0.7`：正面；
- 分数小于或等于 `0.3`：负面；
- 其余分数：中性。

## 项目目录

```text
Evaluate_Project/
├─ Codings/             # 评论采集、数据处理、模型训练和报告生成代码
├─ Functional_files/    # 自定义词典、停用词表、训练语料和 SnowNLP 模型
├─ .env.example         # Cookie 环境变量示例
├─ requirements.txt     # Python 依赖
└─ 双击在线采集.bat     # Windows 在线采集入口
```

`Original_Datas/` 和 `outputs/` 用于存放本地数据与运行结果，默认不会提交到 Git 仓库。

## 运行环境

项目在 Windows 和 Python 3.13 环境下开发。安装依赖：

```powershell
python -m pip install -r requirements.txt
```

PDF 导出依赖本机安装的 Microsoft Word。没有 Word 时，图表和 DOCX 文件仍可生成。

## 准备数据

离线分析使用 Excel 文件作为输入，表格需要包含以下三列：

| 昵称 | 性别 | 评论 |
| --- | --- | --- |
| 示例用户 | 保密 | 老师讲得很清楚 |

可以把文件保存为 `Original_Datas/评论.xlsx`，也可以在运行时通过 `--input` 指定其他文件或目录。输入目录中有多个 Excel 文件时，程序会先合并数据再进行处理。

## 使用方法

### 分析已有评论

在项目根目录运行：

```powershell
python Codings/main.py offline --input "Original_Datas/评论.xlsx" --output "outputs/result"
```

程序会在输出目录中保存处理后的评论、关键词统计和分析摘要。

### 生成报告

完成数据处理后运行：

```powershell
python Codings/main.py report --output "outputs/result"
```

程序会生成情绪分布图、情感分数分布图、关键词图、词云和 Word 报告；如果本机可以调用 Microsoft Word，还会同时导出 PDF。

### 在线采集评论

Windows 下可以双击 `双击在线采集.bat`，按提示输入视频地址和 Cookie。采集的数据会保存到 `Original_Datas/`。

也可以在 PowerShell 中运行：

```powershell
$env:BILIBILI_COOKIE = "自己的 B站 Cookie"
python Codings/main.py crawl --url "B站视频地址" --input "Original_Datas/评论.xlsx" --output "outputs/result"
```

评论采集依赖 B站接口，接口或签名规则变化时可能无法正常使用。Cookie 只应保存在本机环境变量中，不要写入代码或提交到仓库。

### 重新训练情感模型

仓库已经包含训练后的模型。如需使用 `Functional_files/针对评教的SnowNLP模型/` 中的正负语料重新训练，可以运行：

```powershell
python Codings/snownlp评教模型训练.py
```

## 输出内容

项目的主要输出包括：

- 带情感分数和分类结果的评论数据；
- 关键词及其出现频数；
- 正面、中性、负面评论分布图；
- 情感分数分布图和关键词词云；
- DOCX 格式的评论分析报告；
- 在 Microsoft Word 可用时导出的 PDF 报告。

SnowNLP 的分数用于辅助整理评论，不能代替人工判断。反问、讽刺或依赖上下文的评论可能出现分类偏差。
