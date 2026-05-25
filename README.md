# excle_verify_easy

极简版 Excel 模板校验工具，只依赖 `openpyxl`。

## 功能

```text
读取模板示例行并推断规则
生成带规则约束的填写模板
从“序号=1”定位真实数据区
校验数据并生成带时间戳的 output Excel
错误值改为 /，标红并居中
警告值标黄
自动修正值标蓝并写入修正明细
支持日期、年月、年份、年龄、枚举、列表、数值、身份证号、手机号等规则
支持英文分号、空格、末尾句号、字体颜色等基础规范化
```

## 运行

```bash
pip install -r requirements.txt
```

## 配置

公共配置统一放在 `config/` 目录：

```text
config/app_config.json          运行配置：模板、数据、sheet、表头行、示例行、输出目录
config/auto_match_rules.json    候选清单匹配配置
```

`settings.py` 只作为兼容入口读取这些配置，并继续提供：

```text
TEMPLATE_PATH
DATA_PATH
SHEET_NAME
HEADER_ROWS
SAMPLE_ROW
AUTO_MATCH_CONFIG_PATH
```

`app_config.json` 示例：

```json
{
  "template_path": "input/template/员工信息（市城规总院）2.xlsx",
  "data_path": "input/data/员工信息（市城规总院）.xlsx",
  "sheet_name": "auto",
  "header_rows": [1, 2],
  "sample_row": 3,
  "output": {
    "data_dir": "output/data",
    "template_dir": "output/template",
    "data_prefix": "output",
    "template_prefix": "填写模板"
  }
}
```

只生成带规则约束的填写模板：

```bash
python generate_template/main.py
```

只校验数据并生成修正后的校验结果：

```bash
python verify/main.py
```

兼容旧流程，同时执行“生成填写模板 + 校验数据”：

```bash
python main.py
```

输出目录：

```text
output/template/填写模板_模板名_YYYYMMDD_HHMMSS_microseconds.xlsx
output/data/output_数据文件名_YYYYMMDD_HHMMSS_microseconds.xlsx
```

## 候选清单匹配

候选清单匹配默认关闭，配置文件在：

```text
config/auto_match_rules.json
```

人工只维护标准值清单，不需要维护 aliases：

```json
{
  "enabled": true,
  "rules": [
    {
      "name": "公司名称清单",
      "field_keywords": ["公司", "单位"],
      "auto_replace_threshold": 0.88,
      "min_score_gap": 0.12,
      "candidates": [
        "小米公司",
        "华为公司",
        "城建集团",
        "一建公司",
        "交投集团"
      ]
    }
  ]
}
```

开启后，只会对命中 `field_keywords` 的字段生效。

处理逻辑：

```text
填写值在 candidates 中：不修改
填写值不在 candidates 中，但可信匹配到唯一候选项：自动修正
填写值不在 candidates 中，且匹配过低或候选项歧义：改为 /，标红居中
```

生成填写模板时，如果某列字段命中候选清单规则，该列会自动生成下拉选择约束，只允许从 `candidates` 中选择标准值。

错误和修正说明会保持简洁，不输出 score/gap。

当前算法组合：

```text
标准化
SequenceMatcher
字符 Jaccard
bigram Jaccard
编辑距离相似度
top1/top2 歧义保护
```

## 规则边界

```text
模板必须有示例行
数据字段顺序默认和模板一致
字段类型主要由模板示例值决定
枚举和列表逻辑保持原有规则
候选清单匹配必须人工开启并配置标准候选项
```
