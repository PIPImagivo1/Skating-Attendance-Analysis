# 🛼 轮滑活动智能分析系统

![Python][image-1]
![Pandas][image-2]
![License][image-3]

自动化处理轮滑社团活动数据，实现：
- 📅 多日期考勤合并
- 👤 会员信息智能匹配
- 📊 可视化数据报告生成

## 功能亮点
- **智能纠错**：自动处理格式不一致的序号
- **跨平台支持**：兼容Excel和CSV文件
- **数据安全**：原始文件只读，结果自动归档

## 快速开始
### 安装依赖
```bash
pip install -r src/requirements.txt
```

### 运行程序
```bash
python src/main.py
```

### 文件结构要求
```text
data/raw/
├── member_info.xlsx    # 会员表（需包含'序号'和'姓名'列）
└── 2025.03.31.xlsx    # 考勤文件（两列：序号、签退状态）
```

## 可视化演示
![运行演示][image-4]

## 开发文档
[查看详细技术文档][1]

## 开源协议
本项目采用 [MIT License][2]

[1]:	docs/DEVELOPMENT.md
[2]:	LICENSE

[image-1]:	https://img.shields.io/badge/Python-3.10%2B-blue
[image-2]:	https://img.shields.io/badge/Pandas-2.0-lightgrey
[image-3]:	https://img.shields.io/badge/License-MIT-green
[image-4]:	docs/demo.gif