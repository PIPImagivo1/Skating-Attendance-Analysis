# -*- coding: utf-8 -*-
"""
轮滑活动统计系统最终版
适配路径：/Users/pite/Desktop/SkatingProject/
功能：
1. 自动合并多个日期的考勤文件
2. 关联会员信息表
3. 生成带人名的统计报告
"""

import pandas as pd
from pathlib import Path
import warnings
from datetime import datetime

# 禁用openpyxl警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ===================== 配置区域 =====================
# 必须修改：确认以下路径与你的实际结构一致
BASE_DIR = Path("/Users/pite/Desktop/SkatingProject")
RAW_DATA_DIR = BASE_DIR / "data/raw"                    # 原始考勤文件目录
MEMBER_FILE = RAW_DATA_DIR / "轮滑协会24-25会员上学期信息统计.xlsx"  # 会员表路径
PROCESSED_DIR = BASE_DIR / "data/processed"             # 结果输出目录
# ==================================================

def validate_paths():
    """验证关键路径是否存在"""
    errors = []
    if not RAW_DATA_DIR.exists():
        errors.append(f"原始数据目录不存在：{RAW_DATA_DIR}")
    if not MEMBER_FILE.exists():
        errors.append(f"会员表不存在：{MEMBER_FILE}")
    
    if errors:
        print("❌ 路径配置错误：")
        for error in errors:
            print(error)
        print("\n请检查：")
        print(f"1. 确认桌面存在 SkatingProject 文件夹")
        print(f"2. 确认 data/raw/ 内有考勤文件和会员表")
        exit(1)

def load_member_info():
    """
    加载会员信息表
    返回：{序号: 姓名} 的字典
    """
    try:
        # 假设会员表有标题行，序号在首列，姓名在第二列
        df = pd.read_excel(MEMBER_FILE, usecols=[0,1], names=['序号', '姓名'], header=0)
        return dict(zip(df['序号'], df['姓名']))
    except Exception as e:
        print(f"❌ 会员表读取失败：{MEMBER_FILE}")
        print(f"错误详情：{str(e)}")
        print("\n请检查：")
        print("1. 文件是否被其他程序打开")
        print("2. 表格是否包含'序号'和'姓名'两列")
        exit(1)

def process_attendance_files(member_dict):
    """
    处理所有考勤文件
    返回：{序号: 参与次数} 的字典
    """
    attendance = {}
    
    for file in RAW_DATA_DIR.glob("20*.xlsx"):  # 只处理以20开头的日期文件
        if "轮滑协会" in file.name:  # 跳过会员表
            continue
            
        try:
            # 读取无标题行的考勤文件
            df = pd.read_excel(file, header=None, names=['序号', '签退状态'])
            
            # 提取日期（从文件名）
            date_str = file.stem.replace(".", "-")  # 2025.3.31 → 2025-3-31
            print(f"📅 正在处理 {date_str} 的数据...")
            
            for _, row in df.iterrows():
                if pd.notna(row['签退状态']) and row['签退状态'] in [1, 2]:
                    serial = int(row['序号'])
                    attendance[serial] = attendance.get(serial, 0) + 1
                    
        except Exception as e:
            print(f"⚠️ 文件处理失败：{file.name}")
            print(f"错误原因：{str(e)}")
            continue
    
    return attendance

def generate_report(member_dict, attendance_data):
    """生成最终统计报告"""
    report = []
    
    for serial, count in attendance_data.items():
        name = member_dict.get(serial, f"未知会员_{serial}")
        report.append({
            '姓名': name,
            '会员序号': serial,
            '参与次数': count,
            '应减打卡': count * 2
        })
    
    # 按参与次数降序排序
    df = pd.DataFrame(report).sort_values('参与次数', ascending=False)
    
    # 保存结果
    PROCESSED_DIR.mkdir(exist_ok=True)  # 确保输出目录存在
    output_file = PROCESSED_DIR / "轮滑活动统计_最终版.xlsx"
    df.to_excel(output_file, index=False)
    
    return df, output_file

def main():
    print("="*50)
    print("轮滑活动统计系统 v2.0")
    print("="*50 + "\n")
    
    # 步骤1：路径验证
    print("🔍 正在验证文件路径...")
    validate_paths()
    
    # 步骤2：加载会员信息
    print("\n📋 正在加载会员表...")
    member_info = load_member_info()
    print(f"✅ 已加载 {len(member_info)} 位会员信息")
    
    # 步骤3：处理考勤数据
    print("\n📊 正在分析考勤文件...")
    attendance_data = process_attendance_files(member_info)
    print(f"✅ 已处理 {len(attendance_data)} 位参与记录")
    
    # 步骤4：生成报告
    print("\n📑 正在生成统计报告...")
    final_report, output_path = generate_report(member_info, attendance_data)
    
    # 结果展示
    print("\n🎉 处理完成！")
    print(f"📂 结果文件位置：{output_path}")
    print("\n📋 统计结果预览：")
    print(final_report.head())
    
    # 保存处理日志
    log_file = PROCESSED_DIR / "processing_log.txt"
    with open(log_file, "w") as f:
        f.write(f"最后处理时间：{datetime.now()}\n")
        f.write(f"处理文件数量：{len(list(RAW_DATA_DIR.glob('20*.xlsx')))}个\n")
        f.write(f"有效参与人次：{sum(attendance_data.values())}次\n")
    
    print(f"\n⏱ 本次运行日志已保存至：{log_file}")

if __name__ == "__main__":
    main()