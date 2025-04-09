def check_member_consistency(member_dict, data_dir):
    """检查是否有未注册成员参与活动"""
    all_ids = set()
    for file in Path(data_dir).glob("*.xlsx"):
        df = pd.read_excel(file)
        all_ids.update(df['序号'].unique())
    
    missing = all_ids - member_dict.keys()
    if missing:
        print(f"发现{len(missing)}个未注册序号：{missing}")