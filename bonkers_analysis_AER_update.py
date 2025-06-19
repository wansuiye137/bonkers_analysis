import pandas as pd
import numpy as np
from pathlib import Path
import re
import os
import glob
from datetime import datetime

def update_aer_history(new_csv_file, output_dir='bonkers_analysis'):
    # 1. 使用通配符查找所有历史文件
    history_files = sorted(glob.glob(str(Path(output_dir) / 'AER_history_*.xlsx')))

    # 2. 处理历史文件存在/不存在的情况
    if not history_files:
        print("未找到历史文件，创建新的数据框...")
        df_history = pd.DataFrame(columns=['Type', 'Account', 'Bank', 'TermMonths'])
        current_date = None
    else:
        # 获取最新的历史文件
        latest_history = history_files[-1]
        print(f"找到最新历史文件: {os.path.basename(latest_history)}")

        # 提取当前历史文件日期
        match = re.search(r'AER_history_(\d{4}-\d{2}-\d{2})\.xlsx', latest_history)
        if not match:
            print(f"无法从文件名提取日期: {latest_history}")
            return
        current_date = match.group(1)

        # 读取历史数据并保留所有列
        df_history = pd.read_excel(latest_history)
        print(f"历史数据形状: {df_history.shape}")

    # 3. 创建key列用于匹配
    df_history['key'] = df_history[['Type', 'Account', 'Bank', 'TermMonths']].astype(str).apply('|'.join, axis=1)

    # 4. 处理新CSV文件
    df_new = pd.read_csv(new_csv_file)
    new_date = pd.to_datetime(df_new['RunDate'].iloc[0]).date()
    new_date_str = new_date.strftime('%Y-%m-%d')
    df_new['key'] = df_new[['Type', 'Account', 'Bank', 'TermMonths']].astype(str).apply('|'.join, axis=1)

    # 5. 处理重复key
    if df_new['key'].duplicated().any():
        df_new = df_new.drop_duplicates(subset='key', keep='first')

    # 6. 创建新日期列数据
    new_data = df_new.set_index('key')

    # 7. 合并新数据到历史数据
    df_history = df_history.set_index('key')

    # 创建新日期列
    new_col_name = str(new_date_str)
    df_history[new_col_name] = np.nan

    # 更新新日期列的值
    for key in df_history.index:
        if key in new_data.index:
            df_history.loc[key, new_col_name] = new_data.loc[key, 'AER']

    # 添加新账户
    new_keys = set(new_data.index) - set(df_history.index)
    if new_keys:
        new_accounts = new_data.loc[list(new_keys)].copy()
        new_accounts[new_col_name] = new_accounts['AER']
        df_history = pd.concat([df_history, new_accounts], axis=0)

    # 8. 重置索引并清理
    df_history.reset_index(inplace=True)
    df_history.drop(columns=['key'], inplace=True)

    # 9. 重新排序列
    info_cols = ['Type', 'Account', 'Bank', 'TermMonths']

    # 获取所有日期列（包括历史日期列）
    date_cols = [col for col in df_history.columns if re.match(r'^\d{4}-\d{2}-\d{2}$', str(col))]
    date_cols = sorted(date_cols)  # 按日期排序
    other_cols = [col for col in df_history.columns if col not in info_cols + date_cols]

    # 重新排列列顺序
    df_history = df_history[info_cols + other_cols + date_cols]

    # 10. 创建新文件名（使用新日期）
    new_filename = f"AER_history_{new_date_str}.xlsx"
    new_filepath = Path(output_dir) / new_filename

    # 11. 保存新历史文件
    df_history.to_excel(new_filepath, index=False)

    print(f"更新完成! 新增日期: {new_date_str}")
    print(f"数据形状: {df_history.shape}")
    print(f"生成新历史文件: {new_filename}")

#更新文件
update_aer_history('bonkers_2025-06-19.csv')