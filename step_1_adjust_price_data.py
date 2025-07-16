import pandas as pd
import numpy as np
from datetime import datetime

# 设置需要模拟的日期范围
start_date = '2025-07-01'  # 开始日期
end_date = '2026-06-30'    # 结束日期

# 价格整体调整变量（单位：元）
# 正数表示上调，负数表示下调，0表示不调整
PRICE_ADJUSTMENT = 0 # 例如：-0.5 表示所有价格下调0.5元

# 排除调整的产品列表
# 这些产品的价格数据将不会被调整，保持原始值
EXCLUDE_PRODUCTS = [" "]  # 可以添加多个产品，如 ["牛腩", "大米龙", "牛霖"]

# 读取原始数据
df = pd.read_excel('价格数据模版.xlsx')

# 确保日期列是datetime格式
df['日期'] = pd.to_datetime(df['日期'])

# 计算涨跌幅
df['涨跌幅'] = df['收盘价'].pct_change() * 100

# 保存所有原始数据
original_data = df.copy()

def adjust_prices(row):
    close_price = row['收盘价']  # 收盘价必须保持不变
    open_price = row['开盘价']
    price_change_pct = row['涨跌幅'] if not pd.isna(row['涨跌幅']) else 0
    
    # 检查开盘价是否需要模拟或修正
    simulate_open = False
    
    if pd.isna(open_price):
        # 情况1：开盘价缺失
        simulate_open = True
    elif open_price <= 0:
        # 情况2：开盘价为0或负数
        simulate_open = True
    else:
        # 情况3：检查是否为离群点
        price_diff_pct = abs(open_price - close_price) / close_price
        
        # 离群点判断标准
        if price_diff_pct > 0.01:  # 开盘价与收盘价差异超过1%（日内波动限制）
            print(f"  检测到日内波动过大：{open_price:.3f}（与收盘价{close_price:.3f}差异{price_diff_pct:.1%}），将重新模拟")
            simulate_open = True
        elif open_price > close_price * 2 or open_price < close_price * 0.5:  # 开盘价是收盘价的2倍以上或一半以下
            print(f"  检测到异常开盘价：{open_price:.3f}（收盘价：{close_price:.3f}），将重新模拟")
            simulate_open = True
    
    # 如果需要模拟开盘价
    if simulate_open:
        # 根据涨跌幅和收盘价推算合理的开盘价（严格限制在1%以内）
        if abs(price_change_pct) > 0:
            # 有趋势时，开盘价与收盘价的关系应该合理
            if price_change_pct > 0:  # 上涨趋势
                # 开盘价应该低于收盘价，但波动不超过1%
                open_variation = np.random.uniform(0, min(abs(price_change_pct) * 0.4, 1)) / 100  # 最多1%的差异
                open_price = close_price * (1 - open_variation)
            else:  # 下跌趋势
                # 开盘价应该高于收盘价，但波动不超过1%
                open_variation = np.random.uniform(0, min(abs(price_change_pct) * 0.4, 1)) / 100  # 最多1%的差异
                open_price = close_price * (1 + open_variation)
        else:
            # 无明显趋势时，开盘价在收盘价附近小幅波动（限制在±1%）
            open_variation = np.random.uniform(-0.01, 0.01)  # ±1%的波动
            open_price = close_price * (1 + open_variation)
        
        # 严格确保日内波动不超过1%
        open_price = max(open_price, close_price * 0.99)  # 确保开盘价不低于收盘价的99%
        open_price = min(open_price, close_price * 1.01)  # 确保开盘价不高于收盘价的101%
    
    # 根据开盘价和收盘价的关系确定趋势
    if close_price > open_price:
        trend = "up"  # 上涨
        trend_strength = (close_price - open_price) / open_price
    elif close_price < open_price:
        trend = "down"  # 下跌
        trend_strength = (open_price - close_price) / open_price
    else:
        trend = "flat"  # 平盘
        trend_strength = 0
    
    # 根据趋势强度确定波动范围
    base_volatility = 0.003  # 基础波动率0.3%
    if trend_strength > 0.03:  # 大幅波动（超过3%）
        volatility = base_volatility + 0.008
    elif trend_strength > 0.01:  # 中等波动（1%-3%）
        volatility = base_volatility + 0.005
    else:  # 小幅波动（1%以内）
        volatility = base_volatility + 0.002
    
    # 按照金融逻辑模拟最高价和最低价
    if trend == "up":  # 上涨趋势
        # 最高价：通常高于收盘价
        high_extra = np.random.uniform(0, volatility * close_price * 1.5)
        high_price = close_price + high_extra
        # 最低价：通常接近或低于开盘价
        low_reduction = np.random.uniform(0, volatility * open_price)
        low_price = open_price - low_reduction
        
    elif trend == "down":  # 下跌趋势
        # 最高价：通常接近或高于开盘价
        high_extra = np.random.uniform(0, volatility * open_price)
        high_price = open_price + high_extra
        # 最低价：通常低于收盘价
        low_reduction = np.random.uniform(0, volatility * close_price * 1.5)
        low_price = close_price - low_reduction
        
    else:  # 平盘趋势
        # 最高价和最低价围绕开盘价/收盘价对称分布
        price_center = (open_price + close_price) / 2
        high_extra = np.random.uniform(0, volatility * price_center)
        low_reduction = np.random.uniform(0, volatility * price_center)
        high_price = price_center + high_extra
        low_price = price_center - low_reduction
    
    # 确保金融逻辑正确性
    high_price = max(high_price, max(open_price, close_price))  # 最高价至少等于开盘价和收盘价中的较高者
    low_price = min(low_price, min(open_price, close_price))    # 最低价至多等于开盘价和收盘价中的较低者
    
    # 确保价格不会出现负数，并且保持适当的精度
    result = {
        '开盘价': round(max(open_price, 0), 3),  # 总是返回开盘价（模拟或修正后的）
        '最高价': round(max(high_price, 0), 3),
        '最低价': round(max(low_price, 0), 3)
    }
    
    return pd.Series(result)

# 应用调整（只对指定日期范围内的数据进行模拟）
np.random.seed(42)  # 设置随机种子以确保结果可重现

# 创建日期掩码
date_mask = (df['日期'] >= start_date) & (df['日期'] <= end_date)

# 只对日期范围内的数据进行模拟
adjusted = df[date_mask].apply(adjust_prices, axis=1)

# 更新数据框（只更新日期范围内的数据）
df.loc[date_mask, '开盘价'] = adjusted['开盘价']  # 模拟时间范围内：使用模拟的开盘价
df.loc[date_mask, '最高价'] = adjusted['最高价']
df.loc[date_mask, '最低价'] = adjusted['最低价']

# 对日期范围外的数据保持原值
df.loc[~date_mask, ['开盘价', '最高价', '最低价']] = original_data.loc[~date_mask, ['开盘价', '最高价', '最低价']]

# 确保收盘价绝对不变（对所有数据）
df['收盘价'] = original_data['收盘价']

# 应用价格整体调整（不包括最新日期的数据和排除的产品）
if PRICE_ADJUSTMENT != 0:
    print(f"正在应用价格调整：{PRICE_ADJUSTMENT:+.2f}元")
    
    # 获取最新日期
    latest_date = df['日期'].max()
    
    # 创建掩码：排除最新日期的数据
    non_latest_mask = df['日期'] < latest_date
    
    # 创建掩码：排除指定的产品
    non_excluded_mask = ~df['SKU名称'].isin(EXCLUDE_PRODUCTS)
    
    # 组合掩码：既不是最新日期，也不是排除的产品
    adjustment_mask = non_latest_mask & non_excluded_mask
    
    # 只对最高价和最低价进行调整，保持开盘价和收盘价不变
    price_columns = ['最高价', '最低价']  # 移除开盘价和收盘价
    for col in price_columns:
        df.loc[adjustment_mask, col] = df.loc[adjustment_mask, col] + PRICE_ADJUSTMENT
        # 确保价格不会变成负数
        df.loc[adjustment_mask, col] = df.loc[adjustment_mask, col].clip(lower=0)
    
    # 统计调整情况
    adjusted_rows = adjustment_mask.sum()
    latest_rows = (~non_latest_mask).sum()
    excluded_rows = (~non_excluded_mask).sum()
    
    print(f"价格调整完成：")
    print(f"  - 调整了 {adjusted_rows} 条历史数据的最高价和最低价{'+' if PRICE_ADJUSTMENT > 0 else ''}{PRICE_ADJUSTMENT:.2f}元")
    print(f"  - 保持了 {latest_rows} 条最新数据不变")
    print(f"  - 收盘价保持原始值不变，开盘价按1%波动限制模拟")
    if EXCLUDE_PRODUCTS:
        print(f"  - 排除了 {excluded_rows} 条产品数据不调整（排除产品：{', '.join(EXCLUDE_PRODUCTS)}）")

# 删除涨跌幅列（不保存到输出文件）
if '涨跌幅' in df.columns:
    df = df.drop('涨跌幅', axis=1)

# 保存调整后的数据
df.to_excel('价格数据初步调整.xlsx', index=False)

# 生成最近两天的收盘价比对表格
# 获取最近两天的日期
latest_dates = sorted(df['日期'].unique())[-2:]

# 创建比对数据框
comparison_df = df[df['日期'].isin(latest_dates)][['日期', 'SKU名称', '收盘价']]

# 将数据透视为宽格式
comparison_pivot = comparison_df.pivot(index='SKU名称', columns='日期', values='收盘价')

# 计算涨跌幅
latest_date = latest_dates[-1]
previous_date = latest_dates[-2]
comparison_pivot['涨跌幅'] = ((comparison_pivot[latest_date] - comparison_pivot[previous_date]) / comparison_pivot[previous_date] * 100).round(2)

# 重命名列为更友好的格式
comparison_pivot.columns = [d.strftime('%Y-%m-%d') if isinstance(d, pd.Timestamp) else d for d in comparison_pivot.columns]

# 保存比对表格
comparison_pivot.to_excel('网页的第一个表格的比对.xlsx')

# 打印模拟的日期范围信息
print(f"数据已调整完成，新文件已保存为'价格数据初步调整.xlsx'")
print(f"模拟日期范围：{start_date} 至 {end_date}")
print(f"模拟数据条数：{date_mask.sum()} 条")
print(f"保持原值数据条数：{(~date_mask).sum()} 条")
if PRICE_ADJUSTMENT != 0:
    exclude_info = f"，排除产品：{', '.join(EXCLUDE_PRODUCTS)}" if EXCLUDE_PRODUCTS else ""
    print(f"价格整体调整：最高价和最低价{PRICE_ADJUSTMENT:+.2f}元（仅历史数据{exclude_info}）")
else:
    print("价格整体调整：无调整")
print(f"\n最近两天({latest_dates[-2].strftime('%Y-%m-%d')}至{latest_dates[-1].strftime('%Y-%m-%d')})的收盘价比对表格已保存为'网页的第一个表格的比对.xlsx'") 