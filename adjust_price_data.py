import pandas as pd
import numpy as np

# 读取原始数据
df = pd.read_excel('价格数据模版.xlsx')

# 计算涨跌幅
df['涨跌幅'] = df['收盘价'].pct_change() * 100

# 保存原始收盘价
original_close = df['收盘价'].copy()

def adjust_prices(row):
    close_price = row['收盘价']
    price_change_pct = row['涨跌幅'] if not pd.isna(row['涨跌幅']) else 0
    
    # 根据涨跌幅调整价格波动范围
    if abs(price_change_pct) > 3:  # 大幅波动（超过3%）
        price_range = close_price * 0.015  # 扩大到1.5%基础波动
        # 在大涨的情况下，倾向于开盘价低于收盘价
        if price_change_pct > 3:
            open_diff = np.random.uniform(-price_range, -price_range/3)  # 开盘价偏向低位
            extra_range = 0.008  # 额外0.8%的波动空间
        # 在大跌的情况下，倾向于开盘价高于收盘价
        elif price_change_pct < -3:
            open_diff = np.random.uniform(price_range/3, price_range)  # 开盘价偏向高位
            extra_range = 0.008
        else:
            open_diff = np.random.uniform(-price_range, price_range)
            extra_range = 0.005
    elif abs(price_change_pct) > 1:  # 中等波动（1%-3%）
        price_range = close_price * 0.008  # 0.8%基础波动
        if price_change_pct > 1:
            open_diff = np.random.uniform(-price_range, 0)  # 开盘价偏向低位
            extra_range = 0.005
        else:
            open_diff = np.random.uniform(0, price_range)  # 开盘价偏向高位
            extra_range = 0.005
    else:  # 小幅波动（1%以内）
        price_range = close_price * 0.005  # 0.5%基础波动
        open_diff = np.random.uniform(-price_range/2, price_range/2)
        extra_range = 0.003
    
    open_price = close_price + open_diff
    
    # 根据涨跌幅调整最高最低价的范围
    if price_change_pct > 0:  # 上涨
        # 确保最高价有足够空间反映上涨趋势
        high_range = close_price * (extra_range + abs(price_change_pct) * 0.003)  # 随涨幅增加而增加
        low_range = close_price * extra_range
        high_price = max(open_price, close_price) + abs(np.random.uniform(0, high_range))
        low_price = min(open_price, close_price) - abs(np.random.uniform(0, low_range))
    elif price_change_pct < 0:  # 下跌
        # 确保最低价有足够空间反映下跌趋势
        high_range = close_price * extra_range
        low_range = close_price * (extra_range + abs(price_change_pct) * 0.003)  # 随跌幅增加而增加
        high_price = max(open_price, close_price) + abs(np.random.uniform(0, high_range))
        low_price = min(open_price, close_price) - abs(np.random.uniform(0, low_range))
    else:  # 平盘
        high_price = max(open_price, close_price) + abs(np.random.uniform(0, close_price * extra_range))
        low_price = min(open_price, close_price) - abs(np.random.uniform(0, close_price * extra_range))
    
    # 确保价格不会出现负数，并且保持适当的精度
    return pd.Series({
        '开盘价': round(max(open_price, 0), 3),
        '最高价': round(max(high_price, 0), 3),
        '最低价': round(max(low_price, 0), 3)
    })

# 应用调整
np.random.seed(42)  # 设置随机种子以确保结果可重现
adjusted = df.apply(adjust_prices, axis=1)

# 更新数据框
df['开盘价'] = adjusted['开盘价']
df['最高价'] = adjusted['最高价']
df['最低价'] = adjusted['最低价']

# 确保收盘价保持不变
df['收盘价'] = original_close

# 保存调整后的数据
df.to_excel('价格数据模版_调整后.xlsx', index=False)

print("数据已调整完成，新文件已保存为'价格数据模版_调整后.xlsx'") 