import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
import os

# 创建保存结果的目录（如果不存在）
if not os.path.exists('analysis_results'):
    os.makedirs('analysis_results')

# 读取Excel文件
df = pd.read_excel('价格数据模版_调整后.xlsx')

# 将日期列转换为datetime格式
df['日期'] = pd.to_datetime(df['日期'])

# 获取所有唯一的产品名称
products = df['SKU名称'].unique()

# 为每个产品创建图表
for product in products:
    # 筛选当前产品的数据
    product_df = df[df['SKU名称'] == product]
    
    # 计算移动平均线
    product_df['3MA'] = product_df['收盘价'].rolling(window=3).mean()
    product_df['5MA'] = product_df['收盘价'].rolling(window=5).mean()
    
    # 计算Y轴范围
    y_min = product_df['最低价'].min()
    y_max = product_df['最高价'].max()
    y_range = y_max - y_min
    y_min_display = y_min - (y_range * 0.2)  # 最低价减去20%范围
    y_max_display = y_max + (y_range * 0.2)  # 最高价加上20%范围
    
    # 创建图表
    fig = go.Figure()
    
    # 添加K线图
    fig.add_trace(
        go.Candlestick(
            x=product_df['日期'],
            open=product_df['开盘价'],
            high=product_df['最高价'],
            low=product_df['最低价'],
            close=product_df['收盘价'],
            name='K线',
            increasing_line_color='red',
            decreasing_line_color='green'
        )
    )
    
    # 添加3日移动平均线
    fig.add_trace(
        go.Scatter(
            x=product_df['日期'],
            y=product_df['3MA'],
            name='3MA',
            line=dict(color='#FF851B', width=1)
        )
    )
    
    # 添加5日移动平均线
    fig.add_trace(
        go.Scatter(
            x=product_df['日期'],
            y=product_df['5MA'],
            name='5MA',
            line=dict(color='#0074D9', width=1)
        )
    )

    # 获取最新价格和日期
    latest_price = product_df['收盘价'].iloc[-1]
    latest_date = product_df['日期'].iloc[-1]
    
    # 添加最新价格标注
    fig.add_annotation(
        x=latest_date,
        y=latest_price,
        text=f'最新参考成交指数: {latest_price:.3f}<br>{latest_date.strftime("%Y-%m-%d")}',
        showarrow=True,
        arrowhead=2,
        arrowsize=1,
        arrowwidth=2,
        arrowcolor='rgba(0,0,0,0.5)',
        ax=-80,
        ay=-40,
        font=dict(size=12, color='black'),
        bgcolor='white',
        bordercolor='black',
        borderwidth=1,
        align='left'
    )
    
    # 更新布局
    fig.update_layout(
        title=dict(
            text=f'{product}价格走势图',
            x=0.5,
            y=0.95,
            font=dict(size=20)
        ),
        yaxis_title='价格指数',
        xaxis_title='日期',
        template='plotly_white',
        xaxis_rangeslider_visible=False,  # 禁用下方的范围滑块
        height=700,  # 设置图表高度
        plot_bgcolor='white',  # 设置背景颜色
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor='lightgrey',
            fixedrange=True,  # 禁用Y轴缩放
            range=[y_min_display, y_max_display],  # 设置Y轴显示范围
            title='价格指数'
        ),
        xaxis=dict(
            showgrid=True,
            gridcolor='lightgrey',
            type='date',
            tickformat='%Y-%m-%d',  # 设置日期格式
            rangeslider=dict(visible=False)
        ),
        margin=dict(t=100, l=50, r=50, b=50)  # 调整边距
    )
    
    # 生成文件名（使用当前时间戳避免文件重名）
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'analysis_results/price_chart_{product}_{timestamp}.html'
    
    # 保存图表为HTML文件
    fig.write_html(filename)
    print(f'已生成{product}的价格走势图：{filename}') 