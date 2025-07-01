import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
import os

# 创建保存结果的目录（如果不存在）
if not os.path.exists('analysis_results'):
    os.makedirs('analysis_results')

# 读取价格数据和成交量数据
df_price = pd.read_excel('价格数据初步调整.xlsx')
df_volume = pd.read_excel('成交量数据模板.xlsx')

# 将日期列转换为datetime格式
df_price['日期'] = pd.to_datetime(df_price['日期'])
df_volume['日期'] = pd.to_datetime(df_volume['日期'])

# 合并价格和成交量数据
df_merged = pd.merge(
    df_price,
    df_volume[['日期', 'SKU名称', '成交指数']],
    how='left',
    on=['日期', 'SKU名称']
)

# 获取所有有成交量数据的唯一产品名称
products_with_volume = df_merged[df_merged['成交指数'].notna()]['SKU名称'].unique()

# 为每个产品创建图表
for product in products_with_volume:
    # 筛选当前产品的数据
    product_df = df_merged[df_merged['SKU名称'] == product].copy()
    
    # 计算移动平均线
    product_df['3MA'] = product_df['收盘价'].rolling(window=3).mean()
    product_df['5MA'] = product_df['收盘价'].rolling(window=5).mean()
    product_df['20MA'] = product_df['收盘价'].rolling(window=20).mean()
    
    # 计算价格Y轴范围
    y_min = product_df['最低价'].min()
    y_max = product_df['最高价'].max()
    y_range = y_max - y_min
    y_min_display = y_min - (y_range * 0.2)
    y_max_display = y_max + (y_range * 0.2)
    
    # 创建子图
    fig = make_subplots(
        rows=2, 
        cols=1,
        shared_xaxes=True,
        vertical_spacing=0.05,
        row_heights=[0.7, 0.3]  # 设置主图和成交量图的高度比例
    )
    
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
            decreasing_line_color='green',
            increasing_line_width=3,
            decreasing_line_width=3
        ),
        row=1, col=1
    )
    
    # 添加移动平均线
    fig.add_trace(
        go.Scatter(
            x=product_df['日期'],
            y=product_df['3MA'],
            name='3MA',
            line=dict(color='#FF851B', width=2)
        ),
        row=1, col=1
    )
    
    fig.add_trace(
        go.Scatter(
            x=product_df['日期'],
            y=product_df['5MA'],
            name='5MA',
            line=dict(color='#0074D9', width=2)
        ),
        row=1, col=1
    )

    fig.add_trace(
        go.Scatter(
            x=product_df['日期'],
            y=product_df['20MA'],
            name='20MA',
            line=dict(color='#B10DC9', width=2)
        ),
        row=1, col=1
    )

    # 添加成交量柱状图
    colors = ['red' if row['收盘价'] >= row['开盘价'] else 'green' 
             for _, row in product_df.iterrows()]
    
    fig.add_trace(
        go.Bar(
            x=product_df['日期'],
            y=product_df['成交指数'],
            name='成交量',
            marker_color=colors,
            opacity=0.8
        ),
        row=2, col=1
    )

    # 获取最新价格和日期
    latest_price = product_df['收盘价'].iloc[-1]
    latest_date = product_df['日期'].iloc[-1]
    latest_volume = product_df['成交指数'].iloc[-1]
    
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
        font=dict(size=14, color='black'),
        bgcolor='white',
        bordercolor='black',
        borderwidth=1,
        align='left',
        row=1, col=1
    )
    
    # 更新布局
    fig.update_layout(
        title=dict(
            text=f'{product}指数',  # 简化标题格式
            x=0.5,
            y=0.95,
            font=dict(size=32)
        ),
        template='plotly_white',
        height=1000,  # 增加总图表高度
        plot_bgcolor='white',
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(size=20)
        ),
        yaxis=dict(
            title='价格指数',
            showgrid=True,
            gridcolor='lightgrey',
            fixedrange=True,
            range=[y_min_display, y_max_display],
            title_font=dict(size=24),
            tickfont=dict(size=20)
        ),
        yaxis2=dict(
            title='成交量',
            showgrid=True,
            gridcolor='lightgrey',
            title_font=dict(size=24),
            tickfont=dict(size=20)
        ),
        xaxis=dict(
            showgrid=True,
            gridcolor='lightgrey',
            rangeslider=dict(visible=False),
            title_font=dict(size=20),
            tickfont=dict(size=20)
        ),
        xaxis2=dict(
            title='日期',
            showgrid=True,
            gridcolor='lightgrey',
            title_font=dict(size=20),
            tickfont=dict(size=20)
        ),
        margin=dict(t=100, l=50, r=50, b=50)
    )
    
    # 生成文件名（使用当前时间戳避免文件重名）
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'analysis_results/{product}成交量价_{timestamp}.html'
    
    # 保存图表为HTML文件
    fig.write_html(filename)
    print(f'已生成{product}的价格和成交量走势图：{filename}') 