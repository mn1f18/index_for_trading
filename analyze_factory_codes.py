import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
from datetime import datetime

def analyze_factory_frequency(excel_file, plate_name_filter="件套", output_dir="analysis_results"):
    # 创建输出目录
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 设置pandas显示选项，避免中文显示为省略号
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', None)

    # 定义新老厂号
    old_factories = {
        'SIF3225', 'SIF337', 'SIF385', 'SIF42', 
        'SIF4507', 'SIF504', 'SIF2058', 'SIF4333'
    }

    new_factories = {
        'SIF4302', 'SIF4400', 'SIF3470', 'SIF3000',
        'SIF1662', 'SIF2880', 'SIF1110', 'SIF457',
        'SIF3181', 'SIF51'
    }

    try:
        # 读取Excel文件
        df = pd.read_excel(excel_file)
    except Exception as e:
        print(f"读取Excel文件时出错: {str(e)}")
        return

    # 获取列名（假设列名为"厂号"和"盘名称"，如果不是请修改）
    factory_column = "厂号"
    plate_name_column = "盘名称"

    # 根据盘名称进行筛选
    filtered_df = df[df[plate_name_column].str.contains(plate_name_filter, na=False)]
    
    # 如果没有找到匹配的记录，返回提示
    if len(filtered_df) == 0:
        print(f"\n没有找到包含'{plate_name_filter}'的记录")
        print(f"可用的盘名称：")
        for name in sorted(df[plate_name_column].unique()):
            print(f"- {name}")
        return

    # 获取所有匹配的盘名称
    unique_plate_names = filtered_df[plate_name_column].unique()
    print(f"\n找到以下包含'{plate_name_filter}'的盘名称：")
    for name in unique_plate_names:
        print(f"- {name}")

    # 统计厂号出现频率
    factory_counts = filtered_df[factory_column].value_counts()
    total_records = len(filtered_df)

    # 分别统计新厂和老厂的频率
    old_factory_counts = {}
    new_factory_counts = {}

    # 统计在我们关注的厂号列表中的记录数
    valid_records = 0
    for factory, count in factory_counts.items():
        if factory in old_factories:
            old_factory_counts[factory] = count
            valid_records += count
        elif factory in new_factories:
            new_factory_counts[factory] = count
            valid_records += count

    # 创建一个函数来格式化输出统计信息
    def print_factory_stats(title, factory_dict, valid_records):
        print(f"\n=== {title} ===")
        total_category = sum(factory_dict.values()) if factory_dict else 0
        for factory, count in factory_dict.items():
            percentage = (count / valid_records) * 100
            print(f"{factory}: {count}次 ({percentage:.2f}%)")
        category_percentage = (total_category / valid_records) * 100
        print(f"小计: {total_category}次 ({category_percentage:.2f}%)")
        return total_category

    # 打印统计信息
    print(f"\n在总共{total_records}条记录中，符合新老厂号列表的记录共{valid_records}条")
    total_old = print_factory_stats("老厂出现频率", old_factory_counts, valid_records)
    total_new = print_factory_stats("新厂出现频率", new_factory_counts, valid_records)

    # 创建plotly图表
    fig = make_subplots(
        rows=3, cols=1,
        subplot_titles=('老厂出现频率', '新厂出现频率', '新老厂总体比例'),
        row_heights=[0.4, 0.4, 0.2]  # 调整每个子图的高度比例
    )

    # 添加老厂数据
    if old_factory_counts:
        fig.add_trace(
            go.Bar(
                x=list(old_factory_counts.keys()),
                y=list(old_factory_counts.values()),
                text=[f"{count}次<br>({(count/valid_records*100):.2f}%)" for count in old_factory_counts.values()],
                textposition='auto',
                name='老厂',
                marker_color='royalblue'
            ),
            row=1, col=1
        )

    # 添加新厂数据
    if new_factory_counts:
        fig.add_trace(
            go.Bar(
                x=list(new_factory_counts.keys()),
                y=list(new_factory_counts.values()),
                text=[f"{count}次<br>({(count/valid_records*100):.2f}%)" for count in new_factory_counts.values()],
                textposition='auto',
                name='新厂',
                marker_color='lightgreen'
            ),
            row=2, col=1
        )

    # 添加总体比例数据
    total_comparison = {
        '老厂': total_old,
        '新厂': total_new
    }
    
    fig.add_trace(
        go.Bar(
            x=list(total_comparison.keys()),
            y=list(total_comparison.values()),
            text=[f"{count}次<br>({(count/valid_records*100):.2f}%)" for count in total_comparison.values()],
            textposition='auto',
            marker_color=['royalblue', 'lightgreen'],
            showlegend=False
        ),
        row=3, col=1
    )

    # 更新布局
    title_text = f'新老厂号出现频率分析 - {plate_name_filter}'
    fig.update_layout(
        title_text=title_text,
        showlegend=False,
        font=dict(size=12),
        height=1000  # 增加总高度以适应新的子图
    )

    # 更新x轴和y轴标签
    for i in range(1, 4):
        fig.update_xaxes(title_text='厂号' if i < 3 else '', row=i, col=1)
        fig.update_yaxes(title_text='出现次数', row=i, col=1)

    # 调整第三个子图的样式
    fig.update_xaxes(tickangle=0, row=3, col=1)  # 确保标签水平显示
    
    # 生成时间戳
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    # 使用搜索关键词作为文件名
    base_filename = f"厂号分析_{plate_name_filter}_{timestamp}"
    
    # 保存为HTML文件（可交互）
    html_path = os.path.join(output_dir, f"{base_filename}.html")
    fig.write_html(html_path)

    # 打印总体统计信息
    print("\n=== 总体统计 ===")
    print(f"符合新老厂号列表的记录总数: {valid_records}")
    print(f"老厂总出现次数: {total_old} ({(total_old/valid_records*100):.2f}%)")
    print(f"新厂总出现次数: {total_new} ({(total_new/valid_records*100):.2f}%)")
    
    # 打印文件保存位置
    print(f"\n分析结果已保存至：")
    print(f"HTML文件（可交互）：{html_path}")
    print(f"\n提示：HTML文件可以直接在浏览器中打开，支持放大、缩小、悬停查看详情等交互功能")

if __name__ == "__main__":
    # 可以通过修改参数来筛选不同的盘名称
    analyze_factory_frequency("2025-06-25报盘.xlsx", plate_name_filter="件套", output_dir="analysis_results") 