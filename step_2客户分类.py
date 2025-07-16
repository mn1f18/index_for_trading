import pandas as pd
import os
from datetime import datetime

# ========== 客户产品关心列表配置 ==========
# 在这里定义每个客户关心的产品列表
CUSTOMER_PRODUCTS = {
    "客户A": ["牛前八件套", "大米龙", "牛霖", "牛腩" , "牛前后", "去盖臀肉"],
    "客户B": ["牛前八件套", "龟排腱"],
    # "客户C": ["龟排腱", "牛霖"],  # 可以添加更多客户
}

# ========== 文件配置 ==========
SOURCE_FILE = "价格数据初步调整.xlsx"  # 源数据文件
OUTPUT_PREFIX = "价格数据初步调整_"      # 输出文件前缀

def load_source_data():
    """读取源数据文件"""
    try:
        df = pd.read_excel(SOURCE_FILE)
        print(f"✅ 成功读取源数据文件：{SOURCE_FILE}")
        print(f"   数据行数：{len(df)}")
        print(f"   包含产品：{list(df['SKU名称'].unique())}")
        return df
    except FileNotFoundError:
        print(f"❌ 错误：找不到源数据文件 {SOURCE_FILE}")
        return None
    except Exception as e:
        print(f"❌ 读取源数据文件时出错：{str(e)}")
        return None

def filter_customer_data(df, customer_name, product_list):
    """为指定客户筛选数据"""
    # 筛选客户关心的产品
    customer_df = df[df['SKU名称'].isin(product_list)].copy()
    
    if len(customer_df) == 0:
        print(f"⚠️  警告：{customer_name} 的产品列表在源数据中未找到匹配数据")
        return None
    
    # 按日期和产品名称排序
    customer_df = customer_df.sort_values(['日期', 'SKU名称']).reset_index(drop=True)
    
    return customer_df

def save_customer_file(df, customer_name):
    """保存客户数据到Excel文件"""
    filename = f"{OUTPUT_PREFIX}{customer_name}.xlsx"
    
    try:
        df.to_excel(filename, index=False)
        print(f"✅ 已生成 {customer_name} 数据文件：{filename}")
        print(f"   包含数据行数：{len(df)}")
        print(f"   包含产品：{list(df['SKU名称'].unique())}")
        print(f"   日期范围：{df['日期'].min().strftime('%Y-%m-%d')} 至 {df['日期'].max().strftime('%Y-%m-%d')}")
        return True
    except Exception as e:
        print(f"❌ 保存 {customer_name} 数据文件时出错：{str(e)}")
        return False

def generate_summary_report(source_df):
    """生成处理摘要报告"""
    print("\n" + "="*60)
    print("📊 客户数据分类摘要报告")
    print("="*60)
    
    print(f"源数据文件：{SOURCE_FILE}")
    print(f"总数据行数：{len(source_df)}")
    print(f"总产品数量：{len(source_df['SKU名称'].unique())}")
    print(f"日期范围：{source_df['日期'].min().strftime('%Y-%m-%d')} 至 {source_df['日期'].max().strftime('%Y-%m-%d')}")
    
    print(f"\n客户配置：")
    for customer, products in CUSTOMER_PRODUCTS.items():
        print(f"  {customer}：{len(products)} 个产品 -> {', '.join(products)}")
    
    print(f"\n生成的文件：")
    for customer in CUSTOMER_PRODUCTS.keys():
        filename = f"{OUTPUT_PREFIX}{customer}.xlsx"
        if os.path.exists(filename):
            print(f"  ✅ {filename}")
        else:
            print(f"  ❌ {filename} (生成失败)")

def main():
    """主函数"""
    print("🚀 开始客户数据分类处理...")
    print(f"处理时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 读取源数据
    source_df = load_source_data()
    if source_df is None:
        return
    
    # 检查源数据中的SKU名称列
    if 'SKU名称' not in source_df.columns:
        print("❌ 错误：源数据文件中未找到 'SKU名称' 列")
        print(f"   现有列名：{list(source_df.columns)}")
        return
    
    print(f"\n🔍 开始为 {len(CUSTOMER_PRODUCTS)} 个客户筛选数据...")
    
    success_count = 0
    
    # 为每个客户筛选和保存数据
    for customer_name, product_list in CUSTOMER_PRODUCTS.items():
        print(f"\n--- 处理 {customer_name} ---")
        print(f"关心的产品：{', '.join(product_list)}")
        
        # 筛选数据
        customer_df = filter_customer_data(source_df, customer_name, product_list)
        
        if customer_df is not None:
            # 检查产品匹配情况
            found_products = list(customer_df['SKU名称'].unique())
            missing_products = [p for p in product_list if p not in found_products]
            
            if missing_products:
                print(f"⚠️  未找到的产品：{', '.join(missing_products)}")
            
            # 保存文件
            if save_customer_file(customer_df, customer_name):
                success_count += 1
        else:
            print(f"❌ {customer_name} 数据筛选失败")
    
    # 生成摘要报告
    generate_summary_report(source_df)
    
    print(f"\n🎉 处理完成！成功生成 {success_count}/{len(CUSTOMER_PRODUCTS)} 个客户数据文件")

if __name__ == "__main__":
    main() 