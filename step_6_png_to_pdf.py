from PIL import Image
import glob
import os

def ensure_dir(directory):
    """确保目录存在，如果不存在则创建"""
    if not os.path.exists(directory):
        os.makedirs(directory)
        print(f'创建目录：{directory}')

def convert_png_to_pdf(png_file, output_dir='output_reports', password=None):
    # 确保输出目录存在
    ensure_dir(output_dir)
    
    # 打开PNG图片
    image = Image.open(png_file)
    
    # 如果图片是RGBA模式，转换为RGB模式（PDF不支持RGBA）
    if image.mode == 'RGBA':
        image = image.convert('RGB')
    
    # 获取文件名（不含路径）并生成PDF文件路径
    base_name = os.path.basename(png_file)
    pdf_name = os.path.splitext(base_name)[0] + '.pdf'
    pdf_file = os.path.join(output_dir, pdf_name)
    
    # 保存为PDF，添加密码保护
    if password:
        image.save(pdf_file, 'PDF', resolution=100.0, encrypt=True, encrypt_password=password)
        print(f'已将 {png_file} 转换为加密PDF {pdf_file}')
    else:
        image.save(pdf_file, 'PDF', resolution=100.0)
        print(f'已将 {png_file} 转换为 {pdf_file}')

if __name__ == '__main__':
    # 转换最新的价格指数报告
    png_file = '关键部位均权指数MK20250715-A-2.png'
    output_dir = 'output_reports'  # 设置输出目录
    
    if os.path.exists(png_file):
        convert_png_to_pdf(png_file, output_dir)
    else:
        print(f'文件 {png_file} 不存在') 