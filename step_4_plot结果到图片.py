from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import os
from datetime import datetime
import time
import glob

def convert_plot_to_image():
    try:
        # 设置Chrome选项
        chrome_options = Options()
        chrome_options.add_argument('--headless')  # 无界面模式
        chrome_options.add_argument('--disable-gpu')  # 禁用GPU加速
        chrome_options.add_argument('--no-sandbox')  # 禁用沙盒模式
        chrome_options.add_argument('--disable-dev-shm-usage')  # 禁用/dev/shm使用
        # 性能优化选项
        chrome_options.add_argument('--disable-extensions')
        chrome_options.add_argument('--disable-browser-side-navigation')
        chrome_options.add_argument('--disable-infobars')
        chrome_options.add_argument('--disable-notifications')
        chrome_options.add_argument('--disable-popup-blocking')
        chrome_options.add_argument('--disable-save-password-bubble')
        chrome_options.add_argument('--disable-translate')
        chrome_options.add_argument('--disable-web-security')
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--disable-logging')
        chrome_options.add_argument('--log-level=3')
        chrome_options.add_argument('--silent')

        # 获取当前工作目录
        current_dir = os.getcwd()
        
        # 获取analysis_results目录下最新的HTML文件
        html_files = glob.glob(os.path.join(current_dir, 'analysis_results', '*.html'))
        if not html_files:
            print("未找到HTML文件！")
            return
            
        latest_html = max(html_files, key=os.path.getctime)
        html_url = f"file:///{latest_html}"
        print(f"找到最新的HTML文件: {os.path.basename(latest_html)}")
        print("正在启动Chrome浏览器...")
        
        # 使用ChromeDriverManager自动管理驱动
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        # 设置固定的窗口大小
        driver.set_window_size(1920, 900)
        
        print("正在加载页面...")
        driver.get(html_url)

        # 等待页面加载完成
        time.sleep(5)

        # 注入CSS来隐藏滚动条并设置页面样式
        driver.execute_script("""
            document.body.style.overflow = 'hidden';
            document.documentElement.style.overflow = 'hidden';
            document.body.style.margin = '0';
            document.body.style.padding = '0';
            
            // 确保所有plotly图表容器可见
            var plots = document.querySelectorAll('.js-plotly-plot');
            plots.forEach(function(plot) {
                plot.style.visibility = 'visible';
                plot.style.display = 'block';
                plot.style.height = '950px';  // 改回原来的高度
            });
        """)

        # 等待一下以确保样式应用完成
        time.sleep(2)

        # 生成输出文件名（保持与HTML文件名一致，仅改扩展名）
        output_filename = os.path.splitext(latest_html)[0] + '.png'

        print("正在截取页面...")
        # 截取整个页面
        driver.save_screenshot(output_filename)
        
        print(f"图片生成成功！文件保存为: {output_filename}")
        
        # 获取生成的图片文件大小
        file_size = os.path.getsize(output_filename) / 1024  # 转换为KB
        print(f"图片文件大小: {file_size:.2f} KB")

    except Exception as e:
        print(f"转换过程中出现错误: {str(e)}")
    
    finally:
        if 'driver' in locals():
            driver.quit()

def batch_convert_plots():
    try:
        # 获取当前工作目录
        current_dir = os.getcwd()
        
        # 获取analysis_results目录下所有的HTML文件
        html_files = glob.glob(os.path.join(current_dir, 'analysis_results', '*.html'))
        
        if not html_files:
            print("未找到HTML文件！")
            return
            
        print(f"找到 {len(html_files)} 个HTML文件")
        
        # 设置Chrome选项（与单个转换相同的设置）
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-extensions')
        chrome_options.add_argument('--disable-browser-side-navigation')
        chrome_options.add_argument('--disable-infobars')
        chrome_options.add_argument('--disable-notifications')
        chrome_options.add_argument('--disable-popup-blocking')
        chrome_options.add_argument('--disable-save-password-bubble')
        chrome_options.add_argument('--disable-translate')
        chrome_options.add_argument('--disable-web-security')
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--disable-logging')
        chrome_options.add_argument('--log-level=3')
        chrome_options.add_argument('--silent')

        print("正在启动Chrome浏览器...")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.set_window_size(1920, 900)

        for html_file in html_files:
            try:
                html_url = f"file:///{html_file}"
                print(f"\n正在处理: {os.path.basename(html_file)}")
                
                driver.get(html_url)
                time.sleep(5)
                
                # 注入CSS和JavaScript
                driver.execute_script("""
                    document.body.style.overflow = 'hidden';
                    document.documentElement.style.overflow = 'hidden';
                    document.body.style.margin = '0';
                    document.body.style.padding = '0';
                    
                    // 确保所有plotly图表容器可见
                    var plots = document.querySelectorAll('.js-plotly-plot');
                    plots.forEach(function(plot) {
                        plot.style.visibility = 'visible';
                        plot.style.display = 'block';
                        plot.style.height = '950px';  // 改回原来的高度
                    });
                """)

                time.sleep(2)
                
                # 生成输出文件名
                output_filename = os.path.splitext(html_file)[0] + '.png'
                
                # 截取页面
                driver.save_screenshot(output_filename)
                
                file_size = os.path.getsize(output_filename) / 1024
                print(f"已生成图片: {os.path.basename(output_filename)}")
                print(f"图片大小: {file_size:.2f} KB")
                
            except Exception as e:
                print(f"处理 {os.path.basename(html_file)} 时出错: {str(e)}")
                continue

    except Exception as e:
        print(f"批量转换过程中出现错误: {str(e)}")
    
    finally:
        if 'driver' in locals():
            driver.quit()

def convert_specific_html(html_path, is_volume_chart=False):
    try:
        if not os.path.exists(html_path):
            print(f"错误：文件 {html_path} 不存在！")
            return
            
        # 设置Chrome选项
        chrome_options = Options()
        chrome_options.add_argument('--headless')  # 无界面模式
        chrome_options.add_argument('--disable-gpu')  # 禁用GPU加速
        chrome_options.add_argument('--no-sandbox')  # 禁用沙盒模式
        chrome_options.add_argument('--disable-dev-shm-usage')  # 禁用/dev/shm使用
        chrome_options.add_argument('--disable-extensions')
        chrome_options.add_argument('--disable-browser-side-navigation')
        chrome_options.add_argument('--disable-infobars')
        chrome_options.add_argument('--disable-notifications')
        chrome_options.add_argument('--disable-popup-blocking')
        chrome_options.add_argument('--disable-save-password-bubble')
        chrome_options.add_argument('--disable-translate')
        chrome_options.add_argument('--disable-web-security')
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--disable-logging')
        chrome_options.add_argument('--log-level=3')
        chrome_options.add_argument('--silent')

        print(f"正在处理文件: {os.path.basename(html_path)}")
        print("正在启动Chrome浏览器...")
        
        # 使用ChromeDriverManager自动管理驱动
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        # 根据图表类型设置窗口大小
        if is_volume_chart:
            driver.set_window_size(1920, 1100)  # 成交量价图表使用更高的窗口
        else:
            driver.set_window_size(1920, 900)   # 其他图表保持原有高度
        
        # 构建文件URL
        html_url = f"file:///{html_path}"
        
        print("正在加载页面...")
        driver.get(html_url)

        # 等待页面加载完成
        time.sleep(5)

        # 注入CSS来隐藏滚动条并设置页面样式
        driver.execute_script("""
            document.body.style.overflow = 'hidden';
            document.documentElement.style.overflow = 'hidden';
            document.body.style.margin = '0';
            document.body.style.padding = '0';
            
            // 确保所有plotly图表容器可见
            var plots = document.querySelectorAll('.js-plotly-plot');
            plots.forEach(function(plot) {
                plot.style.visibility = 'visible';
                plot.style.display = 'block';
                plot.style.height = '950px';  // 改回原来的高度
            });
        """)

        # 等待一下以确保样式应用完成
        time.sleep(2)

        # 生成输出文件名（保持与HTML文件名一致，仅改扩展名）
        output_filename = os.path.splitext(html_path)[0] + '.png'

        print("正在截取页面...")
        # 截取整个页面
        driver.save_screenshot(output_filename)
        
        print(f"图片生成成功！文件保存为: {output_filename}")
        
        # 获取生成的图片文件大小
        file_size = os.path.getsize(output_filename) / 1024  # 转换为KB
        print(f"图片文件大小: {file_size:.2f} KB")

    except Exception as e:
        print(f"转换过程中出现错误: {str(e)}")
    
    finally:
        if 'driver' in locals():
            driver.quit()

def convert_volume_chart(html_path):
    try:
        if not os.path.exists(html_path):
            print(f"错误：文件 {html_path} 不存在！")
            return
            
        # 设置Chrome选项
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-extensions')
        chrome_options.add_argument('--disable-browser-side-navigation')
        chrome_options.add_argument('--disable-infobars')
        chrome_options.add_argument('--disable-notifications')
        chrome_options.add_argument('--disable-popup-blocking')
        chrome_options.add_argument('--disable-save-password-bubble')
        chrome_options.add_argument('--disable-translate')
        chrome_options.add_argument('--disable-web-security')
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--disable-logging')
        chrome_options.add_argument('--log-level=3')
        chrome_options.add_argument('--silent')

        print(f"正在处理文件: {os.path.basename(html_path)}")
        print("正在启动Chrome浏览器...")
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        # 专门为成交量价图表设置更高的窗口
        driver.set_window_size(1920, 1100)
        
        html_url = f"file:///{html_path}"
        
        print("正在加载页面...")
        driver.get(html_url)

        time.sleep(5)

        driver.execute_script("""
            document.body.style.overflow = 'hidden';
            document.documentElement.style.overflow = 'hidden';
            document.body.style.margin = '0';
            document.body.style.padding = '0';
            
            var plots = document.querySelectorAll('.js-plotly-plot');
            plots.forEach(function(plot) {
                plot.style.visibility = 'visible';
                plot.style.display = 'block';
                plot.style.height = '950px';
            });
        """)

        time.sleep(2)

        output_filename = os.path.splitext(html_path)[0] + '.png'

        print("正在截取页面...")
        driver.save_screenshot(output_filename)
        
        print(f"图片生成成功！文件保存为: {output_filename}")
        
        file_size = os.path.getsize(output_filename) / 1024
        print(f"图片文件大小: {file_size:.2f} KB")

    except Exception as e:
        print(f"转换过程中出现错误: {str(e)}")
    
    finally:
        if 'driver' in locals():
            driver.quit()

def convert_custom_name_plots(custom_name):
    try:
        # 获取当前工作目录
        current_dir = os.getcwd()
        
        # 搜索包含自定义名称的HTML文件
        pattern = os.path.join(current_dir, 'analysis_results', f'*{custom_name}*.html')
        html_files = glob.glob(pattern)
        
        if not html_files:
            print(f"未找到包含 '{custom_name}' 的HTML文件！")
            return
            
        print(f"找到 {len(html_files)} 个包含 '{custom_name}' 的HTML文件")
        
        # 设置Chrome选项
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-extensions')
        chrome_options.add_argument('--disable-browser-side-navigation')
        chrome_options.add_argument('--disable-infobars')
        chrome_options.add_argument('--disable-notifications')
        chrome_options.add_argument('--disable-popup-blocking')
        chrome_options.add_argument('--disable-save-password-bubble')
        chrome_options.add_argument('--disable-translate')
        chrome_options.add_argument('--disable-web-security')
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--disable-logging')
        chrome_options.add_argument('--log-level=3')
        chrome_options.add_argument('--silent')

        print("正在启动Chrome浏览器...")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.set_window_size(1920, 900)

        for html_file in html_files:
            try:
                
                html_url = f"file:///{html_file}"
                print(f"\n正在处理: {os.path.basename(html_file)}")
                
                driver.get(html_url)
                time.sleep(5)
                
                # 注入CSS和JavaScript
                driver.execute_script("""
                    document.body.style.overflow = 'hidden';
                    document.documentElement.style.overflow = 'hidden';
                    document.body.style.margin = '0';
                    document.body.style.padding = '0';
                    
                    // 确保所有plotly图表容器可见
                    var plots = document.querySelectorAll('.js-plotly-plot');
                    plots.forEach(function(plot) {
                        plot.style.visibility = 'visible';
                        plot.style.display = 'block';
                        plot.style.height = '950px';
                    });
                """)

                time.sleep(2)
                
                # 生成输出文件名
                output_filename = os.path.splitext(html_file)[0] + '.png'
                
                # 截取页面
                driver.save_screenshot(output_filename)
                
                file_size = os.path.getsize(output_filename) / 1024
                print(f"已生成图片: {os.path.basename(output_filename)}")
                print(f"图片大小: {file_size:.2f} KB")
                
            except Exception as e:
                print(f"处理 {os.path.basename(html_file)} 时出错: {str(e)}")
                continue

        print(f"\n完成！已成功转换 {len(html_files)} 个包含 '{custom_name}' 的HTML文件")

    except Exception as e:
        print(f"批量转换过程中出现错误: {str(e)}")
    
    finally:
        if 'driver' in locals():
            driver.quit()

if __name__ == "__main__":
    print("请选择转换模式：")
    print("1. 仅转换最新的plotly图表")
    print("2. 批量转换所有plotly图表")
    print("3. 转换最新的成交量价图表")
    print("4. 转换包含指定名称的所有图表")
    
    choice = input("请输入选项（1、2、3或4）: ")
    
    if choice == "1":
        convert_plot_to_image()
    elif choice == "2":
        batch_convert_plots()
    elif choice == "3":
        # 获取当前工作目录
        current_dir = os.getcwd()
        # 搜索所有包含"成交量价"的HTML文件
        volume_html_files = glob.glob(os.path.join(current_dir, 'analysis_results', '*成交量价_*.html'))
        
        if not volume_html_files:
            print("未找到成交量价相关的HTML文件！")
        else:
            # 获取最新的文件
            latest_volume_html = max(volume_html_files, key=os.path.getctime)
            print(f"找到最新的成交量价图表文件: {os.path.basename(latest_volume_html)}")
            convert_volume_chart(latest_volume_html)  # 使用专门的函数处理成交量价图表
    elif choice == "4":
        custom_name = input("请输入要搜索的名称（如：客户A）: ").strip()
        if custom_name:
            convert_custom_name_plots(custom_name)
        else:
            print("名称不能为空！")
    else:
        print("无效的选项！") 