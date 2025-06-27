from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import os
from datetime import datetime
import time

def convert_html_to_image():
    try:
        # 设置Chrome选项
        chrome_options = Options()
        chrome_options.add_argument('--headless')  # 无界面模式
        chrome_options.add_argument('--start-maximized')  # 最大化窗口
        chrome_options.add_argument('--disable-gpu')  # 禁用GPU加速
        chrome_options.add_argument('--no-sandbox')  # 禁用沙盒模式
        chrome_options.add_argument('--disable-dev-shm-usage')  # 禁用/dev/shm使用
        # 性能优化选项
        chrome_options.add_argument('--disable-extensions')  # 禁用扩展
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

        # 获取当前工作目录的绝对路径
        current_dir = os.getcwd()
        html_path = os.path.join(current_dir, 'onepage.html')
        html_url = f"file:///{html_path}"

        print("正在启动Chrome浏览器...")
        # 使用ChromeDriverManager自动管理驱动
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        # 设置正常的窗口大小
        driver.set_window_size(1366, 768)
        
        print("正在加载页面...")
        driver.get(html_url)

        # 等待页面加载完成
        time.sleep(5)

        # 获取页面实际高度并调整窗口大小
        page_height = driver.execute_script('return document.documentElement.scrollHeight')
        driver.set_window_size(1366, page_height)

        # 再次等待以确保调整后的页面稳定
        time.sleep(2)

        # 生成文件名
        current_date = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f'价格指数报告_{current_date}.png'

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

if __name__ == "__main__":
    convert_html_to_image() 