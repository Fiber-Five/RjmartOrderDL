import json
import logging
import os
import time
from datetime import datetime, timedelta
from typing import Dict, List, Optional
from DrissionPage import Chromium, ChromiumOptions

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("rjmart_export.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class RJMartExporter:
    def __init__(self, browser_path: str, download_path: str, owner: str):
        """初始化RJMart导出器
        
        Args:
            browser_path: Chrome浏览器路径
            download_path: 文件下载基础路径
            owner: 账户所有者名称
        """
        self.owner = owner
        self.download_base_path = download_path
        # 只在这里定义用户下载路径，不要重复创建
        self.user_download_path = os.path.join(download_path, owner)
        os.makedirs(self.user_download_path, exist_ok=True)
        self.browser = self._init_browser(browser_path, self.user_download_path)
        self.tab = None

    def _init_browser(self, browser_path: str, download_path: str) -> Chromium:
        """初始化无头浏览器实例"""
        # 不需要在这里创建目录，因为已经在__init__中创建过了
        options = ChromiumOptions().set_browser_path(browser_path)
        
        # 设置无头模式
        options.headless(True)
        
        # 设置下载路径
        options.set_download_path(download_path)
        
        # 无头模式下的其他必要设置
        options.set_argument('--disable-gpu')
        options.set_argument('--no-sandbox')
        options.set_argument('--disable-dev-shm-usage')
        options.set_argument('--disable-software-rasterizer')
        options.set_argument('--window-size=1920,1080')
        
        # 设置默认下载行为
        options.set_pref('download.default_directory', download_path)
        options.set_pref('download.prompt_for_download', False)
        options.set_pref('download.directory_upgrade', True)
        options.set_pref('safebrowsing.enabled', True)

        logger.info(f"初始化无头浏览器，下载路径: {download_path}")
        return Chromium(addr_or_opts=options)

    def login(self, username: str, password: str) -> bool:
        """登录RJMart系统

        Args:
            username: 登录用户名
            password: 登录密码

        Returns:
            bool: 登录是否成功
        """
        try:
            self.tab = self.browser.new_tab()
            self.tab.get('https://www.rjmart.cn/Login')
            self.tab.ele('@name=username').clear().input(username)
            self.tab.ele('@type=password').clear().input(password)
            self.tab.ele('@type=submit').click()
            self.tab.wait(3)
            logger.info(f"用户 {username} 登录成功")
            return True
        except Exception as e:
            logger.error(f"登录失败: {e}")
            return False

    def set_date_range(self, start_date: str) -> bool:
        """设置日期范围

        Args:
            start_date: 开始日期，格式：YYYY-MM-DD

        Returns:
            bool: 设置是否成功
        """
        try:
            # 访问订单列表页面
            self.tab.get('https://www.rjmart.cn/PM/orderList')
            self.tab.wait(3)  # 增加页面加载等待时间

            # 设置结束日期为昨天
            yesterday = datetime.now() - timedelta(days=1)
            end_date = yesterday.strftime('%Y-%m-%d')

            # 设置开始日期
            start_js = self._get_date_set_js('start', start_date)
            end_js = self._get_date_set_js('end', end_date)

            # 设置开始日期并等待
            if not self.tab.run_js(start_js):
                raise Exception("未找到开始日期输入框")
            self.tab.wait(2)  # 等待开始日期设置生效

            # 设置结束日期并等待
            if not self.tab.run_js(end_js):
                raise Exception("未找到结束日期输入框")
            self.tab.wait(2)  # 等待结束日期设置生效

            # 点击查询按钮
            search_button = self.tab.ele('@class=zen_btn zen_btn-primary')
            if search_button:
                search_button.click()
                self.tab.wait(3)  # 增加查询等待时间

                # 验证日期范围是否设置成功
                start_input = self.tab.ele('input.ZenDatePicker-input-start')
                end_input = self.tab.ele('input.ZenDatePicker-input-end')

                if not (start_input and end_input):
                    raise Exception("无法验证日期输入框")

                actual_start = start_input.attr('value')
                actual_end = end_input.attr('value')

                if actual_start != start_date or actual_end != end_date:
                    logger.warning(f"日期范围设置可能不正确: 预期 {start_date} 至 {end_date}, 实际 {actual_start} 至 {actual_end}")
                    # 重试一次
                    self.tab.wait(2)
                    self.tab.run_js(start_js)
                    self.tab.wait(2)
                    self.tab.run_js(end_js)
                    self.tab.wait(2)
                    search_button.click()
                    self.tab.wait(3)

            logger.info(f"日期范围设置成功: {start_date} 至 {end_date}")
            return True
        except Exception as e:
            logger.error(f"设置日期范围失败: {e}")
            return False

    def export_data(self) -> bool:
        """导出所有类型的数据"""
        try:
            # 导出订单明细和商品明细
            self._export_details()

            # 导出列表
            self._export_list()

            # 下载和清理文件
            self._handle_downloads()

            logger.info("所有数据导出完成")
            return True
        except Exception as e:
            logger.error(f"导出数据失败: {e}")
            return False

    def _export_details(self):
        """导出订单明细和商品明细"""
        for i in range(2):  # 导出两种明细
            export_type = "导出订单明细" if i == 0 else "导出商品明细"
            logger.info(f"开始导出 {export_type}")

            export_button = self.tab.ele('xpath://div[@class="operateArea zen_il-bl zen_v-m"]//div[@class="ZenDropMenu-trigger"]')
            if not export_button:
                raise Exception("未找到导出按钮")
            logger.info("找到导出按钮")

            export_button.hover()
            self.tab.wait(2)
            logger.info("悬停在导出按钮上")

            type_element = self.tab.ele(f'xpath://span[@class="ZenSelect-item-text " and text()="{export_type}"]')
            if not type_element:
                raise Exception(f"未找到{export_type}选项")
            logger.info(f"找到{export_type}选项")

            type_element.click()
            self.tab.wait(2)
            logger.info(f"点击了{export_type}选项")

            self._close_export_dialog()
            self.tab.wait(2)
            logger.info(f"完成{export_type}导出")

    def _export_list(self):
        """导出列表数据"""
        self.tab.wait(2)
        export_list_button = self.tab.ele('xpath://div[@class="operateArea zen_il-bl zen_v-m"]//button[@type="button"]/span[text()="导出列表"]')
        export_list_button.click()
        self.tab.wait(2)

    def _handle_downloads(self):
        """处理下载项和清理记录"""
        try:
            self.tab.wait(3)
            # 修改元素查找方式
            download_list = self.tab.eles('xpath://div[contains(@class, "ZenTable-table-body")]/div')
            total_items = len(download_list) - 1
            logger.info(f"共有{total_items}个下载项")

            # 下载文件
            for i in range(1, 3):
                # 修改元素查找方式
                name = download_list[i].ele('xpath://div[1]/span').text
                base_name = name.split('-')[0]
                new_name = f"{base_name}_{self.owner}"
                
                target_file = os.path.join(self.user_download_path, f"{new_name}.xlsx")
                if os.path.exists(target_file):
                    os.remove(target_file)
                    logger.info(f"已删除已存在的文件: {target_file}")
                
                # 修改元素查找方式
                download_button = download_list[i].ele('xpath://div[contains(@class, "ZenTable-table-td")]//span[text()="下载"]')
                download_button.click.to_download(self.user_download_path, new_name)
                
                # 等待文件下载完成
                timeout = 30
                start_time = time.time()
                while time.time() - start_time < timeout:
                    if os.path.exists(target_file):
                        logger.info(f"已下载: {new_name} 到 {self.user_download_path}")
                        break
                    time.sleep(1)
                else:
                    logger.warning(f"下载超时: {new_name}")

            # 清理记录
            for i in range(total_items, 0, -1):
                # 修改元素查找方式
                name = download_list[i].ele('xpath://div[1]/span').text
                delete_button = download_list[i].ele('xpath://div[contains(@class, "ZenTable-table-td")]//span[text()="删除"]')
                delete_button.click()
                self.tab.wait(1)
                logger.info(f"已删除记录: {name}")
            
        except Exception as e:
            logger.error(f"处理下载项时出错: {e}")
            raise

    def _close_export_dialog(self):
        """关闭导出对话框"""
        try:
            close_button = self.tab.ele('xpath://span[@class="closeBtn zen_cur-p"]')
            if close_button:
                close_button.click()
                self.tab.wait(1)
        except Exception as e:
            logger.warning(f"关闭导出对话框失败: {e}")

    @staticmethod
    def _get_date_set_js(input_type: str, date: str) -> str:
        """生成设置日期的JavaScript代码

        Args:
            input_type: 输入框类型 ('start' 或 'end')
            date: 要设置的日期

        Returns:
            str: JavaScript代码
        """
        return f"""
            var dateInput = document.querySelector('input.ZenDatePicker-input-{input_type}');
            if (dateInput) {{
                dateInput.value = '{date}';
                var event = new Event('change', {{ bubbles: true }});
                dateInput.dispatchEvent(event);
                var inputEvent = new Event('input', {{ bubbles: true }});
                dateInput.dispatchEvent(inputEvent);
                return true;
            }}
            return false;
        """

    def update_user(self, new_owner: str):
        """更新当前用户信息并设置新的下载路径"""
        self.owner = new_owner
        self.user_download_path = os.path.join(self.download_base_path, new_owner)
        os.makedirs(self.user_download_path, exist_ok=True)
        
        # 更新浏览器下载路径
        self.browser.set.download_path(self.user_download_path)
        logger.info(f"已更新用户信息: {new_owner}, 下载路径: {self.user_download_path}")

    def close(self):
        """关闭浏览器"""
        try:
            if self.browser:
                self.browser.quit()
                logger.info("浏览器已关闭")
        except Exception as e:
            logger.error(f"关闭浏览器时出错: {e}")

def process_account(account: Dict, settings: Dict, exporter: RJMartExporter) -> bool:
    """处理单个账户的导出操作

    Args:
        account: 账户信息字典
        settings: 全局设置字典
        exporter: RJMart导出器实例

    Returns:
        bool: 处理是否成功
    """
    logger.info(f"开始处理用户 {account['owner']} 的数据导出")

    try:
        # 关闭之前的标签页
        if exporter.tab:
            try:
                exporter.tab.close()
                logger.info("已关闭上一个标签页")
            except Exception as e:
                logger.warning(f"关闭标签页时出错: {e}")
            exporter.tab = None

        # 执行导出流程
        if exporter.login(account['username'], account['password']):
            if exporter.set_date_range(settings['start_date']):
                if exporter.export_data():
                    logger.info(f"用户 {account['owner']} 数据导出成功")
                    return True
        return False
    except Exception as e:
        logger.error(f"用户 {account['owner']} 处理失败: {e}")
        return False

def main():
    try:
        # 读取配置文件
        with open('config.json', 'r', encoding='utf-8') as f:
            config = json.load(f)
            accounts = config['accounts']
            settings = config['settings']

        if not accounts:
            logger.error("没有找到账户配置")
            return

        # 使用第一个账户的浏览器配置初始化导出器
        download_path = settings['download_path']
        os.makedirs(download_path, exist_ok=True)
        
        # 创建一个共享的导出器实例
        exporter = RJMartExporter(
            browser_path=settings['browser_path'],
            download_path=download_path,
            owner=accounts[0]['owner']
        )

        # 处理每个账户
        success_count = 0
        try:
            for i, account in enumerate(accounts):
                logger.info(f"开始处理第 {i+1}/{len(accounts)} 个账户: {account['owner']}")
                
                # 更新导出器的用户信息
                exporter.update_user(account['owner'])
                
                if process_account(account, settings, exporter):
                    success_count += 1

            # 输出处理结果统计
            logger.info(f"处理完成: 成功 {success_count} 个, 失败 {len(accounts) - success_count} 个")

        finally:
            # 所有账户处理完后关闭浏览器
            exporter.close()

    except Exception as e:
        logger.error(f"程序执行出错: {e}")
    finally:
        logger.info("程序结束")


if __name__ == "__main__":
    main()
