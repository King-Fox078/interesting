"""
移动端预约系统主入口
"""

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner
from kivy.uix.progressbar import ProgressBar
from kivy.uix.popup import Popup
from kivy.uix.tabbedpanel import TabbedPanel, TabbedPanelItem
from kivy.core.window import Window
from kivy.clock import Clock
from kivy.storage.jsonstore import JsonStore
from kivy.utils import platform
from kivy.core.text import LabelBase
from kivy.resources import resource_add_path
import threading
import time
import json
import requests
import random
from datetime import datetime
import os


# 设置中文字体
def setup_chinese_font():
    """设置中文字体"""
    try:
        # 尝试使用系统字体
        if platform == 'android':
            # Android系统字体
            resource_add_path('/system/fonts')
            LabelBase.register('Roboto', 'DroidSansFallback.ttf')
        elif platform == 'win':
            # Windows系统字体
            resource_add_path('C:/Windows/Fonts')
            LabelBase.register('Roboto', 'msyh.ttc')  # 微软雅黑
        elif platform == 'linux':
            # Linux系统字体
            resource_add_path('/usr/share/fonts')
            LabelBase.register('Roboto', 'DejaVuSans.ttf')
        elif platform == 'macosx':
            # macOS系统字体
            resource_add_path('/System/Library/Fonts')
            LabelBase.register('Roboto', 'Arial.ttf')
    except:
        # 如果系统字体不可用，使用默认字体
        pass


# 设置窗口大小
if platform == 'android':
    Window.softinput_mode = 'below_target'
else:
    Window.size = (400, 700)

# 设置中文字体
setup_chinese_font()


class BookingDataManager:
    """预约数据管理类"""

    def __init__(self):
        self.store = JsonStore('booking_config.json')
        self.load_default_config()

    def load_default_config(self):
        """加载默认配置"""
        if not self.store.exists('config'):
            default_config = {
                'ip_pool': [
                    '192.168.1.100:8080',
                    '192.168.1.101:8080',
                    '192.168.1.102:8080'
                ],
                'user_agents': [
                    'Mozilla/5.0 (Linux; Android 10; SM-G975F) AppleWebKit/537.36',
                    'Mozilla/5.0 (iPhone; CPU iPhone OS 14_0 like Mac OS X) AppleWebKit/605.1.15',
                    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                ],
                'booking_sites': {
                    'mao_memorial': {
                        'url': 'https://jnt.mfu.com.cn/page/user',
                        'max_threads': 5,
                        'requests_per_second': 2,
                        'timeout': 10
                    },
                    'tiananmen': {
                        'url': 'https://yuyue.tamgw.beijing.gov.cn/',
                        'max_threads': 5,
                        'requests_per_second': 2,
                        'timeout': 10
                    }
                },
                'user_info': {
                    'name': '',
                    'phone': '',
                    'id_card': '',
                    'visit_date': '',
                    'visit_time': '09:00',
                    'visitor_count': 1
                }
            }
            self.store.put('config', **default_config)

    def get_data(self):
        """获取配置数据"""
        return self.store.get('config')

    def save_data(self, config):
        """保存配置数据"""
        self.store.put('config', **config)

    def update_user_info(self, user_info):
        """更新用户信息"""
        config = self.get_data()
        config['user_info'] = user_info
        self.save_data(config)

    def update_ip_pool(self, ip_pool):
        """更新IP池"""
        config = self.get_data()
        config['ip_pool'] = ip_pool
        self.save_data(config)


class MobileBookingSystem:
    """移动端预约系统"""

    def __init__(self, callback=None):
        self.data_manager = BookingDataManager()
        self.is_booking = False
        self.booking_threads = []
        self.success_count = 0
        self.fail_count = 0
        self.callback = callback
        self.lock = threading.Lock()

    def start_booking(self, site_name, user_info, max_workers=5):
        """开始预约"""
        if self.is_booking:
            return False

        self.is_booking = True
        self.success_count = 0
        self.fail_count = 0

        # 在新线程中运行预约任务
        thread = threading.Thread(target=self._run_booking, args=(site_name, user_info, max_workers))
        thread.daemon = True
        thread.start()
        self.booking_threads.append(thread)

        return True

    def stop_booking(self):
        """停止预约"""
        self.is_booking = False
        for thread in self.booking_threads:
            if thread.is_alive():
                thread.join(timeout=1)
        self.booking_threads.clear()

    def _run_booking(self, site_name, user_info, max_workers):
        """运行预约任务"""
        config = self.data_manager.get_data()
        site_config = config['booking_sites'].get(site_name, {})

        for i in range(max_workers):
            if not self.is_booking:
                break

            thread = threading.Thread(target=self._booking_worker, args=(site_config, user_info, f"Worker-{i}"))
            thread.daemon = True
            thread.start()
            self.booking_threads.append(thread)
            time.sleep(0.2)

        # 等待所有工作线程完成
        for thread in self.booking_threads:
            if thread.is_alive():
                thread.join()

    def _booking_worker(self, site_config, user_info, worker_name):
        """预约工作线程"""
        url = site_config.get('url', '')
        timeout = site_config.get('timeout', 10)
        requests_per_second = site_config.get('requests_per_second', 2)

        while self.is_booking:
            try:
                # 频率控制
                time.sleep(1.0 / requests_per_second)

                # 发送请求
                response = self._make_request(url, user_info, timeout)

                if response and self._check_booking_success(response):
                    with self.lock:
                        self.success_count += 1
                    self._notify_success(f"{worker_name} 预约成功!")
                    break
                else:
                    with self.lock:
                        self.fail_count += 1
                    self._notify_progress(f"{worker_name} 预约失败，重试中...")

            except Exception as e:
                with self.lock:
                    self.fail_count += 1
                self._notify_progress(f"{worker_name} 异常: {str(e)}")
                time.sleep(1)

    def _make_request(self, url, user_info, timeout):
        """发送请求"""
        try:
            headers = {
                'User-Agent': random.choice(self.data_manager.get_data()['user_agents']),
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                'Accept-Language': 'zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3',
                'Accept-Encoding': 'gzip, deflate',
                'Connection': 'keep-alive',
                'Upgrade-Insecure-Requests': '1',
            }

            # 构建请求数据
            data = {
                'name': user_info.get('name', ''),
                'phone': user_info.get('phone', ''),
                'id_card': user_info.get('id_card', ''),
                'visit_date': user_info.get('visit_date', ''),
                'visit_time': user_info.get('visit_time', '09:00'),
                'visitor_count': user_info.get('visitor_count', 1)
            }

            response = requests.post(url, headers=headers, data=data, timeout=timeout)
            return response

        except Exception as e:
            return None

    def _check_booking_success(self, response):
        """检查预约是否成功"""
        if not response or response.status_code != 200:
            return False

        # 根据响应内容判断是否成功
        content = response.text.lower()
        success_indicators = ['success', '预约成功', '预约已提交', '提交成功']

        return any(indicator in content for indicator in success_indicators)

    def _notify_success(self, message):
        """通知成功"""
        if self.callback:
            Clock.schedule_once(lambda dt: self.callback('success', message), 0)

    def _notify_progress(self, message):
        """通知进度"""
        if self.callback:
            Clock.schedule_once(lambda dt: self.callback('progress', message), 0)

    def get_statistics(self):
        """获取统计信息"""
        return {
            'success_count': self.success_count,
            'fail_count': self.fail_count,
            'total_requests': self.success_count + self.fail_count
        }


class BookingApp(App):
    """预约应用主类"""

    def __init__(self):
        super().__init__()
        self.booking_system = MobileBookingSystem(callback=self.on_booking_event)
        self.data_manager = BookingDataManager()

    def build(self):
        """构建应用界面"""
        self.title = "预约系统"

        # 创建主布局
        main_layout = BoxLayout(orientation='vertical', padding=10, spacing=10)

        # 创建标签页
        tab_panel = TabbedPanel()

        # 预约页面
        booking_tab = TabbedPanelItem(text='预约')
        booking_tab.add_widget(self.create_booking_tab())
        tab_panel.add_widget(booking_tab)

        # 设置页面
        settings_tab = TabbedPanelItem(text='设置')
        settings_tab.add_widget(self.create_settings_tab())
        tab_panel.add_widget(settings_tab)

        # 日志页面
        log_tab = TabbedPanelItem(text='日志')
        log_tab.add_widget(self.create_log_tab())
        tab_panel.add_widget(log_tab)

        main_layout.add_widget(tab_panel)

        return main_layout

    def create_booking_tab(self):
        """创建预约页面"""
        layout = BoxLayout(orientation='vertical', spacing=10)

        # 网站选择
        site_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=40)
        site_layout.add_widget(Label(text='预约网站:', size_hint_x=0.3))
        self.site_spinner = Spinner(
            text='毛主席纪念堂',
            values=('毛主席纪念堂', '天安门'),
            size_hint_x=0.7
        )
        site_layout.add_widget(self.site_spinner)
        layout.add_widget(site_layout)

        # 用户信息输入
        info_layout = GridLayout(cols=2, spacing=5, size_hint_y=None)
        info_layout.bind(minimum_height=info_layout.setter('height'))

        # 姓名
        info_layout.add_widget(Label(text='姓名:', size_hint_y=None, height=40))
        self.name_input = TextInput(multiline=False, size_hint_y=None, height=40)
        info_layout.add_widget(self.name_input)

        # 手机号
        info_layout.add_widget(Label(text='手机号:', size_hint_y=None, height=40))
        self.phone_input = TextInput(multiline=False, size_hint_y=None, height=40)
        info_layout.add_widget(self.phone_input)

        # 身份证号
        info_layout.add_widget(Label(text='身份证号:', size_hint_y=None, height=40))
        self.id_card_input = TextInput(multiline=False, size_hint_y=None, height=40)
        info_layout.add_widget(self.id_card_input)

        # 预约日期
        info_layout.add_widget(Label(text='预约日期:', size_hint_y=None, height=40))
        self.date_input = TextInput(multiline=False, size_hint_y=None, height=40)
        info_layout.add_widget(self.date_input)

        # 预约时间
        info_layout.add_widget(Label(text='预约时间:', size_hint_y=None, height=40))
        self.time_input = TextInput(text='09:00', multiline=False, size_hint_y=None, height=40)
        info_layout.add_widget(self.time_input)

        # 预约人数
        info_layout.add_widget(Label(text='预约人数:', size_hint_y=None, height=40))
        self.count_input = TextInput(text='1', multiline=False, size_hint_y=None, height=40)
        info_layout.add_widget(self.count_input)

        # 线程数
        info_layout.add_widget(Label(text='线程数:', size_hint_y=None, height=40))
        self.thread_input = TextInput(text='5', multiline=False, size_hint_y=None, height=40)
        info_layout.add_widget(self.thread_input)

        layout.add_widget(info_layout)

        # 控制按钮
        button_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=50)

        self.start_button = Button(text='开始预约', on_press=self.start_booking)
        button_layout.add_widget(self.start_button)

        self.stop_button = Button(text='停止预约', on_press=self.stop_booking, disabled=True)
        button_layout.add_widget(self.stop_button)

        layout.add_widget(button_layout)

        # 进度条
        self.progress_bar = ProgressBar(max=100, size_hint_y=None, height=20)
        layout.add_widget(self.progress_bar)

        # 统计信息
        self.stats_label = Label(text='成功: 0 | 失败: 0 | 总计: 0', size_hint_y=None, height=30)
        layout.add_widget(self.stats_label)

        # 加载保存的用户信息
        self.load_user_info()

        return layout

    def create_settings_tab(self):
        """创建设置页面"""
        layout = BoxLayout(orientation='vertical', spacing=10)

        # IP池设置
        ip_layout = BoxLayout(orientation='vertical', spacing=5)
        ip_layout.add_widget(Label(text='IP池配置:', size_hint_y=None, height=30))

        self.ip_input = TextInput(
            multiline=True,
            size_hint_y=None,
            height=100,
            hint_text='每行一个IP地址，格式: 192.168.1.100:8080'
        )
        ip_layout.add_widget(self.ip_input)

        # 加载IP池
        config = self.data_manager.get_data()
        ip_pool = config.get('ip_pool', [])
        self.ip_input.text = '\n'.join(ip_pool)

        layout.add_widget(ip_layout)

        # 保存按钮
        save_button = Button(text='保存设置', on_press=self.save_settings, size_hint_y=None, height=40)
        layout.add_widget(save_button)

        return layout

    def create_log_tab(self):
        """创建日志页面"""
        layout = BoxLayout(orientation='vertical')

        # 日志文本框
        self.log_text = TextInput(
            multiline=True,
            readonly=True,
            size_hint=(1, 1)
        )
        layout.add_widget(self.log_text)

        # 清除日志按钮
        clear_button = Button(text='清除日志', on_press=self.clear_log, size_hint_y=None, height=40)
        layout.add_widget(clear_button)

        return layout

    def load_user_info(self):
        """加载用户信息"""
        config = self.data_manager.get_data()
        user_info = config.get('user_info', {})

        self.name_input.text = user_info.get('name', '')
        self.phone_input.text = user_info.get('phone', '')
        self.id_card_input.text = user_info.get('id_card', '')
        self.date_input.text = user_info.get('visit_date', '')
        self.time_input.text = user_info.get('visit_time', '09:00')
        self.count_input.text = str(user_info.get('visitor_count', 1))

    def save_user_info(self):
        """保存用户信息"""
        user_info = {
            'name': self.name_input.text,
            'phone': self.phone_input.text,
            'id_card': self.id_card_input.text,
            'visit_date': self.date_input.text,
            'visit_time': self.time_input.text,
            'visitor_count': int(self.count_input.text or '1')
        }
        self.data_manager.update_user_info(user_info)

    def save_settings(self, instance):
        """保存设置"""
        try:
            # 保存IP池
            ip_text = self.ip_input.text.strip()
            ip_list = [ip.strip() for ip in ip_text.split('\n') if ip.strip()]
            self.data_manager.update_ip_pool(ip_list)

            # 保存用户信息
            self.save_user_info()

            self.add_log("设置保存成功")
            self.show_popup("成功", "设置已保存")

        except Exception as e:
            self.add_log(f"保存设置失败: {e}")
            self.show_popup("错误", f"保存设置失败: {e}")

    def start_booking(self, instance):
        """开始预约"""
        try:
            # 验证输入
            if not all([self.name_input.text, self.phone_input.text,
                        self.id_card_input.text, self.date_input.text]):
                self.show_popup("错误", "请填写完整的用户信息")
                return

            # 获取用户信息
            user_info = {
                'name': self.name_input.text,
                'phone': self.phone_input.text,
                'id_card': self.id_card_input.text,
                'visit_date': self.date_input.text,
                'visit_time': self.time_input.text,
                'visitor_count': int(self.count_input.text or '1')
            }

            # 确定预约网站
            site_map = {
                '毛主席纪念堂': 'mao_memorial',
                '天安门': 'tiananmen'
            }
            site_name = site_map.get(self.site_spinner.text, 'mao_memorial')

            # 获取线程数
            max_workers = int(self.thread_input.text or '5')

            # 开始预约
            success = self.booking_system.start_booking(site_name, user_info, max_workers)

            if success:
                self.start_button.disabled = True
                self.stop_button.disabled = False
                self.add_log(f"开始预约: {self.site_spinner.text}")
            else:
                self.show_popup("错误", "预约已在运行中")

        except Exception as e:
            self.add_log(f"启动预约失败: {e}")
            self.show_popup("错误", f"启动预约失败: {e}")

    def stop_booking(self, instance):
        """停止预约"""
        self.booking_system.stop_booking()
        self.start_button.disabled = False
        self.stop_button.disabled = True
        self.add_log("预约已停止")

    def on_booking_event(self, event_type, message):
        """预约事件回调"""
        if event_type == 'success':
            self.add_log(f"✅ {message}")
            self.show_popup("成功", message)
            self.start_button.disabled = False
            self.stop_button.disabled = True
        elif event_type == 'progress':
            self.add_log(f"🔄 {message}")

        # 更新统计信息
        stats = self.booking_system.get_statistics()
        self.stats_label.text = f"成功: {stats['success_count']} | 失败: {stats['fail_count']} | 总计: {stats['total_requests']}"

    def add_log(self, message):
        """添加日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        self.log_text.text += log_message

    def clear_log(self, instance):
        """清除日志"""
        self.log_text.text = ""

    def show_popup(self, title, message):
        """显示弹窗"""
        popup = Popup(
            title=title,
            content=Label(text=message),
            size_hint=(None, None),
            size=(300, 200)
        )
        popup.open()


if __name__ == '__main__':
    BookingApp().run()