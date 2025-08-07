import time
import random
import re
import threading
import win32gui
import win32con
import win32api
import requests
import json
import tkinter as tk
from tkinter import scrolledtext, ttk
from pywinauto import Application
from pywinauto.keyboard import send_keys
from pywinauto.uia_element_info import UIAElementInfo

class WeChatAutoReply:
    def __init__(self):
        self.app = None
        self.wechat_window = None
        self.wechat_hwnd = None
        self.wechat_element = None
        self.wechat_class_names = ["WeChatMainWndForPC", "WeChatMainWindow", "WeChatWindow"]
        self.wechat_path = r'C:\Program Files (x86)\Tencent\WeChat\WeChat.exe'
        self.debug = True
        self.visible_attr = self.detect_visible_attribute()
        self.input_offset_y = -80
        self.chat_list_container = None
        self.current_chat_identifier = None
        self.chat_switch_offset_x = 50
        self.other_message_ids = {}
        self.active_chat_element = None
        self.max_chat_count = 10
        self.other_msg_keywords = ["接收消息", "对方消息"]
        self.self_msg_keywords = ["发送消息", "自己消息"]

        self.time_pattern = re.compile(r'^[上下]午\d+:\d{2}$')

        self.public_account_keywords = ["公众号", "订阅号", "服务号", "official account", "weixin", "腾讯新闻"]

        self.root = tk.Tk()
        self.root.title("微信自动回复助手")
        self.root.geometry("800x600")
        self.is_running = False
        self.init_gui()


    def init_gui(self):

        control_frame = ttk.Frame(self.root, padding="10")
        control_frame.pack(fill=tk.X)
        self.start_btn = ttk.Button(control_frame, text="启动监控", command=self.toggle_monitor)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.status_var = tk.StringVar(value="未运行")
        status_label = ttk.Label(control_frame, textvariable=self.status_var)
        status_label.pack(side=tk.LEFT, padx=20)

        log_frame = ttk.LabelFrame(self.root, text="操作进程", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        style = ttk.Style()
        style.configure("TButton", font=("微软雅黑", 10))
        style.configure("TLabel", font=("微软雅黑", 10))
        style.configure("TLabelframe", font=("微软雅黑", 10))
        style.configure("TLabelframe.Label", font=("微软雅黑", 10))


    def append_log(self, message):
        self.log_text.config(state=tk.NORMAL)
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)


    def log(self, message):
        print(f"[DEBUG] {message}")
        self.append_log(message)

    def detect_visible_attribute(self):
        try:
            test_elem = UIAElementInfo()
            if hasattr(test_elem, 'is_visible'):
                return 'is_visible'
            elif hasattr(test_elem, 'visible'):
                return 'visible'
            else:
                return None
        except:
            return None

    def is_element_visible(self, elem):
        if not self.visible_attr:
            return True
        try:
            return getattr(elem, self.visible_attr, True)
        except:
            return True

    def get_wechat_hwnd(self):
        hwnds = []
        for class_name in self.wechat_class_names:
            hwnd = win32gui.FindWindow(class_name, None)
            if hwnd != 0:
                hwnds.append(hwnd)
                self.log(f"通过类名 {class_name} 找到窗口句柄: {hwnd}")
        if not hwnds:
            def callback(hwnd, extra):
                title = win32gui.GetWindowText(hwnd)
                if "微信" in title and win32gui.IsWindowVisible(hwnd):
                    extra.append(hwnd)
                    self.log(f"通过标题找到窗口句柄: {hwnd}, 标题: {title}")
            win32gui.EnumWindows(callback, hwnds)
        return hwnds[0] if hwnds else None

    def connect_wechat(self):
        try:
            self.wechat_hwnd = self.get_wechat_hwnd()
            if not self.wechat_hwnd:
                raise Exception("未找到微信窗口句柄")
            self.wechat_element = UIAElementInfo(self.wechat_hwnd)
            self.app = Application(backend="uia").connect(handle=self.wechat_hwnd)
            self.wechat_window = self.app.window(handle=self.wechat_hwnd)
            self.log("成功连接到微信客户端")
            placement = win32gui.GetWindowPlacement(self.wechat_hwnd)
            if placement[1] == win32con.SW_SHOWMINIMIZED:
                win32gui.ShowWindow(self.wechat_hwnd, win32con.SW_RESTORE)

                for _ in range(3):
                    if not self.is_running:
                        return False
                    time.sleep(1)
            win32gui.MoveWindow(self.wechat_hwnd, 600, 100, 1200, 900, True)
            win32gui.SetForegroundWindow(self.wechat_hwnd)

            for _ in range(3):
                if not self.is_running:
                    return False
                time.sleep(1)
            return True
        except Exception as e:
            self.log(f"连接微信失败: {str(e)}")
            try:
                self.app = Application(backend="uia").start(self.wechat_path)

                for _ in range(20):
                    if not self.is_running:
                        return False
                    time.sleep(1)
                return self.connect_wechat()
            except Exception as e:
                self.log(f"启动微信失败: {str(e)}")
                return False

    def get_new_messages(self, chat_identifier):
        """获取新消息，新增时间格式过滤，并检查后续五条消息是否为自己的消息"""
        try:
            if not chat_identifier:
                self.log("未指定聊天标识，跳过消息检测")
                return []
            if chat_identifier not in self.other_message_ids:
                self.other_message_ids[chat_identifier] = set()
                self.log(f"初始化 [{chat_identifier}] 的消息ID集合")
            self.log(f"开始检查 [{chat_identifier}] 的新消息...")

            message_container = None
            for elem in self.wechat_element.descendants():

                if not self.is_running:
                    return []
                try:
                    if not self.is_element_visible(elem):
                        continue
                    if (elem.control_type in ["List", "Pane"] and
                            elem.rectangle.width() > 500 and
                            elem.rectangle.height() > 400 and
                            ("消息" in elem.name or "chat" in elem.name.lower())):
                        message_container = elem
                        self.log(f"找到消息容器（{elem.control_type}）")
                        break
                except:
                    continue
            if not message_container:
                self.log("未找到精准消息容器，使用宽松匹配")
                for elem in self.wechat_element.descendants():

                    if not self.is_running:
                        return []
                    try:
                        if (self.is_element_visible(elem) and
                                elem.control_type in ["List", "Pane"] and
                                elem.rectangle.width() > 400 and
                                elem.rectangle.height() > 300):
                            message_container = elem
                            self.log(f"宽松匹配找到消息容器（{elem.control_type}）")
                            break
                    except:
                        continue

            if not message_container:
                self.log("未找到任何消息容器，提取所有文本元素")
                message_elements = [e for e in self.wechat_element.descendants()
                                    if e.control_type == "Text" and self.is_element_visible(e)]
            else:
                message_elements = [e for e in message_container.descendants()
                                    if e.control_type == "Text" and self.is_element_visible(e)]
            self.log(f"[{chat_identifier}] 找到文本元素总数: {len(message_elements)} 个")
            new_other_messages = []
            window_rect = self.wechat_element.rectangle
            window_center_x = window_rect.left + (window_rect.width() // 2)

            message_elements = sorted(message_elements, key=lambda e: e.rectangle.top)
            for idx, elem in enumerate(message_elements):

                if not self.is_running:
                    return []
                try:
                    if not self.is_element_in_active_chat(elem):
                        continue
                    elem_text = elem.name.strip()
                    elem_name = elem.name.lower() if elem.name else ""
                    elem_rect = elem.rectangle

                    if self.time_pattern.match(elem_text):
                        self.log(f"[{chat_identifier}] 过滤时间格式: {elem_text}（序号{idx}）")
                        continue

                    if not elem_text:
                        continue

                    elem_id = f"{elem.runtime_id}_{elem_rect.left}_{elem_rect.top}_{hash(elem_text)}"

                    pos_sender = "other" if elem_rect.left < window_center_x else "self"
                    keyword_sender = "other" if any(k in elem_name for k in self.other_msg_keywords) else "self" if any(
                        k in elem_name for k in self.self_msg_keywords) else None
                    sender = keyword_sender if keyword_sender is not None else pos_sender

                    if sender == "other":
                        if elem_id not in self.other_message_ids[chat_identifier]:

                            next_is_self = False
                            for next_idx in range(idx + 1, min(idx + 6, len(message_elements))):
                                next_elem = message_elements[next_idx]
                                next_elem_text = next_elem.name.strip()
                                next_elem_rect = next_elem.rectangle
                                next_pos_sender = "other" if next_elem_rect.left < window_center_x else "self"
                                next_keyword_sender = "other" if any(
                                    k in next_elem.name.lower() for k in self.other_msg_keywords) else "self" if any(
                                    k in next_elem.name.lower() for k in self.self_msg_keywords) else None
                                next_sender = next_keyword_sender if next_keyword_sender is not None else next_pos_sender
                                if next_sender == "self":
                                    next_is_self = True
                                    self.log(
                                        f"[{chat_identifier}] 消息（序号{idx}）: {elem_text} 后有自己的消息（序号{next_idx}）: {next_elem_text}，跳过回复")
                                    break
                            if not next_is_self:
                                self.log(f"[{chat_identifier}] 发现新消息（序号{idx}）: {elem_text}（发送者：{sender}）")
                                new_other_messages.append(elem_text)
                                self.other_message_ids[chat_identifier].add(elem_id)
                            else:
                                self.other_message_ids[chat_identifier].add(elem_id)  # 仍记录消息ID以避免重复
                        else:
                            self.log(f"[{chat_identifier}] 消息已处理（序号{idx}）: {elem_text}")
                    else:
                        self.log(f"[{chat_identifier}] 跳过自己的消息（序号{idx}）: {elem_text}")
                except Exception as e:
                    self.log(f"处理消息元素（序号{idx}）出错: {str(e)}")
                    continue
            self.log(f"[{chat_identifier}] 发现对方新消息总数: {len(new_other_messages)}")
            return new_other_messages
        except Exception as e:
            self.log(f"获取新消息失败: {str(e)}")
            return []

    def is_element_in_active_chat(self, elem):
        if not self.active_chat_element:
            self.log("未获取到活动聊天框元素，默认视为当前聊天消息")
            return True
        try:
            elem_rect = elem.rectangle
            active_rect = self.active_chat_element.rectangle
            return (elem_rect.left >= active_rect.left - 10 and
                    elem_rect.top >= active_rect.top - 10 and
                    elem_rect.right <= active_rect.right + 10 and
                    elem_rect.bottom <= active_rect.bottom + 10)
        except Exception as e:
            self.log(f"判断元素是否在活动聊天框出错: {str(e)}")
            return True

    def send_reply(self, message):
        try:
            self.log(f"准备发送回复: {message}")
            if not self.ensure_active_chat_focused():
                self.log("无法聚焦到活动聊天框，尝试重试...")
                if not self.ensure_active_chat_focused():
                    self.log("聚焦失败，放弃发送")
                    return False
            input_elem = None
            for elem in self.wechat_element.descendants():

                if not self.is_running:
                    return False
                try:
                    if (elem.control_type in ["Edit", "Text"] and
                            elem.is_enabled and
                            self.is_element_visible(elem) and
                            ("输入" in elem.name or "send" in elem.name.lower() or "message" in elem.name.lower())):
                        input_elem = elem
                        self.log(f"找到输入框: {elem.name}")
                        break
                except:
                    continue
            if input_elem:
                try:
                    input_elem.set_focus()

                    for _ in range(1):
                        if not self.is_running:
                            return False
                        time.sleep(1)
                    send_keys(message)

                    for _ in range(1):
                        if not self.is_running:
                            return False
                        time.sleep(1)
                    send_keys("{ENTER}")
                    self.log(f"成功发送回复: {message}")
                    return True
                except Exception as e:
                    self.log(f"输入框发送失败，尝试备用方案: {str(e)}")
            self.log("使用备用方案发送回复")
            win32gui.SetForegroundWindow(self.wechat_hwnd)

            for _ in range(1):
                if not self.is_running:
                    return False
                time.sleep(1)
            rect = win32gui.GetWindowRect(self.wechat_hwnd)
            x = rect[0] + (rect[2] - rect[0]) // 2
            y = rect[3] + self.input_offset_y
            x = max(0, min(x, win32api.GetSystemMetrics(0) - 1))
            y = max(0, min(y, win32api.GetSystemMetrics(1) - 1))
            self.log(f"点击输入区域坐标: ({x}, {y})")
            win32api.SetCursorPos((x, y))

            for _ in range(1):
                if not self.is_running:
                    return False
                time.sleep(0.5)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)

            for _ in range(1):
                if not self.is_running:
                    return False
                time.sleep(0.3)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)

            for _ in range(2):
                if not self.is_running:
                    return False
                time.sleep(0.75)
            send_keys(message)

            for _ in range(1):
                if not self.is_running:
                    return False
                time.sleep(1)
            send_keys("{ENTER}")
            self.log(f"备用方案发送回复成功: {message}")
            return True
        except Exception as e:
            self.log(f"发送回复失败: {str(e)}")
            return False

    def ensure_active_chat_focused(self):
        if not self.active_chat_element:
            self.log("无活动聊天框元素，默认聚焦成功")
            return True
        try:
            rect = self.active_chat_element.rectangle
            x = rect.left + rect.width() // 2
            y = rect.top + rect.height() // 3
            win32api.SetCursorPos((x, y))

            for _ in range(1):
                if not self.is_running:
                    return False
                time.sleep(0.5)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)

            for _ in range(1):
                if not self.is_running:
                    return False
                time.sleep(0.2)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)

            for _ in range(1):
                if not self.is_running:
                    return False
                time.sleep(1)
            return True
        except Exception as e:
            self.log(f"确保活动聊天框聚焦失败: {str(e)}")
            return False

    def generate_reply(self, message):
        self.log(f"为消息生成智能回复: {message}")

        # 豆包API配置（参考：https://www.doubao.com/openapi）
        DOUBAO_API_KEY = "your_doubao_api_key"  # 替换为实际豆包API密钥
        DOUBAO_API_URL = "https://api.doubao.com/chat/completions"
        # DeepSeek API配置（参考：https://www.deepseek.com/openapi）
        DEEPSEEK_API_KEY = "your_deepseek_api_key"  # 替换为实际DeepSeek API密钥
        DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"

        try:

            self.log("尝试调用豆包API生成回复...")
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {DOUBAO_API_KEY}"
            }
            payload = {
                "model": "doubao-pro",  # 豆包模型名称（根据实际版本调整）
                "messages": [
                    {"role": "system", "content": "你是一个微信自动回复助手，回复简洁友好，符合日常聊天语境。"},
                    {"role": "user", "content": message}
                ],
                "max_tokens": 50,
                "temperature": 0.7
            }
            response = requests.post(
                DOUBAO_API_URL,
                headers=headers,
                data=json.dumps(payload),
                timeout=5
            )
            response.raise_for_status()
            reply = response.json()["choices"][0]["message"]["content"].strip()
            self.log(f"豆包API生成回复: {reply}")
            return reply
        except Exception as e:
            self.log(f"豆包API调用失败: {str(e)}，尝试DeepSeek API...")
            try:

                headers = {
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {DEEPSEEK_API_KEY}"
                }
                payload = {
                    "model": "deepseek-chat",  # DeepSeek模型名称（根据实际版本调整）
                    "messages": [
                        {"role": "system", "content": "简洁回复微信消息，语气自然。"},
                        {"role": "user", "content": message}
                    ],
                    "max_tokens": 50,
                    "temperature": 0.7
                }
                response = requests.post(
                    DEEPSEEK_API_URL,
                    headers=headers,
                    data=json.dumps(payload),
                    timeout=5
                )
                response.raise_for_status()
                reply = response.json()["choices"][0]["message"]["content"].strip()
                self.log(f"DeepSeek API生成回复: {reply}")
                return reply
            except Exception as e:
                self.log(f"DeepSeek API调用失败: {str(e)}，使用备用回复...")

                fallback_replies = {

                    "greeting": [
                        "你好呀~ 刚看到消息",
                        "嗨！收到你的问候啦",
                        "哈喽，刚注意到你的消息",
                        "你好呀，消息我看到了",
                        "早呀！刚看到你的问候",
                        "晚上好~ 消息收到了",
                        "嗨，刚看到，最近怎么样？",
                        "你好呀，刚上线就看到了"
                    ],

                    "question": [
                        "你的问题我看到了，稍等回复",
                        "刚看到你的提问，马上解答",
                        "收到你的疑问，这就回复你",
                        "看到你的问题了，稍后详细说",
                        "这个问题我记下了，马上回复",
                        "刚看到你的疑问，正在了解",
                        "你的问题我看到了，整理下思路回复你",
                        "收到，这就为你解答"
                    ],

                    "notification": [
                        "收到你的通知啦，谢谢告知",
                        "嗯嗯，这个我知道了",
                        "好的，已收到你的分享",
                        "消息看到了，感谢通知",
                        "谢谢你的分享，刚看到",
                        "收到这个消息了，谢谢提醒",
                        "好的，这个情况我了解了",
                        "已收到通知，感谢告知"
                    ],

                    "request": [
                        "你的请求我看到了，尽力帮忙",
                        "收到你的求助，马上看看",
                        "刚注意到你的请求，这就处理",
                        "看到了，我来帮你想想办法",
                        "你的需求我收到了，尽量协助",
                        "好的，我来看看怎么帮你",
                        "收到，这就处理你的请求",
                        "没问题，我来想办法解决"
                    ],

                    "invitation": [
                        "收到你的邀约了，稍等回复",
                        "这个安排我看到了，考虑下回复你",
                        "邀约收到啦，稍后给你答复",
                        "看到你的提议了，这就看看时间",
                        "好的，这个邀约我记下了",
                        "收到，我看看日程再回复你",
                        "邀约已看到，稍后确认回复",
                        "你的提议我收到了，考虑后回复"
                    ],

                    "thanks": [
                        "不客气呀~ 应该的",
                        "不用谢，能帮到你就好",
                        "没事没事，不用客气",
                        "不客气，应该做的",
                        "不用谢呀，小事一桩",
                        "没事，不客气呢",
                        "不用客气，很高兴能帮到你"
                    ],

                    "apology": [
                        "没事的，别往心里去",
                        "没关系，我理解的",
                        "没事没事，不用在意",
                        "不要紧的，小事而已",
                        "没关系呀，别自责了",
                        "没事，我不介意的",
                        "没关系，过去了就好"
                    ],

                    "general": [
                        "收到你的消息了",
                        "好的，我看到了",
                        "消息收到啦，稍后回复",
                        "刚看到，马上来处理",
                        "嗯呢，消息已接收",
                        "已读你的消息",
                        "看到了，这就来",
                        "收到啦，稍等片刻",
                        "消息我看到了",
                        "刚注意到消息，马上回复",
                        "收到消息，稍后回复"
                    ]
                }

                def get_fallback_reply(message):
                    message = message.lower()

                    if any(thanks in message for thanks in ["谢谢", "多谢", "感谢", "谢啦", "辛苦了"]):
                        return random.choice(fallback_replies["thanks"])

                    if any(apology in message for apology in ["对不起", "抱歉", "不好意思", "抱歉了", "歉"]):
                        return random.choice(fallback_replies["apology"])

                    if any(greet in message for greet in
                           ["你好", "嗨", "哈喽", "早", "晚", "早上好", "晚上好", "下午好", "最近好吗"]):
                        return random.choice(fallback_replies["greeting"])

                    if any(invite in message for invite in
                           ["约", "一起", "聚会", "吃饭", "时间", "见面", "有空吗", "周末", "哪天"]):
                        return random.choice(fallback_replies["invitation"])

                    if any(req in message for req in
                           ["帮", "助", "请求", "麻烦", "能否", "可以吗", "能不能", "帮个忙"]):
                        return random.choice(fallback_replies["request"])

                    if any(question in message for question in
                           ["？", "吗", "为什么", "怎么", "如何", "哪里", "多少", "何时", "什么"]):
                        return random.choice(fallback_replies["question"])

                    return random.choice(fallback_replies["notification"] + fallback_replies["general"])
                return get_fallback_reply(message)

    def get_chat_list_container(self):
        if self.chat_list_container:
            return self.chat_list_container
        try:
            for elem in self.wechat_element.descendants():

                if not self.is_running:
                    return None
                try:
                    if (elem.control_type == "List" and
                            self.is_element_visible(elem) and
                            elem.rectangle.width() < 300 and
                            elem.rectangle.left < 300):
                        chat_items = [child for child in elem.children()
                                      if child.control_type == "ListItem" and self.is_element_visible(child)]
                        if len(chat_items) > 0:
                            self.chat_list_container = elem
                            self.log(f"找到聊天列表容器，包含 {len(chat_items)} 个聊天项")
                            return elem
                except:
                    continue
            self.log("严格匹配失败，尝试宽松匹配聊天列表")
            for elem in self.wechat_element.descendants():

                if not self.is_running:
                    return None
                try:
                    if elem.control_type == "List" and self.is_element_visible(elem):
                        chat_items = [child for child in elem.children() if child.control_type == "ListItem"]
                        if len(chat_items) > 0:
                            self.chat_list_container = elem
                            self.log(f"宽松匹配找到聊天列表，包含 {len(chat_items)} 个聊天项")
                            return elem
                except:
                    continue
            self.log("未找到聊天列表容器")
            return None
        except Exception as e:
            self.log(f"获取聊天列表失败: {str(e)}")
            return None

    def get_suspected_chats(self):
        try:
            chat_list = self.get_chat_list_container()
            if not chat_list:
                return []
            unread_chats = []
            normal_chats = []
            for chat_item in chat_list.children():

                if not self.is_running:
                    return []
                try:
                    if (chat_item.control_type != "ListItem" or
                            not self.is_element_visible(chat_item) or
                            not chat_item.name.strip()):
                        continue

                    chat_name = chat_item.name.strip().lower()
                    if any(keyword in chat_name for keyword in self.public_account_keywords):
                        self.log(f"跳过公众号聊天: {chat_item.name}")
                        continue
                    has_unread = False
                    for child in chat_item.descendants():

                        if not self.is_running:
                            return []
                        try:
                            if (child.control_type == "Image" and
                                    self.is_element_visible(child) and
                                    ("未读" in child.name or child.rectangle.width() < 20)):
                                has_unread = True
                                break
                        except:
                            continue
                    if has_unread:
                        unread_chats.append((chat_item, chat_item.name.strip()))
                    else:
                        normal_chats.append((chat_item, chat_item.name.strip()))
                except:
                    continue
            suspected_chats = unread_chats + normal_chats
            result = suspected_chats[:self.max_chat_count]
            self.log(f"获取疑似有消息的聊天（含未读）: {[c[1] for c in result]}")
            return result
        except Exception as e:
            self.log(f"获取疑似有消息的聊天失败: {str(e)}")
            return []

    def switch_to_chat(self, chat_item, chat_identifier):
        try:
            if self.current_chat_identifier == chat_identifier:
                self.log(f"已在 [{chat_identifier}] 聊天框，无需切换")
                return True
            self.log(f"开始切换到 [{chat_identifier}] 聊天框")
            win32gui.SetForegroundWindow(self.wechat_hwnd)

            for _ in range(1):
                if not self.is_running:
                    return False
                time.sleep(1)
            rect = chat_item.rectangle
            x = rect.left + self.chat_switch_offset_x
            y = rect.top + rect.height() // 2
            x = max(rect.left + 10, min(x, rect.right - 10))
            y = max(rect.top + 10, min(y, rect.bottom - 10))
            self.log(f"点击聊天项坐标: ({x}, {y})")
            win32api.SetCursorPos((x, y))

            for _ in range(1):
                if not self.is_running:
                    return False
                time.sleep(0.5)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)

            for _ in range(1):
                if not self.is_running:
                    return False
                time.sleep(0.2)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)

            for _ in range(3):
                if not self.is_running:
                    return False
                time.sleep(1)
            self.current_chat_identifier = chat_identifier
            self.active_chat_element = self.get_active_chat_element()
            self.log(f"切换到 [{chat_identifier}] 完成")
            return True
        except Exception as e:
            self.log(f"切换到 [{chat_identifier}] 失败: {str(e)}")
            return False

    def get_active_chat_element(self):
        try:
            for elem in self.wechat_element.descendants():

                if not self.is_running:
                    return None
                try:
                    if (elem.control_type == "Pane" and
                            self.is_element_visible(elem) and
                            elem.rectangle.left > 300 and
                            elem.rectangle.width() > 600 and
                            elem.rectangle.height() > 600):
                        return elem
                except:
                    continue
            for elem in self.wechat_element.descendants():

                if not self.is_running:
                    return None
                try:
                    if (elem.control_type in ["Pane", "List"] and
                            self.is_element_visible(elem) and
                            elem.rectangle.width() > 500 and
                            elem.rectangle.height() > 500):
                        return elem
                except:
                    continue
            self.log("未找到活动聊天框元素")
            return None
        except:
            return None


    def monitor_thread(self):
        if not self.connect_wechat():
            self.append_log("无法连接微信，监控停止")
            self.is_running = False
            self.start_btn.config(text="启动监控")
            self.status_var.set("未运行")
            return
        self.append_log(f"微信消息监控已启动，最多处理{self.max_chat_count}个聊天")
        try:
            while self.is_running:
                self.append_log("\n===== 开始新一轮聊天检测 =====")
                suspected_chats = self.get_suspected_chats()
                if not suspected_chats:
                    self.append_log("未找到任何聊天项，5秒后重试...")

                    for _ in range(5):
                        if not self.is_running:
                            break
                        time.sleep(1)
                    continue
                for idx, (chat_item, chat_identifier) in enumerate(suspected_chats, 1):

                    if not self.is_running:
                        break
                    self.append_log(f"\n----- 处理第{idx}/{self.max_chat_count}个聊天: {chat_identifier} -----")
                    if not self.switch_to_chat(chat_item, chat_identifier):
                        self.append_log(f"❌ 切换到 {chat_identifier} 失败，跳过")
                        continue
                    new_messages = self.get_new_messages(chat_identifier)
                    if new_messages:
                        self.append_log(f"✅ 发现 {len(new_messages)} 条新消息")
                        for msg in new_messages:

                            if not self.is_running:
                                break
                            self.append_log(f"收到消息: {msg}")
                            reply = self.generate_reply(msg)
                            if self.send_reply(reply):
                                self.append_log(f"已回复: {reply}")
                            else:
                                self.append_log(f"❌ 回复发送失败")
                    else:
                        self.append_log(f"ℹ️ 未发现新消息")
                    if idx >= self.max_chat_count:
                        break
                self.append_log("\n===== 本轮检测结束，10秒后再次检测 =====")

                for _ in range(10):
                    if not self.is_running:
                        break
                    time.sleep(1)
        except Exception as e:
            self.append_log(f"监控异常停止: {str(e)}")
        finally:
            self.is_running = False
            self.start_btn.config(text="启动监控")
            self.status_var.set("未运行")


    def toggle_monitor(self):
        if not self.is_running:
            self.is_running = True
            self.start_btn.config(text="停止监控")
            self.status_var.set("运行中")

            threading.Thread(target=self.monitor_thread, daemon=True).start()
        else:
            self.is_running = False
            self.start_btn.config(text="启动监控")
            self.status_var.set("停止中...")


    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    try:
        import pyautogui
    except ImportError:
        import os
        os.system("pip install pyautogui")
        import pyautogui
    bot = WeChatAutoReply()
    bot.run()