import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import requests
import win32com.client
import pythoncom  # 添加pythoncom导入
from apscheduler.schedulers.background import BackgroundScheduler
from datetime import datetime
import json
import time
import threading
import os
import pystray
from pystray import MenuItem as item
from PIL import Image
import sys

class RawMaterialMailerApp:
    def __init__(self, root):
        self.path = self.resource_path(os.path.dirname(os.path.abspath(__file__)))
        self.root = root
        icon = tk.PhotoImage(file=self.resource_path("mail.png"))
        self.root.iconphoto(True, icon)
        self.root.title("原料申请邮件发送工具")
        self.root.geometry("800x600")

        # APScheduler 定时任务调度器
        self.scheduler = BackgroundScheduler()
        self.scheduler.start()

        # 读取配置并直接存储为实例属性
        self.load_config()

        # 初始化token和headers
        self.token = None
        self.headers = {}

        # 设置UI
        self.setup_ui()

        # 自动获取token
        self.get_token()

        # 如果配置为自动启动，则启动定时任务
        if self.auto_start:
            self.start_scheduler()

        # 初始化系统托盘
        self.setup_system_tray()

    def resource_path(self, relative_path):
        """获取资源文件的绝对路径，用于PyInstaller打包后正确访问资源文件"""
        try:
            # PyInstaller创建临时文件夹，并将路径存储在 _MEIPASS 中
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        
        return os.path.join(base_path, relative_path)

    def setup_system_tray(self):
        """设置系统托盘图标"""
        # 创建一个简单的图像作为托盘图标
        # 如果没有图标文件，创建一个默认图像
        try:
            image = Image.open(self.resource_path("mail.png"))
        except:
            # 如果没有图标文件，创建一个默认的简单图像
            image = Image.new('RGB', (64, 64), color = (73, 109, 137))
        
        # 创建系统托盘菜单
        menu = (item('显示', self.show_window), item('退出', self.quit_app))
        
        # 创建系统托盘图标
        self.tray_icon = pystray.Icon("原料申请邮件发送工具", image, "原料申请邮件发送工具", menu)
        
        # 启动系统托盘图标
        self.tray_icon_thread = threading.Thread(target=self.tray_icon.run)
        self.tray_icon_thread.daemon = True
        self.tray_icon_thread.start()

    def show_window(self, icon, item):
        """显示主窗口"""
        self.root.after(0, self._show_window)

    def _show_window(self):
        """实际显示窗口的方法"""
        self.root.deiconify()  # 显示窗口
        self.root.lift()  # 将窗口置于最前
        self.root.focus_force()  # 强制获得焦点

    def hide_window(self):
        """隐藏主窗口到系统托盘"""
        self.root.withdraw()  # 隐藏窗口

    def quit_app(self, icon, item):
        """退出应用程序"""
        self.on_closing()
        self.tray_icon.stop()
        os._exit(0)

    def setup_ui(self):
        # 配置区域
        config_frame = ttk.LabelFrame(self.root, text="服务端配置", padding=10)
        config_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(config_frame, text="登录地址:").grid(
            row=0, column=0, sticky=tk.W, pady=2
        )
        self.server_url_for_login_var = tk.StringVar(value=self.server_url_for_login)
        self.server_url_for_login_entry = ttk.Entry(
            config_frame, textvariable=self.server_url_for_login_var, width=50
        )
        self.server_url_for_login_entry.grid(
            row=0, column=1, columnspan=2, sticky=tk.EW, padx=(5, 0), pady=2
        )

        ttk.Label(config_frame, text="邮件发送地址:").grid(
            row=1, column=0, sticky=tk.W, pady=2
        )
        self.server_url_for_send_mail_var = tk.StringVar(
            value=self.server_url_for_send_mail
        )
        self.server_url_for_send_mail_entry = ttk.Entry(
            config_frame, textvariable=self.server_url_for_send_mail_var, width=50
        )
        self.server_url_for_send_mail_entry.grid(
            row=1, column=1, columnspan=2, sticky=tk.EW, padx=(5, 0), pady=2
        )

        ttk.Label(config_frame, text="抄送列表地址:").grid(
            row=2, column=0, sticky=tk.W, pady=2
        )
        self.server_url_for_copy_list_var = tk.StringVar(
            value=self.server_url_for_copy_list
        )
        self.server_url_for_copy_list_entry = ttk.Entry(
            config_frame, textvariable=self.server_url_for_copy_list_var, width=50
        )
        self.server_url_for_copy_list_entry.grid(
            row=2, column=1, columnspan=2, sticky=tk.EW, padx=(5, 0), pady=2
        )

        ttk.Label(config_frame, text="用户名:").grid(
            row=3, column=0, sticky=tk.W, pady=2
        )
        self.username_var = tk.StringVar(value=self.username)
        self.username_entry = ttk.Entry(
            config_frame, textvariable=self.username_var, width=50
        )
        self.username_entry.grid(
            row=3, column=1, columnspan=2, sticky=tk.EW, padx=(5, 0), pady=2
        )

        ttk.Label(config_frame, text="密码:").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.password_var = tk.StringVar(value=self.password)
        self.password_entry = ttk.Entry(
            config_frame, textvariable=self.password_var, width=50, show="*"
        )
        self.password_entry.grid(
            row=4, column=1, columnspan=2, sticky=tk.EW, padx=(5, 0), pady=2
        )

        ttk.Label(config_frame, text="Token:").grid(
            row=5, column=0, sticky=tk.W, pady=2
        )
        self.token_var = tk.StringVar()
        self.token_entry = ttk.Entry(
            config_frame, textvariable=self.token_var, width=50, state="readonly"
        )
        self.token_entry.grid(
            row=5, column=1, columnspan=2, sticky=tk.EW, padx=(5, 0), pady=2
        )

        ttk.Label(config_frame, text="收件人:").grid(
            row=6, column=0, sticky=tk.W, pady=2
        )
        self.recipient_var = tk.StringVar(value=self.recipient)
        self.recipient_entry = ttk.Entry(
            config_frame, textvariable=self.recipient_var, width=50
        )
        self.recipient_entry.grid(
            row=6, column=1, columnspan=2, sticky=tk.EW, padx=(5, 0), pady=2
        )

        config_frame.columnconfigure(1, weight=1)

        # 定时任务配置区域
        schedule_frame = ttk.LabelFrame(self.root, text="定时任务配置", padding=10)
        schedule_frame.pack(fill=tk.X, padx=10, pady=5)

        # 执行频率配置
        ttk.Label(schedule_frame, text="执行频率 (周):").grid(
            row=0, column=0, sticky=tk.W, pady=2
        )
        self.interval_weeks_var = tk.StringVar(value=str(self.interval_weeks))
        self.interval_weeks_combo = ttk.Combobox(
            schedule_frame,
            textvariable=self.interval_weeks_var,
            values=["1", "2", "3", "4"],
            width=8,
            state="readonly",
        )
        self.interval_weeks_combo.grid(
            row=0, column=1, sticky=tk.W, padx=(5, 0), pady=2
        )

        # 执行日配置
        ttk.Label(schedule_frame, text="执行日:").grid(
            row=0, column=2, sticky=tk.W, padx=(20, 0), pady=2
        )
        self.interval_days_var = tk.StringVar(value=str(self.interval_days))
        self.interval_days_combo = ttk.Combobox(
            schedule_frame,
            textvariable=self.interval_days_var,
            values=["1", "2", "3", "4", "5", "6", "7"],
            width=8,
            state="readonly",
        )
        self.interval_days_combo.grid(row=0, column=3, sticky=tk.W, padx=(5, 0), pady=2)

        # 添加日名称标签
        day_names = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
        day_name = (
            day_names[int(self.interval_days_var.get()) - 1]
            if self.interval_days_var.get().isdigit()
            else "周日"
        )
        self.day_name_label = ttk.Label(schedule_frame, text=f"({day_name})")
        self.day_name_label.grid(row=0, column=4, sticky=tk.W, padx=(5, 0), pady=2)
        
        # 执行时间配置
        ttk.Label(schedule_frame, text="执行时间:").grid(
            row=0, column=5, sticky=tk.W, padx=(20, 0), pady=2
        )
        # 添加小时选择
        self.hour_var = tk.StringVar(value=str(self.hour))
        self.hour_combo = ttk.Combobox(
            schedule_frame,
            textvariable=self.hour_var,
            values=[f"{i:02d}" for i in range(24)],  # 00-23小时
            width=8,
            state="readonly",
        )
        self.hour_combo.grid(row=0, column=6, sticky=tk.W, padx=(5, 0), pady=2)
        
        ttk.Label(schedule_frame, text=":").grid(
            row=0, column=7, sticky=tk.W, pady=2
        )
        # 添加分钟选择
        self.minute_var = tk.StringVar(value=str(self.minute))
        self.minute_combo = ttk.Combobox(
            schedule_frame,
            textvariable=self.minute_var,
            values=[f"{i:02d}" for i in range(60)],  # 00-59分钟
            width=8,
            state="readonly",
        )
        self.minute_combo.grid(row=0, column=8, sticky=tk.W, padx=(5, 0), pady=2)

        self.interval_days_var.trace("w", self.update_day_name)  # 监听选择变化

        # 将自动开始复选框移到下一行
        self.auto_start_var = tk.BooleanVar(value=self.auto_start)
        self.auto_start_check = ttk.Checkbutton(
            schedule_frame, text="启动时自动开始定时任务", variable=self.auto_start_var
        )
        self.auto_start_check.grid(
            row=1, column=0, sticky=tk.W, padx=(0, 0), pady=(5, 0), columnspan=2
        )

        # 控制按钮
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=10, pady=5)

        self.save_config_button = ttk.Button(
            button_frame, text="保存配置", command=self.save_config
        )
        self.save_config_button.pack(side=tk.LEFT, padx=5)

        self.start_button = ttk.Button(
            button_frame, text="开始定时任务", command=self.start_scheduler
        )
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.stop_button = ttk.Button(
            button_frame,
            text="停止定时任务",
            command=self.stop_scheduler,
            state=tk.DISABLED,
        )
        self.stop_button.pack(side=tk.LEFT, padx=5)

        self.manual_send_button = ttk.Button(
            button_frame, text="手动发送", command=self.manual_send
        )
        self.manual_send_button.pack(side=tk.LEFT, padx=5)

        # 状态显示
        status_frame = ttk.LabelFrame(self.root, text="状态", padding=10)
        status_frame.pack(fill=tk.X, padx=10, pady=5)

        self.status_label = ttk.Label(status_frame, text="就绪")
        self.status_label.pack(anchor=tk.W)

        # 日志显示
        log_frame = ttk.LabelFrame(self.root, text="日志", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=15)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 绑定窗口关闭事件，改为最小化到系统托盘
        self.root.protocol("WM_DELETE_WINDOW", self.on_window_close)

        # 初始化定时任务
        if self.auto_start_var.get():
            self.start_scheduler()

    def on_window_close(self):
        """窗口关闭事件 - 最小化到系统托盘"""
        self.hide_window()
        messagebox.showinfo("提示", "程序已最小化到系统托盘，双击托盘图标可恢复显示")

    def update_day_name(self, *args):
        """更新星期名称显示"""
        day_names = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
        try:
            day_index = int(self.interval_days_var.get()) - 1
            if 0 <= day_index <= 6:
                day_name = day_names[day_index]
                self.day_name_label.config(text=f"({day_name})")
        except ValueError:
            pass

    def load_config(self):
        """加载配置文件并直接存储为实例属性"""
        config_file = self.resource_path("config.json")  # 修改：使用resource_path获取配置文件路径
        if os.path.exists(config_file):
            try:
                with open(config_file, "r", encoding="utf-8") as f:
                    config = json.load(f)

                # 直接存储为实例属性
                self.server_url_for_login = config.get("server_url_for_login", "")
                self.server_url_for_send_mail = config.get(
                    "server_url_for_send_mail", ""
                )
                self.server_url_for_copy_list = config.get(
                    "server_url_for_copy_list", ""
                )
                self.username = config.get("username", "")
                self.password = config.get("password", "")
                self.interval_weeks = config.get("interval_weeks", 2)
                self.interval_days = config.get("interval_days", 2)
                self.auto_start = config.get("auto_start", False)
                self.recipient = config.get("recipient", "xu.li@ofi.com")  # 新增收件人配置
                self.hour = config.get("hour", 9)  # 新增小时配置
                self.minute = config.get("minute", 0)  # 新增分钟配置

                return config
            except Exception as e:
                messagebox.showinfo(f"错误", f"配置文件载入错误：{e}")
        else:
            # 设置默认值
            self.server_url_for_login = ""
            self.server_url_for_send_mail = ""
            self.server_url_for_copy_list = ""
            self.username = ""
            self.password = ""
            self.interval_weeks = 2
            self.interval_days = 2
            self.auto_start = False
            self.recipient = ""  # 默认收件人
            self.hour = 9  # 默认小时
            self.minute = 0  # 默认分钟
            # 创建默认配置文件
            self.create_default_config()  # 新增：创建默认配置文件
            messagebox.showinfo("提示", "未找到配置文件，已创建默认配置文件 config.json")

    def create_default_config(self):
        """创建默认配置文件"""
        default_config = {
            "server_url_for_login": "",
            "server_url_for_send_mail": "",
            "server_url_for_copy_list": "",
            "username": "",
            "password": "",
            "interval_weeks": 2,
            "interval_days": 2,
            "auto_start": False,
            "recipient": "xu.li@ofi.com",
            "hour": 9,
            "minute": 0
        }
        config_file = self.resource_path("config.json")  # 使用resource_path获取配置文件路径
        try:
            with open(config_file, "w", encoding="utf-8") as f:
                json.dump(default_config, f, ensure_ascii=False, indent=2)
            self.log_message("默认配置文件创建成功")
        except Exception as e:
            self.log_message(f"创建默认配置文件失败: {e}")

    def save_config(self):
        """保存配置到文件"""
        config = {
            "server_url_for_login": self.server_url_for_login_var.get(),
            "server_url_for_send_mail": self.server_url_for_send_mail_var.get(),
            "server_url_for_copy_list": self.server_url_for_copy_list_var.get(),
            "username": self.username_var.get(),
            "password": self.password_var.get(),
            "interval_weeks": int(self.interval_weeks_var.get()),
            "interval_days": int(self.interval_days_var.get()),
            "auto_start": self.auto_start_var.get(),
            "recipient": self.recipient_var.get(),  # 保存收件人配置
            "hour": int(self.hour_var.get()),  # 保存小时配置
            "minute": int(self.minute_var.get()),  # 保存分钟配置
        }
        try:
            config_file = self.resource_path("config.json")  # 修改：使用resource_path获取配置文件路径
            with open(config_file, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
                self.log_message("配置保存成功")
        except Exception as e:
            self.log_message(f"保存配置文件失败: {e}")

    def log_message(self, message):
        """记录日志"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def update_status(self, message):
        """更新状态栏"""
        self.status_label.config(text=message)
        self.root.update_idletasks()

    def get_token(self):
        # 准备登录数据
        login_data = {
            "username": self.username_var.get().strip(),
            "password": self.password_var.get().strip(),
        }

        if not login_data["username"] or not login_data["password"]:
            raise Exception("用户名和密码不能为空")

        # 发送登录请求 - 直接使用实例属性中的登录URL
        login_url = self.server_url_for_login
        response = requests.post(login_url, json=login_data)
        response.raise_for_status()

        result = response.json()

        # 检查是否登录失败
        if "detail" in result and result["detail"] == "用户名或密码错误":
            raise Exception("用户名或密码错误")

        # 获取token
        self.token = result.get("access_token")  # 根据实际返回格式调整
        self.headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
        }

        if not self.token:
            raise Exception("登录失败：未返回有效token")

        # 更新token显示
        self.token_var.set(self.token)

        self.log_message("认证token获取成功")
        return self.token

    def start_scheduler(self):
        """开始定时任务"""
        # 获取UI中选择的值（这些值已经在UI中限制了选项，所以无需验证）
        interval_weeks = int(self.interval_weeks_var.get())
        interval_day = int(self.interval_days_var.get())
        hour = int(self.hour_var.get())
        minute = int(self.minute_var.get())

        # 先停止现有任务
        self.stop_scheduler()

        # 获取认证token
        self.get_token()

        # 计算星期几（1=周一，7=周日）
        day_names = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
        day_name = day_names[interval_day - 1]

        # 使用cron表达式来实现每周间隔的定时任务
        # 例如：每2周的周五执行一次，cron表达式为：0 0 * * 5/2
        # 这里我们使用 week_of_month 参数来实现
        self.scheduler.add_job(
            self.send_mail_job,
            "cron",
            day_of_week=day_name,
            week="*/" + str(interval_weeks),
            hour=hour,
            minute=minute,
            id="raw_material_mail_job",
        )

        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        self.log_message(
            f"定时任务已启动，每 {interval_weeks} 周的{['周一', '周二', '周三', '周四', '周五', '周六', '周日'][interval_day-1]}执行一次"
        )
        self.update_status(
            f"定时任务运行中，下次执行: {self.scheduler.get_job('raw_material_mail_job').next_run_time}"
        )

    def stop_scheduler(self):
        """停止定时任务"""
        if self.scheduler.get_job("raw_material_mail_job"):
            self.scheduler.remove_job("raw_material_mail_job")

        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

        self.log_message("队列中定时任务已停止")
        self.update_status("队列中定时任务已停止")

    def manual_send(self):
        """手动发送邮件"""
        self.log_message("开始手动发送邮件...")
        self.send_mail_job()

    def get_copy_list(self):
        response = requests.get(self.server_url_for_copy_list, headers=self.headers)
        return response.json()  # 假设返回的是JSON格式

    def send_mail_job(self):
        """发送邮件的主要逻辑"""
        try:
            self.update_status("正在查询申请数据...")

            # 从服务端获取申请数据
            applies = self.get_applies_from_server()

            if len(applies) == 0:
                self.log_message("没有待处理的申请")
                self.update_status("没有待处理的申请")
                return

            self.log_message(f"获取到 {len(applies)} 个待处理申请")

            # 发送邮件
            self.update_status("正在发送邮件...")

            self.send_mail_to_outlook(applies)

            # 更新服务端申请状态
            self.update_applies_status()

            self.log_message(f"邮件发送成功，共处理 {len(applies)} 个申请")
            self.update_status(f"邮件发送完成 - {datetime.now().strftime('%H:%M:%S')}")

        except Exception as e:
            error_msg = f"邮件发送失败: {str(e)}"
            self.log_message(error_msg)
            self.update_status(f"发送失败: {str(e)}")
            messagebox.showerror("错误", error_msg)

    def get_applies_from_server(self):
        """从服务端API获取原料申请数据"""
        # 直接使用实例属性中的邮件发送URL作为API端点
        send_mail_api = self.server_url_for_send_mail

        # 确保有有效的token
        if not self.token:
            self.get_token()

        # 发送GET请求获取申请列表
        response = requests.get(send_mail_api, headers=self.headers)
        response.raise_for_status()

        return response.json()

    def send_mail_to_outlook(self, applies):
        """通过Outlook发送邮件"""
        # 构建邮件内容
        mail_content = ""
        for item in applies:
            mail_content += f"""
            <tr>
            <td><p>{item.get('applyDate', '')}</p></td>
            <td><p>{item.get('rawMaterial_id', '')}</p></td>
            <td><p>{item.get('rawMaterial__name', '')}</p></td>
            <td><p>{item.get('rawMaterial__sapCode', '')}</p></td>
            <td><p>{item.get('applier__username', '')}</p></td>
            <td><p>{item.get('qty', '')}</p></td>
            </tr>
            """

        # 初始化COM库
        pythoncom.CoInitialize()
        try:
            # 创建Outlook应用实例
            outlook = win32com.client.Dispatch("outlook.application")
            mail = outlook.CreateItem(0)  # 0表示邮件项

            # 设置收件人（目标邮件地址预设）
            mail.To = self.recipient_var.get()  # 使用可配置的收件人地址
            copy_list = self.get_copy_list()
            cc_list = [
                item.get("email") for item in copy_list if item.get("email") is not None
            ]  # 修复字段名拼写错误
            mail.CC = ";".join(cc_list)
            # # 设置邮件格式和主题
            mail.BodyFormat = 2  # HTML格式
            mail.Subject = f"{datetime.today().year}年{datetime.today().month}月原料申请"

            # 设置HTML邮件内容
            mail.HTMLBody = f"""
                <p>hi，Team，</p>
                <p>请帮忙安排以下原料：</p>
                <table style="width: 800px; border: 1px solid black;">
                <tbody>
                <tr>
                <th><p>申请日期</p></th>
                <th><p>原料编号</p></th>
                <th><p>原料名称</p></th>
                <th><p>sapCode</p></th>
                <th><p>申请人</p></th>
                <th><p>申请量(kg)</p></th>
                </tr>
                {mail_content}
                </tbody>
                </table>
                <br>
                <span style="font-size:16px">此信息于每周五由系统自动发送！</span><br>
                <span style="font-size:16px">上海同事需要接收此信息，请将邮箱添加到系统的个人信息-我的资料中！</span><br>
            """

            # 设置邮件重要性
            mail.Importance = 1  # 1为低，2为普通，3为高

            # 发送邮件
            mail.Send()
        finally:
            # 释放COM库资源
            pythoncom.CoUninitialize()

    def update_applies_status(self):
        try:
            self.get_token()
            response = requests.post(
                self.server_url_for_send_mail,
                headers=self.headers,
                json={
                    "success": True,
                    "submitDate": datetime.today().strftime("%Y-%m-%d"),
                },
            )
            response.raise_for_status()
            self.log_message("申请状态已更新到服务端")
        except Exception as e:
            self.log_message(f"更新申请状态失败: {e}")

    def on_closing(self):
        """应用关闭时的清理工作"""
        if self.scheduler.running:
            self.scheduler.shutdown()
        self.root.destroy()
        if hasattr(self, 'tray_icon'):
            self.tray_icon.stop()


if __name__ == "__main__":
    root = tk.Tk()
    app = RawMaterialMailerApp(root)
    root.mainloop()