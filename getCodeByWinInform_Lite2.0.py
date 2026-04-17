import sys
import os
import json
import asyncio
import threading
import urllib.request
import subprocess
import pyperclip
import winsound
import tkinter as tk
import queue
import re
import pystray
from PIL import Image, ImageDraw
from winsdk.windows.ui.notifications.management import UserNotificationListener
from winsdk.windows.ui.notifications import NotificationKinds

CONFIG_FILE = 'config.json'
APP_NAME = "VerifyCodeProLite"


def get_exe_path():
    if getattr(sys, 'frozen', False):
        return sys.executable
    return os.path.abspath(__file__)


def get_rel_path(relative_path):
    base_path = os.path.dirname(get_exe_path())
    return os.path.join(base_path, relative_path)

def get_resource_path(relative_path):
    """获取内置资源路径，专用于读取打包进 exe 内部的文件"""
    if getattr(sys, 'frozen', False):
        # 运行在打包后的单文件环境中，资源在 _MEIPASS 临时目录
        base_path = sys._MEIPASS
    else:
        # 运行在开发环境中
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def load_config():
    path = get_rel_path(CONFIG_FILE)
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {
        "LLM_API_KEY": "",
        "LLM_MODEL": "",
        "LLM_BASE_URL": "https://ark.cn-beijing.volces.com/api/v3",
        "EXTRACT_MODE": "REGEX"  # 默认使用本地正则模式
    }


def save_config(config_data):
    with open(get_rel_path(CONFIG_FILE), 'w', encoding='utf-8') as f:
        json.dump(config_data, f, indent=4)


current_config = load_config()


# ---------------- 提取引擎 ----------------

def extract_by_regex(text):
    """本地正则提取规则，智能规避时间/日期干扰"""
    # 1. 预处理：抹除常见的日期、时间干扰项
    text = re.sub(r'\b20\d{2}\b', '', text)  # 去除年份 (如 2024)
    text = re.sub(r'\b\d{1,2}:\d{2}\b', '', text)  # 去除时间 (如 12:30)
    text = re.sub(r'\b\d{1,2}-\d{1,2}\b', '', text)  # 去除日期 (如 11-25)
    text = re.sub(r'\b\d{1,2}月\d{1,2}日\b', '', text)  # 去除中文日期

    # 2. 强特征匹配：寻找关键词后面的 4-6 位字母数字组合
    keyword_match = re.search(r'(验证码|校验码|code|动态码|确认码)[\s:：]*([a-zA-Z0-9]{4,6})\b', text, re.IGNORECASE)
    if keyword_match:
        return keyword_match.group(2)

    # 3. 弱特征兜底：寻找孤立的 4-6 位纯数字或大写字母组合
    fallback_matches = re.findall(r'\b([0-9]{4,6}|[A-Z0-9]{4,6})\b', text)
    for m in fallback_matches:
        # 排除纯小写字母（大概率是普通英文单词）
        if not (m.isalpha() and m.islower()):
            return m

    return None


def call_llm(text):
    """云端大模型提取规则"""
    api_key = current_config.get("LLM_API_KEY", "")
    model = current_config.get("LLM_MODEL", "")
    base_url = current_config.get("LLM_BASE_URL", "")

    if not api_key or not model:
        return None

    prompt = f"请从以下系统通知纯文本中提取验证码(通常是4-6位数字/字母)。要求：只输出验证码本身，绝无任何废话。避开日期和网址。如果没有验证码，严格输出 NONE。通知内容：{text}"
    url = f"{base_url}/chat/completions"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.0,
        "max_tokens": 20
    }

    try:
        req = urllib.request.Request(url, data=json.dumps(payload).encode('utf-8'), headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=5) as response:
            result = json.loads(response.read().decode('utf-8'))
            ans = result.get('choices', [{}])[0].get('message', {}).get('content', '').strip()
            if "NONE" in ans.upper() or len(ans) > 10:
                return None
            return ans
    except Exception:
        return None


# ---------------- 后台监控核心 ----------------

class MonitorTask:
    def __init__(self):
        self.running = False

    async def monitor(self):
        listener = UserNotificationListener.current
        access = await listener.request_access_async()
        if access != 1:
            return

        known_ids = set()
        for n in await listener.get_notifications_async(NotificationKinds.TOAST):
            known_ids.add(n.id)

        TARGET_APPS = ["QQ", "邮件", "Mail", "Outlook", "微信", "Chrome"]

        while self.running:
            try:
                notifications = await listener.get_notifications_async(NotificationKinds.TOAST)
                for n in notifications:
                    if n.id not in known_ids:
                        app_name = n.app_info.display_info.display_name
                        if any(target in app_name for target in TARGET_APPS):
                            bindings = n.notification.visual.bindings
                            if len(bindings) > 0:
                                texts = bindings[0].get_text_elements()
                                full_text = " | ".join([t.text for t in texts if t.text])
                                if full_text:
                                    # 根据用户配置动态选择提取引擎
                                    mode = current_config.get("EXTRACT_MODE", "REGEX")
                                    if mode == "LLM":
                                        code = call_llm(full_text)
                                    else:
                                        code = extract_by_regex(full_text)

                                    if code:
                                        pyperclip.copy(code)
                                        winsound.MessageBeep(winsound.MB_OK)
                        known_ids.add(n.id)
            except Exception:
                pass
            await asyncio.sleep(0.5)

    def start(self):
        self.running = True
        asyncio.run(self.monitor())

    def stop(self):
        self.running = False


monitor_instance = MonitorTask()


def run_monitor_thread():
    monitor_instance.start()


# ---------------- UI 与 系统托盘 ----------------

ui_queue = queue.Queue()


def set_autostart(enable):
    startup = os.path.join(os.getenv('APPDATA'), r'Microsoft\Windows\Start Menu\Programs\Startup')
    path = os.path.join(startup, f"{APP_NAME}.lnk")
    target = get_exe_path()
    work_dir = os.path.dirname(target)

    if enable:
        ps_script = f"""
        $wshell = New-Object -ComObject WScript.Shell
        $shortcut = $wshell.CreateShortcut('{path}')
        $shortcut.TargetPath = '{target}'
        $shortcut.WorkingDirectory = '{work_dir}'
        $shortcut.Save()
        """
        subprocess.run(["powershell", "-Command", ps_script], creationflags=subprocess.CREATE_NO_WINDOW)
    else:
        if os.path.exists(path):
            os.remove(path)


def check_autostart():
    startup = os.path.join(os.getenv('APPDATA'), r'Microsoft\Windows\Start Menu\Programs\Startup')
    return os.path.exists(os.path.join(startup, f"{APP_NAME}.lnk"))


def build_and_run_ui():
    root = tk.Tk()
    root.title("引擎配置")
    root.geometry("380x360")  # 加长窗口以容纳新选项
    root.withdraw()

    def on_closing():
        root.withdraw()

    root.protocol("WM_DELETE_WINDOW", on_closing)

    # 提取模式单选框
    tk.Label(root, text="核心提取模式:").pack(pady=(10, 0))
    var_mode = tk.StringVar(value=current_config.get("EXTRACT_MODE", "REGEX"))
    frame_mode = tk.Frame(root)
    frame_mode.pack(pady=5)
    tk.Radiobutton(frame_mode, text="本地正则匹配 (毫秒级/免费)", variable=var_mode, value="REGEX").pack(anchor="w")
    tk.Radiobutton(frame_mode, text="云端大模型提取 (需联网/需配置密钥)", variable=var_mode, value="LLM").pack(
        anchor="w")

    tk.Label(root, text="大模型 API Key:").pack(pady=5)
    entry_key = tk.Entry(root, width=40, show=".")
    entry_key.insert(0, current_config.get("LLM_API_KEY", ""))
    entry_key.pack()

    tk.Label(root, text="大模型接入点 (Model):").pack(pady=5)
    entry_model = tk.Entry(root, width=40)
    entry_model.insert(0, current_config.get("LLM_MODEL", ""))
    entry_model.pack()

    var_autostart = tk.BooleanVar(value=check_autostart())
    chk_autostart = tk.Checkbutton(root, text="开机静默启动", variable=var_autostart)
    chk_autostart.pack(pady=10)

    lbl_status = tk.Label(root, text="")
    lbl_status.pack(pady=2)

    def save_action():
        lbl_status.config(text="正在保存...", fg="blue")
        root.update()

        current_config["EXTRACT_MODE"] = var_mode.get()
        current_config["LLM_API_KEY"] = entry_key.get()
        current_config["LLM_MODEL"] = entry_model.get()
        save_config(current_config)

        try:
            set_autostart(var_autostart.get())
            lbl_status.config(text="配置已成功保存并生效！", fg="green")
        except Exception as e:
            lbl_status.config(text=f"保存失败: {str(e)}", fg="red")

    btn_save = tk.Button(root, text="保存配置", command=save_action, width=15)
    btn_save.pack()

    def process_queue():
        try:
            msg = ui_queue.get_nowait()
            if msg == "show":
                var_mode.set(current_config.get("EXTRACT_MODE", "REGEX"))
                entry_key.delete(0, tk.END)
                entry_key.insert(0, current_config.get("LLM_API_KEY", ""))
                entry_model.delete(0, tk.END)
                entry_model.insert(0, current_config.get("LLM_MODEL", ""))
                var_autostart.set(check_autostart())
                lbl_status.config(text="")

                root.deiconify()
                root.attributes("-topmost", True)
                root.attributes("-topmost", False)
                root.focus_force()
            elif msg == "quit":
                root.destroy()
                return
        except queue.Empty:
            pass
        root.after(200, process_queue)

    process_queue()
    root.mainloop()


def show_config_window(icon, item):
    ui_queue.put("show")


def quit_app(icon, item):
    monitor_instance.stop()
    ui_queue.put("quit")
    icon.stop()


def load_icon_image():
    # 注意这里改用了 get_resource_path，且填入了你实际的图片名
    icon_path = get_resource_path("new-email-verification.png")

    if os.path.exists(icon_path):
        try:
            # 加载并转换为 RGBA 确保透明度正常
            img = Image.open(icon_path).convert("RGBA")
            return img
        except Exception as e:
            print(f"图标加载失败: {e}")

    # 备用绿色方块
    return Image.new('RGB', (64, 64), color=(40, 167, 69))


def on_tray_ready(icon):
    """托盘图标准备就绪后的回调函数，用于触发气泡通知"""
    icon.visible = True
    # 弹出 Windows 右下角原生提示框
    icon.notify("软件已潜伏在系统托盘，随时待命提取验证码。", "运行提示")


if __name__ == '__main__':
    # 1：启动后台工作线程
    t_monitor = threading.Thread(target=run_monitor_thread, daemon=True)
    t_monitor.start()

    # 2：配置托盘与菜单
    menu = pystray.Menu(
        pystray.MenuItem('引擎配置', show_config_window),
        pystray.MenuItem('完全退出', quit_app)
    )
    tray_icon = pystray.Icon("VerifyCodeProLite", load_icon_image(), "验证码无感提取器", menu)

    # 启动托盘线程，并传入 setup 回调以实现气泡通知
    t_tray = threading.Thread(target=tray_icon.run, kwargs={"setup": on_tray_ready}, daemon=True)
    t_tray.start()

    # 3：主线程死守 Tkinter 维持稳定运行
    build_and_run_ui()