import tkinter as tk
from tkinter import ttk, messagebox
import cv2
from PIL import Image, ImageTk
import threading
import json
import os
import time  # 引入时间库用于延时
from pygrabber.dshow_graph import FilterGraph

# 配置文件名称
CONFIG_FILE = "cam_config.json"

class ConfigManager:
    @staticmethod
    def load_config():
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                return {}
        return {}

    @staticmethod
    def save_config(device_name, options):
        data = ConfigManager.load_config()
        data[device_name] = options
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"保存配置失败: {e}")

class CameraConfigPane:
    def __init__(self, parent, index, app_instance):
        self.index = index
        self.app = app_instance
        
        self.frame = tk.LabelFrame(parent, text=f"通道 {index + 1}", padx=5, pady=5)
        self.frame.grid(row=0, column=index, padx=5, sticky="nsew")

        self.var_enable = tk.BooleanVar(value=True)
        self.chk_enable = tk.Checkbutton(self.frame, text="启用此摄像头", variable=self.var_enable, fg="blue")
        self.chk_enable.pack(anchor="w")

        tk.Label(self.frame, text="选择设备:").pack(anchor="w")
        self.var_device = tk.StringVar()
        self.combo_device = ttk.Combobox(self.frame, textvariable=self.var_device, state="readonly", width=20)
        self.combo_device.pack(fill="x", pady=(0, 5))
        self.combo_device.bind("<<ComboboxSelected>>", self.on_device_selected)

        tk.Label(self.frame, text="分辨率 & 格式:").pack(anchor="w")
        self.var_res = tk.StringVar()
        self.combo_res = ttk.Combobox(self.frame, textvariable=self.var_res, state="readonly", width=20)
        self.combo_res.pack(fill="x", pady=(0, 5))
        
        self.lbl_status = tk.Label(self.frame, text="等待配置", fg="gray", font=("Arial", 8))
        self.lbl_status.pack(anchor="w")

    def update_device_list(self, devices_dict):
        device_names = list(devices_dict.keys())
        self.combo_device['values'] = device_names
        if device_names and self.index < len(device_names):
            if not self.combo_device.get():
                self.combo_device.current(self.index)
                self.on_device_selected(None)

    def on_device_selected(self, event):
        full_dev_name = self.var_device.get()
        if not full_dev_name:
            return
            
        clean_name = full_dev_name.split(": ", 1)[-1] if ": " in full_dev_name else full_dev_name
        cached_data = ConfigManager.load_config()
        
        if clean_name in cached_data and len(cached_data[clean_name]) > 0:
            self.combo_res['values'] = cached_data[clean_name]
            self.combo_res.current(0)
            self.lbl_status.config(text="已加载配置 (缓存)", fg="green")
        else:
            dev_id = self.app.devices_dict.get(full_dev_name)
            if dev_id is not None:
                self.lbl_status.config(text="正在重新扫描硬件...", fg="orange")
                self.combo_res.set("扫描中...")
                self.combo_res['values'] = []
                threading.Thread(target=self.scan_resolutions, args=(dev_id, clean_name), daemon=True).start()

    def scan_resolutions(self, dev_id, clean_name):
        """【修复重点】带重试机制的深度扫描"""
        scan_list = [(1920, 1080), (1280, 720), (800, 600)]
        formats = [
            ("MJPG", cv2.VideoWriter_fourcc(*'MJPG')),
            ("YUY2", cv2.VideoWriter_fourcc(*'YUY2')),
        ]

        available_options = []
        cap = cv2.VideoCapture() # 先初始化对象
        
        # --- 重试打开逻辑 ---
        is_opened = False
        for attempt in range(1, 4): # 最多尝试3次
            # 这里的延时非常重要，给驱动喘息时间
            time.sleep(0.5) 
            try:
                cap.open(dev_id, cv2.CAP_DSHOW)
                if cap.isOpened():
                    is_opened = True
                    break
                else:
                    print(f"[{clean_name}] 打开失败 (尝试 {attempt}/3)...")
            except Exception as e:
                print(f"[{clean_name}] 异常: {e}")
        
        if is_opened:
            # 获取默认作为保底
            try:
                def_w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
                def_h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
                if def_w > 0 and def_h > 0:
                    available_options.append(f"默认 {def_w}x{def_h} (Auto)")
            except: pass

            # 开始循环测试
            for fmt_name, fourcc in formats:
                for w, h in scan_list:
                    try:
                        cap.set(cv2.CAP_PROP_FOURCC, fourcc)
                        cap.set(cv2.CAP_PROP_FRAME_WIDTH, w)
                        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, h)
                        
                        # 必须读取，且给一点点缓冲
                        cap.read() 
                        
                        act_w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
                        act_h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
                        
                        if act_w == w and act_h == h:
                            option_str = f"{fmt_name} {w}x{h}"
                            if option_str not in available_options:
                                available_options.append(option_str)
                    except: pass
            
            cap.release()
        else:
            print(f"[{clean_name}] 最终打开失败，驱动忙碌。")

        def finish_scan():
            if available_options:
                unique_options = list(dict.fromkeys(available_options))
                self.combo_res['values'] = unique_options
                self.combo_res.current(0)
                self.lbl_status.config(text="扫描完成", fg="green")
                ConfigManager.save_config(clean_name, unique_options)
            else:
                self.combo_res['values'] = ["获取失败(请重试)"]
                self.combo_res.current(0)
                self.lbl_status.config(text="驱动忙碌/无信号", fg="red")

        self.app.root.after(0, finish_scan)

    def get_config(self):
        if not self.var_enable.get():
            return None
        full_dev_name = self.var_device.get()
        if not full_dev_name:
            return None
        dev_id = self.app.devices_dict.get(full_dev_name)
        res_str = self.var_res.get()
        config = {"id": dev_id, "fourcc": None, "width": 640, "height": 480}
        try:
            if "默认" in res_str:
                parts = res_str.split(' ')
                dims = parts[1].split('x')
                config['width'] = int(dims[0])
                config['height'] = int(dims[1])
            else:
                parts = res_str.split(' ')
                if len(parts) >= 2:
                    fmt = parts[0]
                    dims = parts[1].split('x')
                    config['width'] = int(dims[0])
                    config['height'] = int(dims[1])
                    if fmt == "MJPG":
                        config['fourcc'] = cv2.VideoWriter_fourcc(*'MJPG')
                    elif fmt == "YUY2":
                        config['fourcc'] = cv2.VideoWriter_fourcc(*'YUY2')
        except: pass
        return config


class MultiCamApp:
    def __init__(self, root):
        self.root = root
        self.root.title("多路监控系统 - 稳定增强版")
        self.root.geometry("1200x700")
        
        self.is_running = False
        self.caps = [None] * 4
        self.devices_dict = {} 

        self._init_gui()
        # 延时一点启动，避免 pygrabber 和 UI 抢占资源
        self.root.after(800, self.refresh_devices)

    def _init_gui(self):
        top_frame = tk.Frame(self.root, pady=10)
        top_frame.pack(side=tk.TOP, fill=tk.X)

        self.configs = []
        for i in range(4):
            cp = CameraConfigPane(top_frame, i, self)
            self.configs.append(cp)

        btn_frame = tk.Frame(top_frame)
        btn_frame.grid(row=0, column=4, padx=20, sticky="nsew")

        self.btn_toggle = tk.Button(btn_frame, text="▶ 启动选中设备", font=("Arial", 12, "bold"), 
                                    bg="#2E7D32", fg="white", width=16, height=2,
                                    command=self.toggle_cameras)
        self.btn_toggle.pack(pady=5)

        btn_refresh = tk.Button(btn_frame, text="⟳ 强制重新扫描", command=self.force_rescan)
        btn_refresh.pack(fill="x")
        
        self.video_frame = tk.Frame(self.root, bg="black")
        self.video_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        self.video_frame.grid_rowconfigure(0, weight=1)
        self.video_frame.grid_rowconfigure(1, weight=1)
        self.video_frame.grid_columnconfigure(0, weight=1)
        self.video_frame.grid_columnconfigure(1, weight=1)

        self.video_labels = []
        for i in range(4):
            lbl = tk.Label(self.video_frame, bg="black", text=f"通道 {i+1} 待机", fg="#666", font=("Arial", 16))
            lbl.grid(row=i//2, column=i%2, sticky="nsew", padx=2, pady=2)
            self.video_labels.append(lbl)

    def force_rescan(self):
        if self.is_running:
            self.stop_cameras()
            # 给一点时间让摄像头完全释放
            self.root.update()
            time.sleep(0.5) 

        if os.path.exists(CONFIG_FILE):
            try:
                os.remove(CONFIG_FILE)
            except: pass
        
        self.refresh_devices()
        for cfg in self.configs:
            cfg.on_device_selected(None)

    def refresh_devices(self):
        try:
            # 获取设备前，强制垃圾回收一下，或者等一下
            time.sleep(0.2)
            graph = FilterGraph()
            devices = graph.get_input_devices()
            self.devices_dict = {}
            for i, name in enumerate(devices):
                display_name = f"{i}: {name}"
                self.devices_dict[display_name] = i
            
            for cfg in self.configs:
                cfg.update_device_list(self.devices_dict)
        except Exception as e:
            print(f"刷新设备列表出错: {e}")

    def toggle_cameras(self):
        if self.is_running:
            self.stop_cameras()
        else:
            # 启动前加一个短延时，防止连续点击
            self.btn_toggle.config(state="disabled")
            self.root.after(200, self.start_cameras)

    def start_cameras(self):
        self.is_running = True
        self.btn_toggle.config(text="⏹ 停止所有", bg="#C62828", state="normal")
        
        active_count = 0
        for i, cfg in enumerate(self.configs):
            settings = cfg.get_config()
            if settings:
                # 开启线程去启动摄像头，避免 UI 卡死
                # 注意：这里为了简化逻辑，依然在主线程循环启动，但加入 Retry
                try:
                    cap = cv2.VideoCapture()
                    # --- 启动时的重试逻辑 ---
                    opened = False
                    for attempt in range(3):
                        cap.open(settings['id'], cv2.CAP_DSHOW)
                        if cap.isOpened():
                            opened = True
                            break
                        time.sleep(0.3) # 失败重试延时
                    
                    if opened:
                        if settings['fourcc']:
                            cap.set(cv2.CAP_PROP_FOURCC, settings['fourcc'])
                        cap.set(cv2.CAP_PROP_FRAME_WIDTH, settings['width'])
                        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, settings['height'])
                        cap.read() 

                        self.caps[i] = cap
                        active_count += 1
                    else:
                        self.video_labels[i].config(text="占用/打开失败", fg="red")
                except Exception as e:
                    print(f"Cam {i} error: {e}")
            else:
                self.video_labels[i].config(text="已禁用", fg="#444", image='')

        if active_count > 0:
            self.update_loop()
        else:
            if self.is_running: # 如果原本想运行但一个都没打开
                self.stop_cameras()
                messagebox.showwarning("提示", "未能成功打开任何摄像头。\n请检查是否被其他程序占用。")

    def stop_cameras(self):
        self.is_running = False
        self.btn_toggle.config(text="▶ 启动选中设备", bg="#2E7D32")
        for i in range(4):
            if self.caps[i]:
                # 释放资源
                self.caps[i].release()
                self.caps[i] = None
            self.video_labels[i].config(image='', text=f"通道 {i+1} 待机")

    def update_loop(self):
        if not self.is_running:
            return

        for i in range(4):
            cap = self.caps[i]
            if cap and cap.isOpened():
                try:
                    ret, frame = cap.read()
                    if ret:
                        label_w = self.video_labels[i].winfo_width()
                        label_h = self.video_labels[i].winfo_height()
                        
                        if label_w > 10 and label_h > 10:
                            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                            img = Image.fromarray(frame)
                            
                            img_w, img_h = img.size
                            ratio = min(label_w / img_w, label_h / img_h)
                            new_w = int(img_w * ratio)
                            new_h = int(img_h * ratio)
                            
                            img = img.resize((new_w, new_h), Image.Resampling.BILINEAR)
                            
                            final_img = Image.new('RGB', (label_w, label_h), (0, 0, 0))
                            pos_x = (label_w - new_w) // 2
                            pos_y = (label_h - new_h) // 2
                            final_img.paste(img, (pos_x, pos_y))
                            
                            imgtk = ImageTk.PhotoImage(image=final_img)
                            self.video_labels[i].imgtk = imgtk
                            self.video_labels[i].config(image=imgtk, text='')
                    else:
                        # 偶尔读不到帧不代表断开，只是一帧丢失，不要立刻报错
                        pass 
                except:
                    pass

        self.root.after(30, self.update_loop)

    def on_close(self):
        self.stop_cameras()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = MultiCamApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()