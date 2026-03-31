import asyncio
import struct
import threading
import queue
import time
import csv
import os
from datetime import datetime
from typing import Optional, List, Tuple
from bleak import BleakClient, BleakScanner, BleakError
from bleak.backends.device import BLEDevice
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from dataclasses import dataclass

# 蓝牙 UUID 常量
HRS_UUID = "0000180d-0000-1000-8000-00805f9b34fb"  # 心率服务
HRM_UUID = "00002a37-0000-1000-8000-00805f9b34fb"  # 心率测量特征

@dataclass
class HeartRateData:
    """心率数据结构"""
    value: int
    timestamp: float
    formatted_time: str
    # 预留步频字段，未来可以从设备读取并记录
    step_frequency: Optional[float] = None

class HeartRateMonitorGUI:
    """心率监测器图形界面"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("蓝牙心率监测器")
        self.root.geometry("800x600")
        
        # 状态变量
        self.is_scanning = False
        self.is_connected = False
        self.selected_device = None
        self.client = None
        self.devices = []
        
        # 心率日志存储
        self.heart_rate_log = []  # 存储心率数据列表
        self.log_file_path = None  # 当前日志文件路径
        self.is_logging = False  # 是否正在记录日志
        
        # 异步事件循环
        self.loop = asyncio.new_event_loop()
        self.event_thread = threading.Thread(target=self.run_event_loop, daemon=True)
        self.event_thread.start()
        
        # 消息队列
        self.message_queue = queue.Queue()
        
        # 创建界面
        self.create_widgets()
        
        # 定期处理消息队列
        self.process_queue()
        
        # 关闭窗口时的清理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def run_event_loop(self):
        """在单独线程中运行异步事件循环"""
        asyncio.set_event_loop(self.loop)
        self.loop.run_forever()
    
    def run_async(self, coro):
        """在事件循环线程中运行异步函数"""
        future = asyncio.run_coroutine_threadsafe(coro, self.loop)
        return future
    
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="蓝牙心率监测器", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=4, pady=(0, 10))
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_frame = ttk.LabelFrame(main_frame, text="状态", padding="10")
        status_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 10))
        
        status_label = ttk.Label(status_frame, textvariable=self.status_var)
        status_label.grid(row=0, column=0, sticky=tk.W)
        
        # 设备信息
        self.device_info_var = tk.StringVar(value="未连接")
        device_label = ttk.Label(status_frame, textvariable=self.device_info_var)
        device_label.grid(row=0, column=1, padx=(20, 0))
        
        # 日志记录状态
        self.log_status_var = tk.StringVar(value="未记录")
        log_status_label = ttk.Label(status_frame, textvariable=self.log_status_var, foreground="blue")
        log_status_label.grid(row=0, column=2, padx=(20, 0))
        
        # 心率显示
        hr_frame = ttk.LabelFrame(main_frame, text="心率数据", padding="15")
        hr_frame.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 10))
        
        hr_frame.columnconfigure(1, weight=1)
        
        # 心率值
        ttk.Label(hr_frame, text="当前心率:", font=("Arial", 12)).grid(row=0, column=0, sticky=tk.W)
        self.hr_value_var = tk.StringVar(value="--")
        hr_value_label = ttk.Label(hr_frame, textvariable=self.hr_value_var, 
                                   font=("Arial", 24, "bold"), foreground="red")
        hr_value_label.grid(row=0, column=1, sticky=tk.W, padx=(10, 0))
        ttk.Label(hr_frame, text="bpm").grid(row=0, column=2, sticky=tk.W, padx=(5, 0))
        
        # 心率记录统计
        ttk.Label(hr_frame, text="已记录:", font=("Arial", 10)).grid(row=0, column=3, sticky=tk.W, padx=(20, 0))
        self.hr_count_var = tk.StringVar(value="0")
        hr_count_label = ttk.Label(hr_frame, textvariable=self.hr_count_var, 
                                   font=("Arial", 10, "bold"))
        hr_count_label.grid(row=0, column=4, sticky=tk.W, padx=(5, 0))
        ttk.Label(hr_frame, text="次").grid(row=0, column=5, sticky=tk.W, padx=(2, 0))
        
        # 控制按钮区域
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=3, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 扫描按钮
        self.scan_button = ttk.Button(control_frame, text="扫描设备", command=self.start_scan)
        self.scan_button.grid(row=0, column=0, padx=(0, 10))
        
        # 连接/断开按钮
        self.connect_button = ttk.Button(control_frame, text="连接设备", 
                                        command=self.toggle_connection, state=tk.DISABLED)
        self.connect_button.grid(row=0, column=1, padx=(0, 10))
        
        # 开始/停止记录按钮
        self.log_button = ttk.Button(control_frame, text="开始记录", 
                                    command=self.toggle_logging, state=tk.DISABLED)
        self.log_button.grid(row=0, column=2, padx=(0, 10))
        
        # 导出日志按钮
        self.export_button = ttk.Button(control_frame, text="导出日志", 
                                       command=self.export_log, state=tk.DISABLED)
        self.export_button.grid(row=0, column=3, padx=(0, 10))
        
        # 清空日志按钮
        ttk.Button(control_frame, text="清空显示", command=self.clear_display_log).grid(row=0, column=4, padx=(0, 10))
        
        # 设备列表
        list_frame = ttk.LabelFrame(main_frame, text="可用设备", padding="10")
        list_frame.grid(row=4, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # 设备列表树状视图
        columns = ("名称", "地址", "信号强度")
        self.device_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=6)
        
        for col in columns:
            self.device_tree.heading(col, text=col)
            self.device_tree.column(col, width=150)
        
        self.device_tree.column("地址", width=200)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.device_tree.yview)
        self.device_tree.configure(yscrollcommand=scrollbar.set)
        
        self.device_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 绑定选择事件
        self.device_tree.bind("<<TreeviewSelect>>", self.on_device_select)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="操作日志", padding="10")
        log_frame.grid(row=5, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, width=80, height=10)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置主框架行权重
        main_frame.rowconfigure(4, weight=1)
        main_frame.rowconfigure(5, weight=1)
    
    def log_message(self, message: str):
        """添加日志消息"""
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        log_entry = f"[{timestamp}] {message}\n"
        
        # 在GUI线程中更新日志
        self.root.after(0, self._update_log, log_entry)
    
    def _update_log(self, message: str):
        """更新日志显示"""
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)
    
    def clear_display_log(self):
        """清空显示日志（仅清空显示，不清除心率记录）"""
        self.log_text.delete(1.0, tk.END)
        self.log_message("显示日志已清空")
    
    def update_status(self, status: str):
        """更新状态栏"""
        self.status_var.set(status)
    
    def update_device_info(self, info: str):
        """更新设备信息"""
        self.device_info_var.set(info)
    
    def update_log_status(self, status: str, color="blue"):
        """更新日志记录状态"""
        self.log_status_var.set(status)
        # 这里可以更新颜色，但需要获取标签引用
    
    def update_heart_rate(self, hr_data: HeartRateData):
        """更新心率显示"""
        self.hr_value_var.set(str(hr_data.value))
        
        # 如果正在记录日志，则保存心率数据
        if self.is_logging:
            self.heart_rate_log.append(hr_data)
            self.hr_count_var.set(str(len(self.heart_rate_log)))
    
    def toggle_logging(self):
        """切换日志记录状态"""
        if not self.is_logging:
            self.start_logging()
        else:
            self.stop_logging()
    
    def start_logging(self):
        """开始记录心率日志"""
        self.is_logging = True
        self.heart_rate_log = []  # 清空之前的记录
        self.hr_count_var.set("0")
        self.log_button.config(text="停止记录")
        self.export_button.config(state=tk.NORMAL)
        self.update_log_status("记录中...", "green")
        self.log_message("开始记录心率日志")
    
    def stop_logging(self):
        """停止记录心率日志"""
        self.is_logging = False
        self.log_button.config(text="开始记录")
        self.update_log_status(f"已记录{len(self.heart_rate_log)}次", "blue")
        self.log_message(f"停止记录心率日志，共记录 {len(self.heart_rate_log)} 次心率")
    
    def export_log(self):
        """导出心率日志到文件"""
        if not self.heart_rate_log:
            messagebox.showwarning("警告", "没有可导出的心率数据")
            return
        
        # 弹出文件保存对话框
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[
                ("CSV文件", "*.csv"),
                ("文本文件", "*.txt"),
                ("所有文件", "*.*")
            ],
            title="保存心率日志",
            initialfile=f"heart_rate_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
        
        if not file_path:
            return  # 用户取消了保存
        
        try:
            # 写入文件（不使用CSV模块，直接写入指定格式）
            with open(file_path, 'w', encoding='utf-8') as f:
                for hr_data in self.heart_rate_log:
                    # 格式：时间戳,日期时间,心率值bpm,步频
                    step_freq = hr_data.step_frequency if hr_data.step_frequency is not None else ''
                    line = f"{hr_data.timestamp},{hr_data.formatted_time},{hr_data.value}bpm,{step_freq}\n"
                    f.write(line)
            
            self.log_file_path = file_path
            self.log_message(f"心率日志已导出到: {file_path}")
            messagebox.showinfo("导出成功", f"心率日志已成功导出到:\n{file_path}\n\n共导出 {len(self.heart_rate_log)} 条记录")
            
        except Exception as e:
            self.log_message(f"导出日志失败: {e}")
            messagebox.showerror("导出失败", f"导出日志时出错:\n{str(e)}")
    
    def export_log_as_json(self):
        """将心率日志导出为JSON格式（与Android APP兼容）"""
        if not self.heart_rate_log:
            messagebox.showwarning("警告", "没有可导出的心率数据")
            return
        
        # 弹出文件保存对话框
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[
                ("JSON文件", "*.json"),
                ("所有文件", "*.*")
            ],
            title="保存心率日志(JSON)",
            initialfile=f"heart_rate_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
        
        if not file_path:
            return  # 用户取消了保存
        
        try:
            import json
            
            # 构建JSON数据结构
            log_data = {
                "device_info": {
                    "name": self.selected_device.name if self.selected_device else "未知设备",
                    "address": self.selected_device.address if self.selected_device else "未知"
                },
                "record_count": len(self.heart_rate_log),
                "start_time": self.heart_rate_log[0].formatted_time if self.heart_rate_log else "",
                "end_time": self.heart_rate_log[-1].formatted_time if self.heart_rate_log else "",
                "heart_rate_data": [
                    {
                        "timestamp": hr_data.timestamp,
                        "datetime": hr_data.formatted_time,
                        "heart_rate": hr_data.value,
                        # 兼容步频字段，可能为 None
                        "step_frequency": hr_data.step_frequency
                    }
                    for hr_data in self.heart_rate_log
                ]
            }
            
            # 写入JSON文件
            with open(file_path, 'w', encoding='utf-8') as jsonfile:
                json.dump(log_data, jsonfile, ensure_ascii=False, indent=2)
            
            self.log_message(f"心率日志(JSON)已导出到: {file_path}")
            messagebox.showinfo("导出成功", f"心率日志已成功导出为JSON格式:\n{file_path}")
            
        except Exception as e:
            self.log_message(f"导出JSON日志失败: {e}")
            messagebox.showerror("导出失败", f"导出JSON日志时出错:\n{str(e)}")
    
    def start_scan(self):
        """开始扫描设备"""
        if self.is_scanning:
            return
            
        self.is_scanning = True
        self.scan_button.config(state=tk.DISABLED)
        self.update_status("正在扫描设备...")
        self.log_message("开始扫描蓝牙设备")
        
        # 清空设备列表
        for item in self.device_tree.get_children():
            self.device_tree.delete(item)
        
        # 在异步线程中开始扫描
        self.run_async(self.scan_devices())
    
    async def scan_devices(self):
        """异步扫描设备"""
        try:
            # 扫描设备
            self.log_message("扫描中...")
            devices = await BleakScanner.discover(
                timeout=10.0,
                service_uuids=[HRS_UUID]
            )
            
            # 更新设备列表
            self.devices = devices
            self.root.after(0, self.update_device_list, devices)
            
            self.log_message(f"扫描完成，找到 {len(devices)} 个设备")
            
        except Exception as e:
            self.log_message(f"扫描错误: {e}")
        finally:
            self.is_scanning = False
            self.root.after(0, self.scan_completed)
    
    def scan_completed(self):
        """扫描完成后的处理"""
        self.scan_button.config(state=tk.NORMAL)
        self.update_status("扫描完成")
    
    def update_device_list(self, devices: List[BLEDevice]):
        """更新设备列表显示"""
        for device in devices:
            name = device.name or "未知设备"
            address = device.address
            rssi = getattr(device, 'rssi', 'N/A')
            
            self.device_tree.insert("", tk.END, values=(name, address, rssi))
    
    def on_device_select(self, event):
        """处理设备选择事件"""
        selection = self.device_tree.selection()
        if selection:
            item = selection[0]
            values = self.device_tree.item(item, "values")
            
            if values and len(values) >= 2:
                device_name = values[0]
                device_address = values[1]
                
                # 查找对应的设备对象
                self.selected_device = None
                for device in self.devices:
                    if device.address == device_address:
                        self.selected_device = device
                        break
                
                if self.selected_device:
                    self.update_status(f"已选择设备: {device_name}")
                    self.connect_button.config(state=tk.NORMAL)
                    self.log_message(f"选择设备: {device_name} ({device_address})")
    
    def toggle_connection(self):
        """切换连接状态"""
        if not self.is_connected:
            self.connect_device()
        else:
            self.disconnect_device()
    
    def connect_device(self):
        """连接设备"""
        if not self.selected_device:
            messagebox.showwarning("警告", "请先选择设备")
            return
        
        self.update_status("正在连接设备...")
        self.log_message(f"正在连接设备: {self.selected_device.address}")
        
        # 禁用按钮
        self.scan_button.config(state=tk.DISABLED)
        self.connect_button.config(state=tk.DISABLED)
        
        # 在异步线程中连接设备
        self.run_async(self.connect_and_monitor())
    
    async def connect_and_monitor(self):
        """异步连接和监控设备"""
        try:
            device_name = self.selected_device.name or self.selected_device.address
            async with BleakClient(self.selected_device) as client:
                self.client = client
                self.is_connected = True
                # 更新UI状态
                self.root.after(0, self.connection_established, device_name)
                # 寻找心率服务
                services = client.services
                hrs_service = None
                
                for service in services:
                    if service.uuid.lower() == HRS_UUID.lower():
                        hrs_service = service
                        break
                if hrs_service:
                    # 寻找心率测量特征
                    hrm_char = None
                    for char in hrs_service.characteristics:
                        if char.uuid.lower() == HRM_UUID.lower():
                            hrm_char = char
                            break
                    if hrm_char:
                        # 订阅心率测量通知
                        await client.start_notify(hrm_char.handle, self.heart_rate_callback)
                        
                        # 保持连接
                        while client.is_connected:
                            await asyncio.sleep(1)
                    else:
                        self.log_message("未找到心率测量特征")
                else:
                    self.log_message("未找到心率服务")
                
        except Exception as e:
            self.log_message(f"连接错误: {e}")
        finally:
            self.is_connected = False
            self.client = None
            self.root.after(0, self.connection_lost)
    
    def connection_established(self, device_name: str):
        """连接建立后的UI更新"""
        self.update_status(f"已连接: ")
        self.update_device_info(device_name)
        self.connect_button.config(text="断开连接", state=tk.NORMAL)
        self.log_button.config(state=tk.NORMAL)  # 启用记录按钮
        self.log_message("设备连接成功，开始监听心率数据")
    
    def connection_lost(self):
        """连接断开后的UI更新"""
        self.update_status("连接已断开")
        self.update_device_info("未连接")
        self.connect_button.config(text="连接设备", state=tk.NORMAL)
        self.scan_button.config(state=tk.NORMAL)
        self.log_button.config(state=tk.DISABLED)  # 禁用记录按钮
        
        # 停止记录日志（如果正在记录）
        if self.is_logging:
            self.stop_logging()
        
        # 重置心率显示
        self.hr_value_var.set("--")
        
        self.log_message("设备连接已断开")
    
    def disconnect_device(self):
        """断开设备连接"""
        if self.client and self.is_connected:
            self.log_message("正在断开设备连接...")
            self.run_async(self.client.disconnect())
    
    def heart_rate_callback(self, sender: int, data: bytearray):
        """心率数据回调函数"""
        try:
            if len(data) < 2:
                return
                
            flag = data[0]
            
            # 解析心率值
            heart_rate_value = data[1]
            if flag & 0x01:  # 16位心率值
                if len(data) >= 3:
                    heart_rate_value = struct.unpack('<H', data[1:3])[0]
                else:
                    return
            
            # 获取当前时间
            current_time = time.time()
            formatted_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # 创建心率数据对象
            hr_data = HeartRateData(
                value=heart_rate_value,
                timestamp=current_time,
                formatted_time=formatted_time
            )
            
            # 在GUI线程中更新显示
            self.root.after(0, self.update_heart_rate, hr_data)
            
            # 记录到日志（如果正在记录）
            if self.is_logging:
                self.root.after(0, self.log_message, f"记录心率: {heart_rate_value} bpm")
            else:
                self.root.after(0, self.log_message, f"心率: {heart_rate_value} bpm")
            
        except Exception as e:
            self.root.after(0, self.log_message, f"解析心率数据时出错: {e}")
    
    def process_queue(self):
        """处理消息队列"""
        try:
            while True:
                message = self.message_queue.get_nowait()
                self.log_message(message)
        except queue.Empty:
            pass
        
        # 每100ms检查一次队列
        self.root.after(100, self.process_queue)
    
    def on_closing(self):
        """关闭窗口时的清理"""
        # 停止记录日志（如果正在记录）
        if self.is_logging:
            self.stop_logging()
        
        # 提示保存未导出的日志
        if self.heart_rate_log and not self.log_file_path:
            response = messagebox.askyesno(
                "保存日志", 
                f"您有 {len(self.heart_rate_log)} 条未保存的心率数据，是否导出保存？"
            )
            if response:
                self.export_log()
        
        if self.is_connected and self.client:
            # 尝试断开连接
            try:
                future = self.run_async(self.client.disconnect())
                future.result(timeout=2)
            except:
                pass
        
        # 停止事件循环
        self.loop.call_soon_threadsafe(self.loop.stop)
        
        # 关闭窗口
        self.root.destroy()

def main():
    """主函数"""
    root = tk.Tk()
    app = HeartRateMonitorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()