import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import psutil
import ctypes
import socket
from threading import Thread, Lock
import time
import sys
import logging


class NetworkBondingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("专业网络叠加工具 - 跃点数负载均衡（实时监控）")

        # 初始化日志记录
        logging.basicConfig(
            filename='network_bonding.log',
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

        self.speed_data = {}
        self.speed_lock = Lock()
        self.original_metrics = {}

        if not self.is_admin():
            messagebox.showwarning("权限警告", "需要管理员权限修改网络设置！\n请右键以管理员身份运行本程序")
            self.root.destroy()

        self.create_widgets()
        self.refresh_adapters()
        self.selected_adapters = []

        # 初始化速度监控
        self.update_speed_thread = Thread(target=self.update_speeds, daemon=True)
        self.update_speed_thread.start()
        self.speed_update_interval = 1
        self.setup_speed_refresh()

    def is_admin(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False
    def create_widgets(self):

        adapter_frame = ttk.LabelFrame(self.root, text="网络适配器选择（实时速度）")
        adapter_frame.pack(padx=10, pady=5, fill="x")

        # 使用Treeview替代Listbox以支持多列
        columns = ("adapter", "sent", "recv")
        self.adapter_tree = ttk.Treeview(adapter_frame, columns=columns, show="headings")

        # 设置列标题
        self.adapter_tree.heading("adapter", text="适配器名称")
        self.adapter_tree.heading("sent", text="发送速度 (KB/s)")
        self.adapter_tree.heading("recv", text="接收速度 (KB/s)")

        # 设置列宽
        self.adapter_tree.column("adapter", width=200)
        self.adapter_tree.column("sent", width=120, anchor="e")
        self.adapter_tree.column("recv", width=120, anchor="e")

        self.adapter_tree.pack(padx=10, pady=5, fill="both", expand=True)

        # 控制框架
        control_frame = ttk.Frame(self.root)
        control_frame.pack(padx=10, pady=5)

        # 模式选择
        mode_label = ttk.Label(control_frame, text="模式:")
        mode_label.pack(side=tk.LEFT, padx=5)

        self.mode_var = tk.StringVar()
        self.mode_combobox = ttk.Combobox(control_frame, textvariable=self.mode_var,
                                          values=["合并网速", "单独使用"], state="readonly", width=12)
        self.mode_combobox.pack(side=tk.LEFT, padx=5)
        self.mode_combobox.current(0)

        # 控制按钮
        self.start_btn = ttk.Button(control_frame, text="启用叠加", command=self.start_bonding)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = ttk.Button(control_frame, text="停止叠加", command=self.stop_bonding, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        refresh_btn = ttk.Button(control_frame, text="刷新适配器", command=self.refresh_adapters)
        refresh_btn.pack(side=tk.LEFT, padx=5)

        # 状态显示
        self.status_label = ttk.Label(self.root, text="状态: 未启用")
        self.status_label.pack(padx=10, pady=5)


    def refresh_adapters(self):
        """刷新适配器列表并初始化监控"""
        adapters = []
        try:
            for name, stats in psutil.net_if_stats().items():
                if stats.isup and not name.lower().startswith(('virtual', 'loopback')):
                    adapters.append(name)
                    io = psutil.net_io_counters(pernic=True).get(name, None)
                    with self.speed_lock:
                        self.speed_data[name] = {
                            'sent': io.bytes_sent if io else 0,
                            'recv': io.bytes_recv if io else 0
                        }
        except Exception as e:
            logging.error(f"刷新适配器失败: {str(e)}")
            messagebox.showerror("错误", f"无法获取网络适配器信息: {str(e)}")

        self.adapter_tree.delete(*self.adapter_tree.get_children())
        for adapter in adapters:
            with self.speed_lock:
                sent = self.speed_data[adapter]['sent']
                recv = self.speed_data[adapter]['recv']
            self.adapter_tree.insert("", tk.END, values=(
                adapter,
                f"{sent / 1024:.2f}" if sent else "0.00",
                f"{recv / 1024:.2f}" if recv else "0.00"
            ))

    def setup_speed_refresh(self):
        self.root.after(self.speed_update_interval * 1000, self.refresh_speeds)

    def refresh_speeds(self):
        # ... [保持原有速度刷新逻辑] ...
        self.setup_speed_refresh()

    def update_speeds(self):
        """后台速度更新线程"""
        while True:
            adapters = [self.adapter_tree.item(item, "values")[0]
                        for item in self.adapter_tree.get_children()]

            for adapter in adapters:
                try:
                    io = psutil.net_io_counters(pernic=True).get(adapter, None)
                    if io:
                        current_sent = io.bytes_sent
                        current_recv = io.bytes_recv

                        with self.speed_lock:
                            prev_sent = self.speed_data[adapter]['sent']
                            prev_recv = self.speed_data[adapter]['recv']

                        sent_speed = (current_sent - prev_sent) / 1024
                        recv_speed = (current_recv - prev_recv) / 1024

                        with self.speed_lock:
                            self.speed_data[adapter]['sent'] = current_sent
                            self.speed_data[adapter]['recv'] = current_recv

                        self.root.after(0, self.update_gui_speeds, adapter, sent_speed, recv_speed)
                except Exception as e:
                    logging.error(f"速度更新失败: {str(e)}")

            time.sleep(self.speed_update_interval)

    def update_gui_speeds(self, adapter, sent, recv):
        """线程安全的GUI更新"""
        for item in self.adapter_tree.get_children():
            if self.adapter_tree.item(item, "values")[0] == adapter:
                self.adapter_tree.item(item, values=(
                    adapter,
                    f"{sent:.2f}" if sent >= 0 else "0.00",
                    f"{recv:.2f}" if recv >= 0 else "0.00"
                ))
                break

    def start_bonding(self):
        self.selected_adapters = self.adapter_tree.selection()
        if not self.selected_adapters:
            messagebox.showerror("选择错误", "请至少选择一个网络适配器")
            return

        adapters = [self.adapter_tree.item(item, "values")[0]
                    for item in self.selected_adapters]
        mode = self.mode_var.get()

        try:
            if mode == "合并网速":
                self.configure_load_balancing(adapters)
            elif mode == "单独使用":
                self.configure_single_adapter(adapters)

            self.status_label.config(text=f"状态: 已启用（{mode}）")
            self.start_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)

        except Exception as e:
            messagebox.showerror("配置失败", f"错误代码: {str(e)}")
    def configure_load_balancing(self, adapters):
        """核心负载均衡配置函数"""
        all_adapters = [self.adapter_tree.item(item, "values")[0]
                        for item in self.adapter_tree.get_children()]

        # 备份原始跃点数配置
        for adapter in all_adapters:
            self.original_metrics[adapter] = self.get_interface_metric(adapter)

        try:
            # 1. 启用TCP自动调优
            self.run_netsh_command(
                'netsh int tcp set global autotuninglevel=normal',
                "启用TCP自动调优失败"
            )

            # 2. 为所有适配器应用互联网模板
            for adapter in all_adapters:
                self.run_netsh_command(
                    f'netsh int tcp set supplemental template=internet interface="{adapter}"',
                    f"配置适配器 {adapter} 的TCP模板失败"
                )

            # 3. 统一设置低跃点数
            for adapter in all_adapters:
                metric = 1 if adapter in adapters else 1000
                self.run_netsh_command(
                    f'netsh interface ipv4 set interface interface="{adapter}" metric={metric}',
                    f"设置适配器 {adapter} 跃点数失败"
                )

            messagebox.showinfo("成功", "网络负载均衡配置已生效！")

        except Exception as e:
            self.restore_original_config()
            raise e

    def run_netsh_command(self, command, error_msg):
        try:
            logging.debug(f"执行命令: {command}")
            result = subprocess.run(
                command,
                shell=True,
                check=True,
                capture_output=True,
                text=True,
                timeout=5
            )
            if result.returncode != 0:
                raise subprocess.CalledProcessError(
                    result.returncode,
                    command,
                    output=result.stdout,
                    stderr=result.stderr
                )
            logging.debug(f"命令输出: {result.stdout}")
            return result.stdout
        except subprocess.TimeoutExpired:
            logging.error(f"命令执行超时: {command}")
            messagebox.showerror("超时", "网络命令执行超时，请检查网络适配器状态")
            raise
        except subprocess.CalledProcessError as e:
            logging.error(f"命令执行失败: {command}\n错误输出: {e.stderr}")
            messagebox.showerror("命令执行失败",
                              f"{error_msg}\n\n错误代码: {e.returncode}\n{e.stderr}")
            raise

    def configure_single_adapter(self, adapters):
        """增强版单适配器配置"""
        if not adapters:
            return

        selected = adapters[0]
        all_adapters = [self.adapter_tree.item(item, "values")[0]
                        for item in self.adapter_tree.get_children()]

        try:
            # 备份原始配置
            for adapter in all_adapters:
                self.original_metrics[adapter] = self.get_interface_metric(adapter)

            # 设置选中适配器跃点数为1，其他为1000
            for adapter in all_adapters:
                metric = 1 if adapter == selected else 1000
                self.run_netsh_command(
                    f'netsh interface ipv4 set interface interface="{adapter}" metric={metric}',
                    f"设置适配器 {adapter} 跃点数失败"
                )

            # 设置默认网关路由
            gateway = self.get_default_gateway(selected)
            if gateway:
                idx = self.get_interface_index(selected)
                if idx:
                    self.run_netsh_command(
                        f'route add 0.0.0.0 mask 0.0.0.0 {gateway} if {idx} metric 1',
                        f"设置默认路由失败"
                    )

        except Exception as e:
            self.restore_original_config()
            raise e

    def stop_bonding(self):
        self.restore_original_config()
        self.status_label.config(text="状态: 未启用")
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)

    def restore_original_config(self):
        """增强版配置恢复"""
        try:
            for adapter, metric in self.original_metrics.items():
                self.run_netsh_command(
                    f'netsh interface ipv4 set interface interface="{adapter}" metric={metric}',
                    f"恢复适配器 {adapter} 跃点数失败"
                )
                self.run_netsh_command(
                    f'netsh int tcp set supplemental template=disabled interface="{adapter}"',
                    f"恢复TCP模板失败"
                )

            self.run_netsh_command(
                'netsh int tcp set global autotuninglevel=disabled',
                "恢复TCP自动调优失败"
            )
            messagebox.showinfo("恢复成功", "网络设置已恢复初始状态")

        except Exception as e:
            messagebox.showerror("恢复失败", f"错误详情: {str(e)}")

    # 增强版辅助函数 --------------------------------------------------
    def get_interface_metric(self, adapter):
        """增强版跃点数获取"""
        try:
            output = self.run_netsh_command(
                f'netsh interface ipv4 show interface "{adapter}"',
                "获取接口信息失败"
            )
            for line in output.split('\n'):
                if "Metric" in line:
                    return int(line.split(':')[-1].strip())
            return None
        except:
            return None

    def get_default_gateway(self, adapter):
        """增强版默认网关获取"""
        try:
            addrs = psutil.net_if_addrs()[adapter]
            for addr in addrs:
                if addr.family == socket.AF_INET and addr.netmask != '0.0.0.0':
                    return addr.gateway
            return None
        except Exception as e:
            logging.error(f"获取网关失败: {str(e)}")
            return None

    def get_interface_index(self, adapter):
        """增强版接口索引获取（支持多语言系统）"""
        try:
            output = self.run_netsh_command(
                'netsh interface show interface',
                "获取接口列表失败"
            )

            # 支持中英文系统
            search_terms = ["已启用", "Enabled"] if "中文" in sys.getwindowsversion().language else ["Enabled"]

            for line in output.split('\n'):
                line = line.strip()
                if adapter in line and any(term in line for term in search_terms):
                    parts = line.split()
                    if len(parts) >= 2 and parts[1] in ["启用", "Enabled"]:
                        return parts[0]  # 索引号在第一列
            return None
        except Exception as e:
            logging.error(f"获取接口索引失败: {str(e)}")
            return None

if __name__ == "__main__":
    root = tk.Tk()
    app = NetworkBondingApp(root)
    root.mainloop()