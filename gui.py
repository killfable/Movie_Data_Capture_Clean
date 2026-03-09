import configparser
import os
import signal
import subprocess
import sys
import threading
import time
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import requests

REPO_ROOT = Path(__file__).resolve().parent
MAIN_PY = REPO_ROOT / "main.py"
DEFAULT_TEST_URL = "https://www.google.com/generate_204"


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Movie Data Capture GUI")
        self.process: subprocess.Popen | None = None

        self.conf_var = tk.StringVar()
        self.over_var = tk.StringVar()
        self.search_var = tk.StringVar()
        self.specify_var = tk.StringVar()
        self.url_var = tk.StringVar()
        self.xlsx_var = tk.StringVar(value="output.xlsx")
        self.proxy_enabled_var = tk.BooleanVar(value=False)
        self.network_test_url_var = tk.StringVar(value=DEFAULT_TEST_URL)

        self.scrape_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()

        self.proxy_enabled_var = tk.BooleanVar(value=False)
        self.network_test_url_var = tk.StringVar(value=DEFAULT_TEST_URL)

        self.scrape_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()

        self.proxy_enabled_var = tk.BooleanVar(value=False)
        self.network_test_url_var = tk.StringVar(value=DEFAULT_TEST_URL)

        self._build_ui()
        self._load_conf_values_into_ui()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        opts = ttk.LabelFrame(frame, text="全局参数")
        opts.pack(fill=tk.X, pady=6)
        ttk.Label(opts, text="配置文件(--conf)").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(opts, textvariable=self.conf_var, width=70).grid(row=0, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(opts, text="选择", command=self.pick_conf).grid(row=0, column=2, padx=6, pady=4)

        ttk.Label(opts, text="搜刮路径(common.source_folder)").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(opts, textvariable=self.scrape_path_var, width=70).grid(row=1, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(opts, text="选择目录", command=self.pick_scrape_dir).grid(row=1, column=2, padx=6, pady=4)

        ttk.Label(opts, text="输出路径(common.success_output_folder)").grid(row=2, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(opts, textvariable=self.output_path_var, width=70).grid(row=2, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(opts, text="选择目录", command=self.pick_output_dir).grid(row=2, column=2, padx=6, pady=4)

        ttk.Label(opts, text="覆盖配置(--over-config，逗号分隔)").grid(row=3, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(opts, textvariable=self.over_var, width=70).grid(row=3, column=1, columnspan=2, sticky="ew", padx=6, pady=4)

        ttk.Checkbutton(opts, text="启用代理(proxy.switch)", variable=self.proxy_enabled_var).grid(
            row=4, column=0, sticky="w", padx=6, pady=4
        )
        ttk.Entry(opts, textvariable=self.network_test_url_var, width=45).grid(row=4, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(opts, text="测试网络连接", command=self.test_network_connection).grid(row=4, column=2, padx=6, pady=4)
        opts.columnconfigure(1, weight=1)

        modes = ttk.LabelFrame(frame, text="常用操作")
        modes.pack(fill=tk.X, pady=6)

        ttk.Label(modes, text="番号搜索").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(modes, textvariable=self.search_var, width=20).grid(row=0, column=1, padx=6, pady=4)
        ttk.Button(modes, text="运行", command=self.run_search).grid(row=0, column=2, padx=6, pady=4)

        ttk.Button(modes, text="列出待处理影片", command=self.run_list_movie).grid(row=1, column=0, padx=6, pady=4, sticky="w")
        ttk.Button(modes, text="自动评分", command=self.run_rate).grid(row=1, column=1, padx=6, pady=4, sticky="w")
        ttk.Button(modes, text="批量整理", command=self.run_batch_organize).grid(row=1, column=2, padx=6, pady=4, sticky="w")

        ttk.Label(modes, text="单文件刮削").grid(row=2, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(modes, textvariable=self.specify_var, width=55).grid(row=2, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(modes, text="选择文件", command=self.pick_file).grid(row=2, column=2, padx=6, pady=4)
        ttk.Button(modes, text="运行", command=self.run_specify).grid(row=2, column=3, padx=6, pady=4)

        ttk.Label(modes, text="URL抓取").grid(row=3, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(modes, textvariable=self.url_var, width=40).grid(row=3, column=1, sticky="ew", padx=6, pady=4)
        ttk.Entry(modes, textvariable=self.xlsx_var, width=20).grid(row=3, column=2, padx=6, pady=4)
        ttk.Button(modes, text="运行", command=self.run_scraping_url).grid(row=3, column=3, padx=6, pady=4)
        modes.columnconfigure(1, weight=1)

        actions = ttk.Frame(frame)
        actions.pack(fill=tk.X, pady=6)
        ttk.Button(actions, text="停止任务", command=self.stop_task).pack(side=tk.LEFT)
        ttk.Button(actions, text="清空日志", command=self.clear_log).pack(side=tk.LEFT, padx=8)

        self.log_text = tk.Text(frame, height=20)
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=6)

    def _load_conf_values_into_ui(self) -> None:
        try:
            conf_path = self._resolve_conf_path()
            parser = configparser.ConfigParser()
            parser.read(conf_path, encoding="utf-8")
            self.scrape_path_var.set(parser.get("common", "source_folder", fallback=""))
            self.output_path_var.set(parser.get("common", "success_output_folder", fallback=""))
            self.proxy_enabled_var.set(parser.getint("proxy", "switch", fallback=0) == 1)
        except Exception:
            pass

    def pick_conf(self) -> None:
        p = filedialog.askopenfilename(title="选择配置文件")
        if p:
            self.conf_var.set(p)
            self._load_conf_values_into_ui()

    def pick_file(self) -> None:
        p = filedialog.askopenfilename(title="选择视频文件")
        if p:
            self.specify_var.set(p)

    def pick_scrape_dir(self) -> None:
        p = filedialog.askdirectory(title="选择搜刮路径")
        if p:
            self.scrape_path_var.set(p)

    def pick_output_dir(self) -> None:
        p = filedialog.askdirectory(title="选择输出路径")
        if p:
            self.output_path_var.set(p)

    def log(self, msg: str) -> None:
        self.log_text.insert(tk.END, msg)
        self.log_text.see(tk.END)

    def clear_log(self) -> None:
        self.log_text.delete("1.0", tk.END)

    def _resolve_conf_path(self) -> Path:
        conf = self.conf_var.get().strip()
        if conf:
            conf_path = Path(conf)
            if not conf_path.exists():
                raise ValueError(f"配置文件不存在: {conf}")
            return conf_path

        for relative in ("config.ini", "static/config-default.ini"):
            candidate = REPO_ROOT / relative
            if candidate.exists():
                return candidate

        raise ValueError("未找到配置文件: config.ini 或 static/config-default.ini")

    def _resolve_network_settings(self) -> tuple[dict[str, str] | None, int]:
        timeout = 8
        proxy_url = ""
        try:
            conf_path = self._resolve_conf_path()
            parser = configparser.ConfigParser()
            parser.read(conf_path, encoding="utf-8")
            timeout = parser.getint("proxy", "timeout", fallback=timeout)
            proxy_url = parser.get("proxy", "url", fallback="").strip()
        except Exception:
            pass

        if self.proxy_enabled_var.get() and proxy_url:
            return ({"http": proxy_url, "https": proxy_url}, timeout)
        return (None, timeout)

    def _common_args(self) -> list[str]:
        args: list[str] = []
        conf = self.conf_var.get().strip()
        if conf:
            if not Path(conf).exists():
                raise ValueError(f"配置文件不存在: {conf}")
            args.extend(["--conf", conf])

        scrape_path = self.scrape_path_var.get().strip()
        if scrape_path:
            args.extend(["--over-config", f"common.source_folder={scrape_path}"])

        output_path = self.output_path_var.get().strip()
        if output_path:
            args.extend(["--over-config", f"common.success_output_folder={output_path}"])

        over = self.over_var.get().strip()
        if over:
            for item in [x.strip() for x in over.split(",") if x.strip()]:
                args.extend(["--over-config", item])

        proxy_switch = "1" if self.proxy_enabled_var.get() else "0"
        args.extend(["--over-config", f"proxy.switch={proxy_switch}"])
        return args

    def _start(self, mode_args: list[str]) -> None:
        if self.process and self.process.poll() is None:
            messagebox.showwarning("提示", "已有任务正在运行")
            return
        try:
            cmd = [sys.executable, str(MAIN_PY), *self._common_args(), *mode_args]
        except ValueError as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        self.log(f"\n$ {' '.join(cmd)}\n")
        popen_kwargs = {
            'cwd': str(REPO_ROOT),
            'stdout': subprocess.PIPE,
            'stderr': subprocess.STDOUT,
            'text': True,
            'bufsize': 1,
        }
        if os.name == 'nt':
            popen_kwargs['creationflags'] = subprocess.CREATE_NEW_PROCESS_GROUP
        elif hasattr(os, 'setsid'):
            popen_kwargs['preexec_fn'] = os.setsid

        self.process = subprocess.Popen(cmd, **popen_kwargs)
        threading.Thread(target=self._drain_output, daemon=True).start()

    def _drain_output(self) -> None:
        assert self.process is not None
        if self.process.stdout is None:
            return
        for line in self.process.stdout:
            self.root.after(0, self.log, line)
        code = self.process.wait()
        self.root.after(0, self.log, f"\n[process exited: {code}]\n")

    def _test_network_worker(self, url: str) -> None:
        proxies, timeout = self._resolve_network_settings()
        proxy_mode = "ON" if proxies else "OFF"
        self.root.after(0, self.log, f"\n[network-test] url={url} proxy={proxy_mode}\n")
        try:
            response = requests.get(url, timeout=timeout, proxies=proxies)
            self.root.after(
                0,
                self.log,
                f"[network-test] OK status={response.status_code} elapsed={response.elapsed.total_seconds():.2f}s\n",
            )
        except Exception as exc:
            self.root.after(0, self.log, f"[network-test] FAILED {exc}\n")

    def test_network_connection(self) -> None:
        url = self.network_test_url_var.get().strip()
        if not url:
            messagebox.showwarning("提示", "请输入测试 URL")
            return
        threading.Thread(target=self._test_network_worker, args=(url,), daemon=True).start()

    def stop_task(self) -> None:
        if not self.process or self.process.poll() is not None:
            self.log("\n[no running task]\n")
            return

        try:
            self.process.send_signal(signal.SIGINT)
            self.log("\n[sent SIGINT]\n")
        except Exception as exc:
            self.log(f"\n[send SIGINT failed: {exc}]\n")

        for _ in range(10):
            if self.process.poll() is not None:
                self.log("[task stopped]\n")
                return
            time.sleep(0.2)

        self.log("[SIGINT timeout, terminating process]\n")
        self.process.terminate()
        for _ in range(5):
            if self.process.poll() is not None:
                self.log("[task terminated]\n")
                return
            time.sleep(0.2)
        self.process.kill()
        self.log("[task killed]\n")

    def run_search(self) -> None:
        number = self.search_var.get().strip()
        if not number:
            messagebox.showwarning("提示", "请输入番号")
            return
        self._start(["--search", number])

    def run_list_movie(self) -> None:
        self._start(["--list-movie"])

    def run_rate(self) -> None:
        self._start(["--rate"])

    def run_batch_organize(self) -> None:
        self._start([])

    def run_specify(self) -> None:
        path = self.specify_var.get().strip()
        if not path:
            messagebox.showwarning("提示", "请选择文件")
            return
        self._start(["--specify-file", path])

    def run_scraping_url(self) -> None:
        url = self.url_var.get().strip()
        xlsx = self.xlsx_var.get().strip()
        if not url or not xlsx:
            messagebox.showwarning("提示", "请输入 URL 和 xlsx 文件名")
            return
        self._start(["--scraping-url", url, xlsx])


def main() -> None:
    root = tk.Tk()
    root.geometry("1080x820")
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
