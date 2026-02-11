
import os
import sys
import tkinter as tk
from tkinter import messagebox
import subprocess
import threading
import time

# 复用核心逻辑和界面
import generate_all

class ImageExportApp(generate_all.AllReportsApp):
    def _build_ui(self):
        # 1. 调用父类构建标准界面
        super()._build_ui()
        
        # 2. 在界面中插入“导出图片”勾选框
        # 我们寻找 data_var 输入框所在的 frame，或者直接在进度条上方插入
        # 为了简单，我们将其放在进度条 self.progress 的上方容器中
        
        # 为了美观，我们查找 self.progress 的父容器 (是 card_frame)
        card_frame = self.progress.master
        
        # 创建一个 Checkbox 容器
        chk_frame = tk.Frame(card_frame, bg="#FFFFFF")
        chk_frame.pack(after=self.progress, fill="x", pady=(5, 0)) # 放在进度条下面一点点

        self.export_imgs_var = tk.BooleanVar(value=False)
        
        chk = tk.Checkbutton(
            chk_frame, 
            text="同时导出为图片 (需要安装 PowerPoint)", 
            variable=self.export_imgs_var,
            font=("Microsoft YaHei UI", 10) if sys.platform == "win32" else ("PingFang SC", 10),
            bg="#FFFFFF", 
            activebackground="#FFFFFF",
            fg="#333333",
            selectcolor="#FFFFFF"
        )
        chk.pack(side="left")

    def _on_generate(self):
        """重写生成逻辑：先生成PPT，再决定是否转图片"""
        
        # 1. 拦截原始的生成线程，改为运行我们自己的混合任务
        # 父类是直接开启线程 _run_generation，我们需要 copy 部分逻辑或者 wrapper
        
        template_path = self.template_var.get()
        data_path = self.data_var.get()

        if not os.path.exists(template_path):
            messagebox.showerror("错误", "模板文件不存在")
            return
        if not os.path.exists(data_path):
            messagebox.showerror("错误", "数据文件不存在")
            return

        self.gen_btn.config(state="disabled", text="⏳ 正在生成...")
        self.progress["value"] = 0
        self.status_var.set("正在初始化...")

        # 在新线程中运行
        threading.Thread(target=self._run_process, args=(template_path, data_path)).start()

    def _run_process(self, template_path, data_path):
        try:
            # 1. 调用父类的生成逻辑 (静态方法复用是个问题，父类的方法混杂了 self)
            # 我们通过组合方式：直接实例化父逻辑太复杂，不如重构父类
            # 但为了不修改 generate_all.py，我们只能 复制粘贴父类的 _run_generation 核心逻辑
            # 或者... 我们可以利用 Python 动态特性调用父类方法，但父类方法是绑定了 GUI 更新的。
            
            # 最佳方案：让父类的 _run_generation 完成后，我们再接手。
            # 但父类 _run_generation 是 threaded 的，且最后会调用 _on_done。
            # 我们 Hook _on_done！
            pass
        except Exception as e:
            pass

    # --- 采用 HOOK 方案 ---
    def _run_generation_wrapped(self):
        # 这个方法没法用，因为父类点击按钮直接触发 thread
        pass
        
    # --- 实际方案：覆盖 _on_done ---
    # 父类生成结束后会调用 _on_done(msg, output_path)
    def _on_done(self, msg, output_path):
        if not self.export_imgs_var.get():
            # 用户没勾选，直接结束
            super()._on_done(msg, output_path)
            return

        # 用户勾选了，开始转图片
        self.root.after(0, lambda: self.status_var.set("📊 正在调用 PowerPoint 导出图片..."))
        
        threading.Thread(target=self._convert_to_images_thread, args=(output_path,)).start()

    def _convert_to_images_thread(self, pptx_path):
        try:
            images_dir = os.path.splitext(pptx_path)[0] + "_图片导出"
            if not os.path.exists(images_dir):
                os.makedirs(images_dir)
            
            error_msg = None
            
            if sys.platform == "win32":
                self._convert_win32(pptx_path, images_dir)
            else:
                self._convert_mac(pptx_path, images_dir)
                
            # 完成
            self.root.after(0, lambda: super(ImageExportApp, self)._on_done("生成并导出图片成功！", pptx_path))
            
        except Exception as e:
            err = str(e)
            print(err)
            self.root.after(0, lambda: messagebox.showerror("导出图片失败", f"PPT生成成功，但导出图片失败。\n可能原因：未安装Office或权限不足。\n\n错误信息：{err}"))
            self.root.after(0, lambda: super(ImageExportApp, self)._on_done("仅PPT生成成功", pptx_path))

    def _convert_win32(self, pptx_path, output_dir):
        import win32com.client
        
        pptx_path = os.path.abspath(pptx_path)
        output_dir = os.path.abspath(output_dir)
        
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        # powerpoint.Visible = True # 保持后台
        
        presentation = powerpoint.Presentations.Open(pptx_path, WithWindow=False)
        
        # 另存为图片
        # ppSaveAsJPG = 17, ppSaveAsPNG = 18
        presentation.SaveAs(os.path.join(output_dir, "Slide.jpg"), 17)
        
        presentation.Close()
        # powerpoint.Quit() # 不退出 App，防止杀掉用户正在用的 PPT

    def _convert_mac(self, pptx_path, output_dir):
        pptx_path = os.path.abspath(pptx_path)
        output_dir = os.path.abspath(output_dir)
        
        # AppleScript 脚本
        scpt = f'''
        tell application "Microsoft PowerPoint"
            set pptOpen to open "{pptx_path}"
            save pptOpen in "{output_dir}" as save as JPG
            close pptOpen
        end tell
        '''
        
        p = subprocess.Popen(['osascript', '-e', scpt], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        out, err = p.communicate()
        if p.returncode != 0:
            raise Exception(f"AppleScript Error: {err.decode('utf-8')}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ImageExportApp(root)
    root.mainloop()
