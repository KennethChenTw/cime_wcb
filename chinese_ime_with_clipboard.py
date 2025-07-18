# --- START OF FILE chinese_ime_with_clipboard_refactored.py ---

import tkinter as tk
from tkinter import ttk, messagebox, font
import pyperclip
import json
import os
import pickle


class ClipboardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文字複製工具")
        self.root.resizable(False, False)

        # 設定檔案
        self.history_file = "clipboard_history.json"
        self.word_tab_file = "word.tab"
        self.settings_file = "app_settings.json"
        
        # 模式控制變數初始化（必須在載入設定之前）
        self.is_chinese_mode = tk.BooleanVar(value=False)
        self.hotkey_var = tk.StringVar(value="ctrl+k")
        self.preselect_mode = tk.BooleanVar(value=False)
        self.vr_candidate_mode = tk.BooleanVar(value=False)
        
        # 載入設定
        self.load_settings()
        
        # 設定視窗位置和大小
        self.apply_window_settings()
        
        # 載入歷史和詞彙
        self.load_history()
        # *** 修改：初始化字典來儲存詞彙，而不是列表 ***
        self.word_dictionary = {}
        self.load_word_tab()

        # 中文輸入候選清單
        self.candidates = []
        self.selection_dialog = None

        # 焦點追蹤
        self.focused_widget = None

        # 設定字型
        self.setup_fonts()
        
        self.setup_ui()
        self.bind_events()

    def load_settings(self):
        """載入設定檔案"""
        default_settings = {
            "window_x": 100,
            "window_y": 100,
            "window_width": 500,  # 增加預設寬度以容納VR選項
            "window_height": 460,
            "font_size": 12,
            "font_family": "Arial",
            # 新增候選視窗設定
            "candidate_window_width": 300,
            "candidate_window_height": 200,
            "candidate_font_size": 12,
            "candidate_font_family": "Arial",
            # 新增VR候選簡碼設定
            "vr_candidate_mode": False
        }
        
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, "r", encoding="utf-8") as f:
                    self.settings = json.load(f)
                # 確保所有必要的設定都存在
                for key, value in default_settings.items():
                    if key not in self.settings:
                        self.settings[key] = value
            except Exception as e:
                messagebox.showerror("錯誤", f"讀取設定檔案失敗: {e}")
                self.settings = default_settings
        else:
            self.settings = default_settings
        
        # 載入VR候選簡碼設定
        self.vr_candidate_mode.set(self.settings.get("vr_candidate_mode", False))

    def save_settings(self):
        """儲存設定檔案"""
        try:
            # 更新當前視窗位置
            self.settings["window_x"] = self.root.winfo_x()
            self.settings["window_y"] = self.root.winfo_y()
            # 儲存VR候選簡碼設定
            self.settings["vr_candidate_mode"] = self.vr_candidate_mode.get()
            
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存設定檔案失敗: {e}")

    def apply_window_settings(self):
        """應用視窗設定"""
        x = self.settings["window_x"]
        y = self.settings["window_y"]
        width = self.settings["window_width"]
        height = self.settings["window_height"]
        
        # 確保視窗位置在螢幕範圍內
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        if x < 0 or x > screen_width - width:
            x = 100
        if y < 0 or y > screen_height - height:
            y = 100
            
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def setup_fonts(self):
        """設定字型"""
        font_size = self.settings["font_size"]
        font_family = self.settings["font_family"]
        
        self.default_font = font.Font(family=font_family, size=font_size)
        self.entry_font = font.Font(family=font_family, size=font_size + 2)
        self.button_font = font.Font(family=font_family, size=font_size - 1)
        self.label_font = font.Font(family=font_family, size=font_size)
        self.title_font = font.Font(family=font_family, size=font_size + 1, weight="bold")
        
        # 設定候選視窗字型
        candidate_font_size = self.settings["candidate_font_size"]
        candidate_font_family = self.settings["candidate_font_family"]
        
        self.candidate_default_font = font.Font(family=candidate_font_family, size=candidate_font_size)
        self.candidate_button_font = font.Font(family=candidate_font_family, size=candidate_font_size - 1)
        self.candidate_label_font = font.Font(family=candidate_font_family, size=candidate_font_size)
        self.candidate_title_font = font.Font(family=candidate_font_family, size=candidate_font_size + 1, weight="bold")

    def calculate_window_size(self):
        """根據字型大小計算視窗大小"""
        base_width = 500  # 增加基礎寬度
        base_height = 460
        font_size = self.settings["font_size"]
        
        # 根據字型大小調整視窗大小
        scale_factor = font_size / 12  # 12是基準字型大小
        new_width = int(base_width * scale_factor)
        new_height = int(base_height * scale_factor)
        
        # 設定最小和最大尺寸
        new_width = max(450, min(800, new_width))  # 增加最小寬度
        new_height = max(400, min(800, new_height))
        
        return new_width, new_height

    def update_window_size(self):
        """更新視窗大小"""
        width, height = self.calculate_window_size()
        self.settings["window_width"] = width
        self.settings["window_height"] = height
        
        # 保持視窗位置，只更新大小
        x = self.root.winfo_x()
        y = self.root.winfo_y()
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def get_candidate_window_position(self):
        """計算候選視窗的位置，使其在主視窗下方並行"""
        main_x = self.root.winfo_x()
        main_y = self.root.winfo_y()
        main_height = self.root.winfo_height()
        
        # 候選視窗放在主視窗正下方，稍微間隔一點
        candidate_x = main_x
        candidate_y = main_y + main_height + 10
        
        # 確保候選視窗不會超出螢幕範圍
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        candidate_width = self.settings["candidate_window_width"]
        candidate_height = self.settings["candidate_window_height"]
        
        # 調整X位置
        if candidate_x + candidate_width > screen_width:
            candidate_x = screen_width - candidate_width - 10
        if candidate_x < 0:
            candidate_x = 0
            
        # 調整Y位置
        if candidate_y + candidate_height > screen_height:
            # 如果下方放不下，就放在主視窗上方
            candidate_y = main_y - candidate_height - 10
            if candidate_y < 0:
                # 如果上方也放不下，就放在右側
                candidate_x = main_x + self.root.winfo_width() + 10
                candidate_y = main_y
                
        return candidate_x, candidate_y

    def is_candidate_window_open(self):
        """檢查候選視窗是否開啟"""
        try:
            return self.selection_dialog is not None and self.selection_dialog.winfo_exists()
        except tk.TclError:
            return False

    def on_focus_in(self, event):
        """當元件獲得焦點時記錄"""
        self.focused_widget = event.widget

    def on_focus_out(self, event):
        """當元件失去焦點時清除記錄"""
        if self.focused_widget == event.widget:
            self.focused_widget = None

    def on_chinese_key_press(self, event):
        """處理中文輸入框的按鍵事件"""
        char = event.char
        
        # 檢查是否為數字且候選視窗未開啟
        if char.isdigit() and not self.is_candidate_window_open():
            # 將數字填入主輸入框
            current_main_text = self.entry.get()
            self.entry.delete(0, tk.END)
            self.entry.insert(0, current_main_text + char)
            return "break"  # 阻止預設行為
        
        # 其他字元維持原有行為
        return None

    def find_word_matches_with_vr(self, input_code):
        """搜尋詞語匹配，支援VR候選簡碼作為後備機制"""
        # 如果VR候選簡碼模式未啟用，直接使用原始搜尋
        if not self.vr_candidate_mode.get():
            return self.find_word_matches(input_code)
        
        # 1. 先嘗試exact match（優先級最高）
        exact_matches = self.find_word_matches(input_code)
        if exact_matches:  # 如果有exact match，直接返回
            return exact_matches
        
        # 2. 如果沒有exact match，且啟用VR模式，則檢查是否為VR簡碼格式
        is_v_shortcode = False
        is_r_shortcode = False
        base_code = ""
        
        # 檢查V候選簡碼 (3碼+V/v 或 4碼+V/v)
        if len(input_code) >= 4 and input_code[-1].upper() == 'V':
            is_v_shortcode = True
            base_code = input_code[:-1]  # 移除最後的V/v
        
        # 檢查R候選簡碼 (3碼+R/r 或 4碼+R/r)  
        elif len(input_code) >= 4 and input_code[-1].upper() == 'R':
            is_r_shortcode = True
            base_code = input_code[:-1]  # 移除最後的R/r
        
        # 3. 如果是VR候選簡碼格式，尋找base_code的候選詞
        if is_v_shortcode or is_r_shortcode:
            base_matches = self.find_word_matches(base_code)
            
            if is_v_shortcode and len(base_matches) > 1:
                # V代表第二個候選(索引1)
                return [base_matches[1]]
            elif is_r_shortcode and len(base_matches) > 2:
                # R代表第三個候選(索引2)
                return [base_matches[2]]
        
        # 4. 如果以上都沒有找到，返回空列表
        return []

    def open_settings_dialog(self):
        """開啟設定對話框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("設定")
        dialog.geometry("450x550")  # 增加高度以容納新選項
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        # 建立捲軸框架
        canvas = tk.Canvas(dialog)
        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # === 功能設定 ===
        feature_frame = tk.LabelFrame(scrollable_frame, text="功能設定", font=self.title_font)
        feature_frame.pack(pady=5, padx=10, fill="x")

        # VR候選簡碼設定
        vr_candidate_var = tk.BooleanVar(value=self.vr_candidate_mode.get())
        vr_check = tk.Checkbutton(feature_frame, text="啟用VR候選簡碼", variable=vr_candidate_var, 
                                  font=self.label_font)
        vr_check.pack(anchor="w", padx=5, pady=2)
        
        # VR候選簡碼說明（更新說明）
        vr_info = tk.Label(feature_frame, text="• 找不到完全匹配時的後備機制\n• 三碼+V/v 或 四碼+V/v：選擇第二個候選詞\n• 三碼+R/r 或 四碼+R/r：選擇第三個候選詞", 
                          font=self.label_font, fg="gray", justify="left")
        vr_info.pack(anchor="w", padx=20, pady=2)

        # === 主視窗設定 ===
        main_frame = tk.LabelFrame(scrollable_frame, text="主視窗設定", font=self.title_font)
        main_frame.pack(pady=5, padx=10, fill="x")

        # 主視窗字型設定
        main_font_frame = tk.LabelFrame(main_frame, text="字型設定")
        main_font_frame.pack(pady=5, padx=5, fill="x")

        tk.Label(main_font_frame, text="字型:", font=self.label_font).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        main_font_family_var = tk.StringVar(value=self.settings["font_family"])
        main_font_combo = ttk.Combobox(main_font_frame, textvariable=main_font_family_var, 
                                     values=["Arial", "Microsoft JhengHei", "新細明體", "標楷體", "Times New Roman", "Courier New"],
                                     state="readonly", width=15)
        main_font_combo.grid(row=0, column=1, padx=5, pady=2)

        tk.Label(main_font_frame, text="大小:", font=self.label_font).grid(row=1, column=0, sticky="w", padx=5, pady=2)
        main_font_size_var = tk.StringVar(value=str(self.settings["font_size"]))
        main_size_combo = ttk.Combobox(main_font_frame, textvariable=main_font_size_var,
                                     values=["8", "9", "10", "11", "12", "13", "14", "15", "16", "18", "20", "24"],
                                     state="readonly", width=15)
        main_size_combo.grid(row=1, column=1, padx=5, pady=2)

        # 主視窗位置設定
        main_position_frame = tk.LabelFrame(main_frame, text="位置設定")
        main_position_frame.pack(pady=5, padx=5, fill="x")

        tk.Label(main_position_frame, text="X位置:", font=self.label_font).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        main_x_var = tk.StringVar(value=str(self.settings["window_x"]))
        main_x_entry = tk.Entry(main_position_frame, textvariable=main_x_var, width=10, font=self.default_font)
        main_x_entry.grid(row=0, column=1, padx=5, pady=2)

        tk.Label(main_position_frame, text="Y位置:", font=self.label_font).grid(row=1, column=0, sticky="w", padx=5, pady=2)
        main_y_var = tk.StringVar(value=str(self.settings["window_y"]))
        main_y_entry = tk.Entry(main_position_frame, textvariable=main_y_var, width=10, font=self.default_font)
        main_y_entry.grid(row=1, column=1, padx=5, pady=2)

        # 主視窗快速設定按鈕
        main_quick_frame = tk.Frame(main_position_frame)
        main_quick_frame.grid(row=2, column=0, columnspan=2, pady=5)
        
        tk.Button(main_quick_frame, text="置中", font=self.button_font,
                 command=lambda: self.set_center_position(main_x_var, main_y_var)).pack(side=tk.LEFT, padx=2)
        tk.Button(main_quick_frame, text="左上角", font=self.button_font,
                 command=lambda: self.set_corner_position(main_x_var, main_y_var, 50, 50)).pack(side=tk.LEFT, padx=2)
        tk.Button(main_quick_frame, text="右上角", font=self.button_font,
                 command=lambda: self.set_corner_position(main_x_var, main_y_var, 
                                                        self.root.winfo_screenwidth() - self.settings["window_width"] - 50, 50)).pack(side=tk.LEFT, padx=2)

        # === 候選視窗設定 ===
        candidate_frame = tk.LabelFrame(scrollable_frame, text="候選視窗設定", font=self.title_font)
        candidate_frame.pack(pady=5, padx=10, fill="x")

        # 候選視窗字型設定
        candidate_font_frame = tk.LabelFrame(candidate_frame, text="字型設定")
        candidate_font_frame.pack(pady=5, padx=5, fill="x")

        tk.Label(candidate_font_frame, text="字型:", font=self.label_font).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        candidate_font_family_var = tk.StringVar(value=self.settings["candidate_font_family"])
        candidate_font_combo = ttk.Combobox(candidate_font_frame, textvariable=candidate_font_family_var, 
                                          values=["Arial", "Microsoft JhengHei", "新細明體", "標楷體", "Times New Roman", "Courier New"],
                                          state="readonly", width=15)
        candidate_font_combo.grid(row=0, column=1, padx=5, pady=2)

        tk.Label(candidate_font_frame, text="大小:", font=self.label_font).grid(row=1, column=0, sticky="w", padx=5, pady=2)
        candidate_font_size_var = tk.StringVar(value=str(self.settings["candidate_font_size"]))
        candidate_size_combo = ttk.Combobox(candidate_font_frame, textvariable=candidate_font_size_var,
                                          values=["8", "9", "10", "11", "12", "13", "14", "15", "16", "18", "20", "24"],
                                          state="readonly", width=15)
        candidate_size_combo.grid(row=1, column=1, padx=5, pady=2)

        # 候選視窗大小設定
        candidate_size_frame = tk.LabelFrame(candidate_frame, text="視窗大小")
        candidate_size_frame.pack(pady=5, padx=5, fill="x")

        tk.Label(candidate_size_frame, text="寬度:", font=self.label_font).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        candidate_width_var = tk.StringVar(value=str(self.settings["candidate_window_width"]))
        candidate_width_entry = tk.Entry(candidate_size_frame, textvariable=candidate_width_var, width=10, font=self.default_font)
        candidate_width_entry.grid(row=0, column=1, padx=5, pady=2)

        tk.Label(candidate_size_frame, text="高度:", font=self.label_font).grid(row=1, column=0, sticky="w", padx=5, pady=2)
        candidate_height_var = tk.StringVar(value=str(self.settings["candidate_window_height"]))
        candidate_height_entry = tk.Entry(candidate_size_frame, textvariable=candidate_height_var, width=10, font=self.default_font)
        candidate_height_entry.grid(row=1, column=1, padx=5, pady=2)

        # 按鈕區域
        button_frame = tk.Frame(scrollable_frame)
        button_frame.pack(pady=20)

        def apply_settings():
            try:
                # 驗證並應用設定
                new_main_x = int(main_x_var.get())
                new_main_y = int(main_y_var.get())
                new_font_size = int(main_font_size_var.get())
                new_font_family = main_font_family_var.get()
                
                new_candidate_font_size = int(candidate_font_size_var.get())
                new_candidate_font_family = candidate_font_family_var.get()
                new_candidate_width = int(candidate_width_var.get())
                new_candidate_height = int(candidate_height_var.get())

                # 更新設定
                self.settings["window_x"] = new_main_x
                self.settings["window_y"] = new_main_y
                self.settings["font_size"] = new_font_size
                self.settings["font_family"] = new_font_family
                self.settings["candidate_font_size"] = new_candidate_font_size
                self.settings["candidate_font_family"] = new_candidate_font_family
                self.settings["candidate_window_width"] = new_candidate_width
                self.settings["candidate_window_height"] = new_candidate_height
                # 更新VR候選簡碼設定
                self.vr_candidate_mode.set(vr_candidate_var.get())

                # 重新設定字型
                self.setup_fonts()
                
                # 更新所有UI元件的字型
                self.update_all_fonts()
                
                # 更新視窗大小和位置
                self.update_window_size()
                self.apply_window_settings()
                
                # 儲存設定
                self.save_settings()
                
                messagebox.showinfo("成功", "設定已套用並儲存")
                dialog.destroy()
                
            except ValueError:
                messagebox.showerror("錯誤", "請輸入有效的數值")

        def preview_main_font():
            try:
                preview_size = int(main_font_size_var.get())
                preview_family = main_font_family_var.get()
                self.show_font_preview(preview_family, preview_size, "主視窗字型預覽", dialog)
            except ValueError:
                messagebox.showerror("錯誤", "請選擇有效的字型大小")

        def preview_candidate_font():
            try:
                preview_size = int(candidate_font_size_var.get())
                preview_family = candidate_font_family_var.get()
                self.show_font_preview(preview_family, preview_size, "候選視窗字型預覽", dialog)
            except ValueError:
                messagebox.showerror("錯誤", "請選擇有效的字型大小")

        tk.Button(button_frame, text="套用並儲存", font=self.button_font, command=apply_settings).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="預覽主視窗字型", font=self.button_font, command=preview_main_font).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="預覽候選字型", font=self.button_font, command=preview_candidate_font).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="取消", font=self.button_font, command=dialog.destroy).pack(side=tk.LEFT, padx=5)

        # 設定捲軸
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 綁定滑鼠滾輪
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)

    def show_font_preview(self, font_family, font_size, title, parent_window):
        """顯示字型預覽視窗"""
        try:
            preview_font = font.Font(family=font_family, size=font_size)
            # 使用系統預設字型作為按鈕字型，確保按鈕可以正常顯示
            button_font = font.Font(family="Arial", size=10)
        except Exception as e:
            messagebox.showerror("錯誤", f"無法建立預覽字型: {e}")
            return
        
        preview_window = tk.Toplevel(parent_window)
        preview_window.title(title)
        preview_window.geometry("400x220")
        preview_window.transient(parent_window)
        preview_window.grab_set()  # 確保預覽視窗獲得焦點
        preview_window.resizable(False, False)
        
        # 設定視窗位置（在父視窗中央）
        parent_x = parent_window.winfo_x()
        parent_y = parent_window.winfo_y()
        parent_width = parent_window.winfo_width()
        parent_height = parent_window.winfo_height()
        
        preview_x = parent_x + (parent_width - 400) // 2
        preview_y = parent_y + (parent_height - 220) // 2
        
        # 確保視窗在螢幕範圍內
        screen_width = preview_window.winfo_screenwidth()
        screen_height = preview_window.winfo_screenheight()
        if preview_x < 0:
            preview_x = 0
        elif preview_x + 400 > screen_width:
            preview_x = screen_width - 400
        if preview_y < 0:
            preview_y = 0
        elif preview_y + 220 > screen_height:
            preview_y = screen_height - 220
            
        preview_window.geometry(f"400x220+{preview_x}+{preview_y}")
        
        # 建立預覽內容
        content_frame = tk.Frame(preview_window)
        content_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        try:
            tk.Label(content_frame, text=f"字型預覽 Font Preview", 
                    font=preview_font).pack(pady=5)
            tk.Label(content_frame, text=f"字型: {font_family}, 大小: {font_size}", 
                    font=preview_font).pack(pady=5)
            tk.Label(content_frame, text="中文測試 English Test 123", 
                    font=preview_font).pack(pady=5)
            tk.Label(content_frame, text="候選詞語選擇 Candidate Selection", 
                    font=preview_font).pack(pady=5)
            tk.Label(content_frame, text="ABCDEFG abcdefg 一二三四五", 
                    font=preview_font).pack(pady=5)
        except Exception as e:
            tk.Label(content_frame, text=f"字型預覽錯誤: {e}", 
                    font=button_font).pack(pady=10)
        
        # 按鈕區域
        button_frame = tk.Frame(content_frame)
        button_frame.pack(pady=15)
        
        def close_preview():
            try:
                preview_window.destroy()
            except:
                pass
        
        # 使用系統字型確保按鈕正常顯示
        close_button = tk.Button(button_frame, text="關閉", font=button_font, command=close_preview)
        close_button.pack()
        
        # 綁定ESC鍵關閉
        preview_window.bind("<Escape>", lambda e: close_preview())
        
        # 綁定Enter鍵關閉
        preview_window.bind("<Return>", lambda e: close_preview())
        
        # 確保視窗關閉時釋放焦點
        def on_destroy():
            try:
                preview_window.grab_release()
                parent_window.focus_set()
            except:
                pass
        
        preview_window.protocol("WM_DELETE_WINDOW", lambda: (on_destroy(), close_preview()))
        
        # 設定焦點到關閉按鈕
        close_button.focus_set()

    def set_center_position(self, x_var, y_var):
        """設定視窗置中"""
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = self.settings["window_width"]
        window_height = self.settings["window_height"]
        
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        x_var.set(str(x))
        y_var.set(str(y))

    def set_corner_position(self, x_var, y_var, x, y):
        """設定視窗到指定角落"""
        x_var.set(str(x))
        y_var.set(str(y))

    def update_all_fonts(self):
        """更新所有UI元件的字型"""
        def update_widget_font(widget, font_obj):
            try:
                widget.configure(font=font_obj)
            except:
                pass

        # 更新主視窗中的所有元件
        self.mode_label.configure(font=self.title_font)
        self.entry.configure(font=self.entry_font)
        self.chinese_entry.configure(font=self.entry_font)
        self.history_label.configure(font=self.title_font)
        self.history_listbox.configure(font=self.default_font)

        # 更新所有標籤
        for widget in self.root.winfo_children():
            self.update_widget_fonts_recursive(widget)

    def update_widget_fonts_recursive(self, widget):
        """遞歸更新widget及其子元件的字型"""
        widget_class = widget.winfo_class()
        
        if widget_class == "Label":
            widget.configure(font=self.label_font)
        elif widget_class == "Button":
            widget.configure(font=self.button_font)
        elif widget_class == "Entry":
            if widget == self.entry or widget == self.chinese_entry:
                widget.configure(font=self.entry_font)
            else:
                widget.configure(font=self.default_font)
        elif widget_class == "Listbox":
            widget.configure(font=self.default_font)
        
        # 遞歸處理子元件
        for child in widget.winfo_children():
            self.update_widget_fonts_recursive(child)

    def setup_ui(self):
        # 模式切換區域
        self.mode_frame = tk.Frame(self.root)
        self.mode_frame.pack(pady=5)
        tk.Label(self.mode_frame, text="模式:", font=self.label_font).pack(side=tk.LEFT)
        self.mode_label = tk.Label(self.mode_frame, text="英文", font=self.title_font, fg="blue")
        self.mode_label.pack(side=tk.LEFT, padx=5)

        # 快速鍵設定與功能選項 - 分成兩行以避免超出視窗
        self.hotkey_frame = tk.Frame(self.root)
        self.hotkey_frame.pack(pady=2)
        
        # 第一行：快速鍵和先上字模式
        self.hotkey_line1 = tk.Frame(self.hotkey_frame)
        self.hotkey_line1.pack(fill="x")
        tk.Label(self.hotkey_line1, text="切換快速鍵:", font=self.label_font).pack(side=tk.LEFT)
        hotkey_combo = ttk.Combobox(self.hotkey_line1, textvariable=self.hotkey_var,
                                   values=["ctrl+k", "ctrl+space", "capslock"],
                                   state="readonly", width=10)
        hotkey_combo.pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(self.hotkey_line1, text="先上字模式", variable=self.preselect_mode, 
                      font=self.label_font).pack(side=tk.LEFT, padx=5)
        
        # 第二行：VR候選簡碼
        self.hotkey_line2 = tk.Frame(self.hotkey_frame)
        self.hotkey_line2.pack(fill="x", pady=(2, 0))
        tk.Checkbutton(self.hotkey_line2, text="VR候選簡碼", variable=self.vr_candidate_mode, 
                      font=self.label_font).pack(side=tk.LEFT)

        # 主輸入框
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(pady=5)
        tk.Label(self.main_frame, text="主輸入框:", font=self.label_font).pack()
        self.entry = tk.Entry(self.main_frame, font=self.entry_font, width=30)
        self.entry.pack(pady=2)

        # 中文輸入框 (初始隱藏)
        self.chinese_frame = tk.Frame(self.root)
        tk.Label(self.chinese_frame, text="中文輸入 (最多6字):", font=self.label_font).pack()
        self.chinese_entry = tk.Entry(self.chinese_frame, font=self.entry_font, width=30)
        self.chinese_entry.pack(pady=2)

        # 候選清單
        self.candidate_frame = tk.Frame(self.chinese_frame)
        self.candidate_frame.pack(pady=2)

        # 按鈕區域
        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(pady=5)
        tk.Button(self.button_frame, text="手動切換模式", font=self.button_font, 
                 command=self.toggle_mode).pack(side=tk.LEFT, padx=2)
        tk.Button(self.button_frame, text="清空輸入框", font=self.button_font, 
                 command=self.clear_entry).pack(side=tk.LEFT, padx=2)
        tk.Button(self.button_frame, text="清除歷史", font=self.button_font, 
                 command=self.clear_history).pack(side=tk.LEFT, padx=2)
        tk.Button(self.button_frame, text="設定", font=self.button_font, 
                 command=self.open_settings_dialog).pack(side=tk.LEFT, padx=2)

        # 歷史紀錄
        self.history_label = tk.Label(self.root, text="歷史紀錄", font=self.title_font)
        self.history_label.pack(pady=5)
        self.history_listbox = tk.Listbox(self.root, width=50, height=8, font=self.default_font)
        self.history_listbox.pack()

        # 載入歷史
        for item in self.history:
            self.history_listbox.insert(tk.END, item)

    def bind_events(self):
        # 主輸入框事件
        self.entry.bind("<Return>", self.on_enter)
        self.entry.bind("<FocusIn>", self.on_focus_in)
        self.entry.bind("<FocusOut>", self.on_focus_out)

        # 中文輸入框事件
        self.chinese_entry.bind("<space>", self.on_chinese_space)
        self.chinese_entry.bind("<KeyRelease>", self.on_chinese_input)
        self.chinese_entry.bind("<Return>", self.on_enter_from_chinese)
        self.chinese_entry.bind("<FocusIn>", self.on_focus_in)
        self.chinese_entry.bind("<FocusOut>", self.on_focus_out)
        # 新增：綁定按鍵事件到中文輸入框，處理數字輸入
        self.chinese_entry.bind("<KeyPress>", self.on_chinese_key_press)

        # 數字鍵選擇候選詞
        for i in range(10):
            self.root.bind(str(i), lambda e, num=i: self.select_candidate_by_number(num))

        # 英文字母與標點處理（只在先上字模式下啟用）
        self.root.bind("<Key>", self.handle_letter_input)

        # 歷史選擇事件
        self.history_listbox.bind("<<ListboxSelect>>", self.on_history_select)
        self.history_listbox.bind("<FocusIn>", self.on_focus_in)
        self.history_listbox.bind("<FocusOut>", self.on_focus_out)

        # 全域快速鍵
        self.root.bind("<Control-k>", lambda e: self.check_hotkey("ctrl+k"))
        self.root.bind("<Control-space>", lambda e: self.check_hotkey("ctrl+space"))
        self.root.bind("<KeyRelease-Caps_Lock>", lambda e: self.check_hotkey("capslock"))

        # 關閉事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def check_hotkey(self, hotkey):
        if self.hotkey_var.get() == hotkey:
            self.toggle_mode()

    def toggle_mode(self):
        self.is_chinese_mode.set(not self.is_chinese_mode.get())
        if self.is_chinese_mode.get():
            self.mode_label.config(text="中文", fg="red")
            self.chinese_frame.pack(after=self.main_frame, pady=5)
            self.chinese_entry.focus()
        else:
            self.mode_label.config(text="英文", fg="blue")
            self.chinese_frame.pack_forget()
            self.entry.focus()
        self.clear_candidates()
        self.close_selection_dialog()

    def on_chinese_input(self, event):
        current_text = self.chinese_entry.get()
        if len(current_text) > 6:
            self.chinese_entry.delete(6, tk.END)
        self.clear_candidates()
        self.candidates = []

    def on_chinese_space(self, event):
        input_text = self.chinese_entry.get().strip()
        if not input_text:
            return "break"

        # 使用支援VR候選簡碼的搜尋方法
        matches = self.find_word_matches_with_vr(input_text)
        
        if len(matches) == 1:
            current_text = self.entry.get()
            self.entry.delete(0, tk.END)
            self.entry.insert(0, current_text + matches[0])
        elif len(matches) > 1:
            if self.preselect_mode.get():
                current_text = self.entry.get()
                self.entry.delete(0, tk.END)
                self.entry.insert(0, current_text + matches[0])
                self.candidates = matches
                self.show_selection_dialog(matches)
            else:
                self.candidates = matches
                self.show_selection_dialog(matches)
        else:
            messagebox.showinfo("提示", f"找不到 '{input_text}' 對應的詞語")

        self.chinese_entry.delete(0, tk.END)
        return "break"

    # *** 修改：使用字典進行高效查詢 ***
    def find_word_matches(self, input_code):
        """
        (重構後) 原始的詞語搜尋方法。
        現在使用字典進行快速查找，而不是遍歷列表。
        """
        # 使用 .get() 方法，如果找不到鍵，就返回一個空列表
        return self.word_dictionary.get(input_code, [])

    def show_selection_dialog(self, matches):
        self.close_selection_dialog()
        
        dialog = tk.Toplevel(self.root)
        dialog.title("選擇詞語")
        
        # 使用設定中的候選視窗大小
        width = self.settings["candidate_window_width"]
        height = self.settings["candidate_window_height"]
        
        # 計算候選視窗位置
        x, y = self.get_candidate_window_position()
        
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.focus_force()
        dialog.resizable(False, False)

        self.selection_dialog = dialog
        self.chinese_entry.config(state=tk.DISABLED)

        def on_close():
            self.chinese_entry.config(state=tk.NORMAL)
            self.chinese_entry.delete(0, tk.END)
            self.chinese_entry.focus()
            self.close_selection_dialog()

        dialog.protocol("WM_DELETE_WINDOW", on_close)

        # 使用候選視窗專用字型
        tk.Label(dialog, text="請選擇詞語:", font=self.candidate_title_font).pack(pady=5)

        listbox = tk.Listbox(dialog, height=8, font=self.candidate_default_font)
        listbox.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

        for i, word in enumerate(matches):
            listbox.insert(tk.END, f"{i}: {word}")

        def on_select():
            selection = listbox.curselection()
            if selection:
                selected_word = matches[selection[0]]
                if self.preselect_mode.get():
                    current_text = self.entry.get()
                    if current_text and len(matches) > 0:
                        first_word = matches[0]
                        if current_text.endswith(first_word):
                            current_text = current_text[:-len(first_word)]
                    self.entry.delete(0, tk.END)
                    self.entry.insert(0, current_text + selected_word)
                else:
                    current_text = self.entry.get()
                    self.entry.delete(0, tk.END)
                    self.entry.insert(0, current_text + selected_word)
            on_close()

        def on_double_click(event):
            on_select()

        listbox.bind("<Double-Button-1>", on_double_click)

        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=5)

        tk.Button(button_frame, text="確定", font=self.candidate_button_font, command=on_select).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="取消", font=self.candidate_button_font, command=on_close).pack(side=tk.LEFT, padx=5)

        def key_handler(num):
            def handler(e=None):
                if num < len(matches):
                    selected_word = matches[num]
                    if self.preselect_mode.get():
                        current_text = self.entry.get()
                        if current_text and len(matches) > 0:
                            first_word = matches[0]
                            if current_text.endswith(first_word):
                                current_text = current_text[:-len(first_word)]
                        self.entry.delete(0, tk.END)
                        self.entry.insert(0, current_text + selected_word)
                    else:
                        current_text = self.entry.get()
                        self.entry.delete(0, tk.END)
                        self.entry.insert(0, current_text + selected_word)
                on_close()
            return handler

        for i in range(10):
            dialog.bind(str(i), key_handler(i))

        dialog.bind("<Escape>", lambda e: on_close())

        def handle_new_input(event):
            if not self.preselect_mode.get():
                return
            
            char = event.char
            allowed_chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!@#$%^&*()-_=+`~[]{}|;:,.<>?/\\""'
            
            if char and char != ' ' and char in allowed_chars:
                self.close_selection_dialog()
                self.chinese_entry.config(state=tk.NORMAL)
                self.chinese_entry.delete(0, tk.END)
                self.chinese_entry.insert(0, char)
                self.chinese_entry.focus()
                self.clear_candidates()
                return "break"

        dialog.bind("<Key>", handle_new_input)

    def close_selection_dialog(self):
        if self.selection_dialog:
            try:
                if self.selection_dialog.winfo_exists():
                    self.selection_dialog.destroy()
            except tk.TclError:
                pass
            finally:
                self.selection_dialog = None
                try:
                    self.chinese_entry.config(state=tk.NORMAL)
                except tk.TclError:
                    pass

    def select_candidate_by_number(self, num):
        """根據數字選擇候選詞"""
        # 只有在中文模式且候選視窗開啟且有候選項目時才處理候選選擇
        if (self.is_chinese_mode.get() and 
            self.is_candidate_window_open() and 
            self.candidates and 
            0 <= num < len(self.candidates)):
            self.select_candidate_append(self.candidates[num])
            return
        
        # 如果焦點在中文輸入框且候選視窗未開啟，讓 on_chinese_key_press 處理
        # 其他情況不做任何處理

    def select_candidate_append(self, word):
        current_text = self.entry.get()
        self.entry.delete(0, tk.END)
        self.entry.insert(0, current_text + word)
        self.clear_candidates()
        self.candidates = []

    def clear_candidates(self):
        for widget in self.candidate_frame.winfo_children():
            widget.destroy()
        self.candidates = []

    def handle_letter_input(self, event):
        return

    # *** 修改：重寫此函式以實現快取機制 ***
    def _parse_and_cache_word_tab(self, cache_file):
        """
        (輔助函式) 解析 word.tab，將結果存入 self.word_dictionary，並建立二進位快取檔案。
        這是「慢速路徑」，只在需要時執行。
        """
        temp_dict = {}
        try:
            # 1. 從 word.tab 解析文字
            with open(self.word_tab_file, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    
                    parts = line.split()
                    if len(parts) >= 2:
                        code = parts[0]
                        words = parts[1:]
                        temp_dict[code] = words
            
            self.word_dictionary = temp_dict

            # 2. 將新建立的字典寫入快取檔案
            with open(cache_file, "wb") as f_cache:
                pickle.dump(self.word_dictionary, f_cache)
                print("詞庫快取已成功建立/更新。")

        except Exception as e:
            messagebox.showerror("錯誤", f"讀取 word.tab 或建立快取失敗: {e}")
            self.word_dictionary = {} # 確保在失敗時字典是空的

    def load_word_tab(self):
        """
        (重構後) 載入 word.tab 檔案。
        優先使用二進位快取以加速啟動，僅在原始檔更新或快取不存在時才重新解析。
        """
        self.word_dictionary = {}
        cache_file = self.word_tab_file + ".cache" # 快取檔案名稱

        # 情境一：word.tab 檔案不存在，創建範例檔和初始快取
        if not os.path.exists(self.word_tab_file):
            sample_data = {
                "AA": ["寸", "尺", "分"], "BB": ["公分", "公尺"], "CC": ["很好", "不錯", "棒"],
                "aaa": ["鑫", "龘", "鑆"], "DD": ["測試"], "ABC": ["第一", "第二", "第三", "第四"],
                "ABCD": ["甲", "乙", "丙", "丁"], "jou": ["捐", "胡"], "jouv": ["測試V"]
            }
            try:
                # 寫入人類可讀的 word.tab
                with open(self.word_tab_file, "w", encoding="utf-8") as f:
                    for code, words in sample_data.items():
                        f.write(f"{code} {' '.join(words)}\n")
                
                # 寫入二進位快取檔
                with open(cache_file, "wb") as f_cache:
                    pickle.dump(sample_data, f_cache)

                self.word_dictionary = sample_data
                messagebox.showinfo("提示", "已創建範例 word.tab 及快取檔案")
            except Exception as e:
                messagebox.showerror("錯誤", f"創建範例 word.tab 失敗: {e}")
            return # 完成處理，直接返回

        # 情境二：word.tab 存在，判斷是否使用快取
        use_cache = False
        if os.path.exists(cache_file):
            try:
                # 比較 word.tab 和快取檔案的最後修改時間
                word_tab_mtime = os.path.getmtime(self.word_tab_file)
                cache_mtime = os.path.getmtime(cache_file)
                if cache_mtime > word_tab_mtime:
                    use_cache = True
            except OSError:
                use_cache = False # 如果無法獲取時間戳，則不使用快取

        if use_cache:
            # --- 快速路徑：從快取載入 ---
            print("偵測到有效快取，正在從快取載入詞庫...")
            try:
                with open(cache_file, "rb") as f:
                    self.word_dictionary = pickle.load(f)
            except Exception as e:
                # 如果快取檔案損毀或讀取失敗，則退回到慢速路徑
                print(f"快取讀取失敗: {e}。將從 word.tab 重新解析。")
                self._parse_and_cache_word_tab(cache_file)
        else:
            # --- 慢速路徑：從 word.tab 解析並建立快取 ---
            print("快取無效或不存在，正在從 word.tab 解析詞庫...")
            self._parse_and_cache_word_tab(cache_file)

    def on_enter(self, event):
        user_input = self.entry.get().strip()
        if not user_input:
            return
        pyperclip.copy(user_input)
        self.add_to_history(user_input)
        self.entry.delete(0, tk.END)

    def on_enter_from_chinese(self, event):
        user_input = self.entry.get().strip()
        if not user_input:
            return
        pyperclip.copy(user_input)
        self.add_to_history(user_input)
        self.entry.delete(0, tk.END)
        self.chinese_entry.delete(0, tk.END)

    def add_to_history(self, text):
        if text in self.history:
            return
        self.history.append(text)
        self.history_listbox.insert(tk.END, text)
        self.save_history()

    def clear_entry(self):
        self.entry.delete(0, tk.END)
        if self.is_chinese_mode.get():
            self.chinese_entry.delete(0, tk.END)
            self.clear_candidates()
        self.close_selection_dialog()

    def clear_history(self):
        self.history = []
        self.history_listbox.delete(0, tk.END)
        self.save_history()
        messagebox.showinfo("提示", "歷史紀錄已清除")

    def on_history_select(self, event):
        selected = self.history_listbox.curselection()
        if selected:
            selected_text = self.history_listbox.get(selected)
            pyperclip.copy(selected_text)

    def load_history(self):
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, "r", encoding="utf-8") as f:
                    self.history = json.load(f)
            except Exception as e:
                self.history = []
                messagebox.showerror("錯誤", f"讀取歷史紀錄失敗: {e}")
        else:
            self.history = []

    def save_history(self):
        try:
            with open(self.history_file, "w", encoding="utf-8") as f:
                json.dump(self.history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存歷史紀錄失敗: {e}")

    def on_close(self):
        self.save_settings()  # 儲存設定包含視窗位置
        self.save_history()
        self.close_selection_dialog()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = ClipboardApp(root)
    root.mainloop()
