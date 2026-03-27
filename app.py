import json
import os
import re
import shutil
import sys
import threading
from pathlib import Path

import wx
from cryptography.fernet import Fernet, InvalidToken

# ── 路径工具 ─────────────────────────────────────────────────────────────────

def get_resource_path(filename: str) -> str:
    """获取资源文件绝对路径，兼容 PyInstaller 打包环境。"""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)


def get_app_dir() -> Path:
    """
    获取应用工作目录。
    打包后使用 ~/SplitIt/；开发模式使用当前目录。
    """
    if hasattr(sys, "_MEIPASS"):
        d = Path.home() / "SplitIt"
    else:
        d = Path.cwd()
    d.mkdir(parents=True, exist_ok=True)
    return d


APP_DIR = get_app_dir()
PROMPT_FILE = str(APP_DIR / "QAprompt.txt")
MARKDOWN_DOC_DIR = str(APP_DIR / "markdown_doc")

# ── API Key 加密存储 ──────────────────────────────────────────────────────────

_CONFIG_DIR = Path.home() / ".split-it"
_API_CONFIG_FILE = _CONFIG_DIR / ".api_config.enc"
_KEY_FILE = _CONFIG_DIR / ".fernet_key"

_DEFAULT_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
_DEFAULT_MODEL = "qwen-plus"


def _get_fernet_key() -> bytes:
    """获取或生成持久化 Fernet 密钥（首次运行时随机生成并写入磁盘）。"""
    _CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    if not _KEY_FILE.exists():
        _KEY_FILE.write_bytes(Fernet.generate_key())
    return _KEY_FILE.read_bytes()


def save_api_config(api_key: str, base_url: str, model: str) -> None:
    payload = json.dumps(
        {"api_key": api_key, "base_url": base_url, "model": model}
    ).encode()
    encrypted = Fernet(_get_fernet_key()).encrypt(payload)
    _API_CONFIG_FILE.write_bytes(encrypted)


def load_api_config() -> dict | None:
    if not _API_CONFIG_FILE.exists():
        return None
    try:
        decrypted = Fernet(_get_fernet_key()).decrypt(_API_CONFIG_FILE.read_bytes())
        return json.loads(decrypted.decode())
    except (InvalidToken, Exception):
        return None


# ── 对话框：API 设置 ──────────────────────────────────────────────────────────

class ApiSettingsDialog(wx.Dialog):
    def __init__(self, parent, config: dict | None = None):
        super().__init__(
            parent,
            title="API 设置",
            size=(560, 260),
            style=wx.DEFAULT_DIALOG_STYLE,
        )
        cfg = config or {}
        self._init_ui(
            cfg.get("api_key", ""),
            cfg.get("base_url", _DEFAULT_BASE_URL),
            cfg.get("model", _DEFAULT_MODEL),
        )
        self.Centre()

    def _init_ui(self, api_key: str, base_url: str, model: str):
        panel = wx.Panel(self)
        grid = wx.FlexGridSizer(rows=3, cols=2, vgap=10, hgap=8)
        grid.AddGrowableCol(1, 1)

        grid.Add(wx.StaticText(panel, label="API Key："), 0, wx.ALIGN_CENTER_VERTICAL)
        self.key_ctrl = wx.TextCtrl(panel, value=api_key, style=wx.TE_PASSWORD)
        grid.Add(self.key_ctrl, 1, wx.EXPAND)

        grid.Add(wx.StaticText(panel, label="Base URL："), 0, wx.ALIGN_CENTER_VERTICAL)
        self.url_ctrl = wx.TextCtrl(panel, value=base_url)
        grid.Add(self.url_ctrl, 1, wx.EXPAND)

        grid.Add(wx.StaticText(panel, label="模型名称："), 0, wx.ALIGN_CENTER_VERTICAL)
        self.model_ctrl = wx.TextCtrl(panel, value=model)
        grid.Add(self.model_ctrl, 1, wx.EXPAND)

        btn_sizer = wx.StdDialogButtonSizer()
        save_btn = wx.Button(panel, wx.ID_OK, "保存")
        cancel_btn = wx.Button(panel, wx.ID_CANCEL, "取消")
        btn_sizer.AddButton(save_btn)
        btn_sizer.AddButton(cancel_btn)
        btn_sizer.Realize()

        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.Add(grid, 0, wx.EXPAND | wx.ALL, 12)
        vbox.Add(btn_sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 12)
        panel.SetSizer(vbox)

        save_btn.Bind(wx.EVT_BUTTON, self._on_save)

    def _on_save(self, event):
        if not self.key_ctrl.GetValue().strip():
            wx.MessageBox("API Key 不能为空！", "验证失败", wx.OK | wx.ICON_ERROR, self)
            return
        if not self.url_ctrl.GetValue().strip():
            wx.MessageBox("Base URL 不能为空！", "验证失败", wx.OK | wx.ICON_ERROR, self)
            return
        event.Skip()

    def get_config(self) -> dict:
        return {
            "api_key": self.key_ctrl.GetValue().strip(),
            "base_url": self.url_ctrl.GetValue().strip(),
            "model": self.model_ctrl.GetValue().strip() or _DEFAULT_MODEL,
        }


# ── 对话框：提示词设置 ────────────────────────────────────────────────────────

class PromptDialog(wx.Dialog):
    def __init__(self, parent, prompt_text: str = ""):
        super().__init__(
            parent,
            title="设置提示词",
            size=(680, 540),
            style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
        )
        self._init_ui(prompt_text)
        self.Centre()

    def _init_ui(self, prompt_text: str):
        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        hint = wx.StaticText(panel, label="提示词内容（必须包含 $DOC_CONTENT$ 占位符）：")
        vbox.Add(hint, 0, wx.ALL, 8)

        self.text_ctrl = wx.TextCtrl(panel, style=wx.TE_MULTILINE | wx.TE_PROCESS_TAB)
        self.text_ctrl.SetValue(prompt_text)
        self.text_ctrl.SetFont(
            wx.Font(11, wx.FONTFAMILY_TELETYPE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        )
        vbox.Add(self.text_ctrl, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

        btn_sizer = wx.StdDialogButtonSizer()
        save_btn = wx.Button(panel, wx.ID_OK, "保存")
        cancel_btn = wx.Button(panel, wx.ID_CANCEL, "取消")
        btn_sizer.AddButton(save_btn)
        btn_sizer.AddButton(cancel_btn)
        btn_sizer.Realize()
        vbox.Add(btn_sizer, 0, wx.EXPAND | wx.ALL, 8)

        panel.SetSizer(vbox)
        save_btn.Bind(wx.EVT_BUTTON, self._on_save)

    def _on_save(self, event):
        if "$DOC_CONTENT$" not in self.text_ctrl.GetValue():
            wx.MessageBox(
                "提示词中必须包含 $DOC_CONTENT$ 占位符！",
                "验证失败",
                wx.OK | wx.ICON_ERROR,
                self,
            )
            return
        event.Skip()

    def get_prompt(self) -> str:
        return self.text_ctrl.GetValue()


# ── 对话框：拆分进度 ──────────────────────────────────────────────────────────

class ProgressDialog(wx.Dialog):
    def __init__(self, parent):
        super().__init__(
            parent,
            title="拆分进度",
            size=(700, 500),
            style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
        )
        self._done = False
        self.stop_event = threading.Event()
        self._init_ui()
        self.Centre()

    def _init_ui(self):
        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        self.log_ctrl = wx.TextCtrl(
            panel,
            style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2 | wx.HSCROLL,
        )
        self.log_ctrl.SetFont(
            wx.Font(11, wx.FONTFAMILY_TELETYPE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        )
        vbox.Add(self.log_ctrl, 1, wx.EXPAND | wx.ALL, 8)

        self.gauge = wx.Gauge(panel, range=100, style=wx.GA_HORIZONTAL | wx.GA_SMOOTH)
        vbox.Add(self.gauge, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

        self.status_label = wx.StaticText(panel, label="处理中，请稍候...")
        vbox.Add(self.status_label, 0, wx.LEFT | wx.BOTTOM, 8)

        btn_hbox = wx.BoxSizer(wx.HORIZONTAL)
        self.abort_btn = wx.Button(panel, label="立即终止")
        self.abort_btn.SetForegroundColour(wx.Colour(180, 0, 0))
        self.close_btn = wx.Button(panel, wx.ID_OK, "关闭")
        self.close_btn.Enable(False)
        btn_hbox.Add(self.abort_btn, 0, wx.RIGHT, 8)
        btn_hbox.Add(self.close_btn, 0)
        vbox.Add(btn_hbox, 0, wx.ALIGN_RIGHT | wx.ALL, 8)

        panel.SetSizer(vbox)
        self.abort_btn.Bind(wx.EVT_BUTTON, self._on_abort)
        self.Bind(wx.EVT_CLOSE, self._on_close)

    def _on_abort(self, event):
        self.stop_event.set()
        self.abort_btn.Enable(False)
        self.abort_btn.SetLabel("正在终止...")
        self.status_label.SetLabel("收到终止指令，等待当前块处理完毕...")

    def _on_close(self, event):
        if self._done:
            event.Skip()

    def append_log(self, text: str):
        wx.CallAfter(self._do_append_log, text)

    def _do_append_log(self, text: str):
        self.log_ctrl.AppendText(text + "\n")

    def set_progress(self, value: int, status: str = ""):
        wx.CallAfter(self._do_set_progress, value, status)

    def _do_set_progress(self, value: int, status: str):
        self.gauge.SetValue(min(value, 100))
        if status:
            self.status_label.SetLabel(status)

    def set_done(self, success: bool = True, cancelled: bool = False):
        wx.CallAfter(self._do_set_done, success, cancelled)

    def _do_set_done(self, success: bool, cancelled: bool):
        self._done = True
        self.abort_btn.Enable(False)
        self.abort_btn.SetLabel("立即终止")
        self.close_btn.Enable(True)
        if cancelled:
            self.status_label.SetLabel("已终止。")
        elif success:
            self.status_label.SetLabel("全部完成！")
        else:
            self.status_label.SetLabel("处理过程中出现错误，请查看日志。")


# ── 主窗口 ────────────────────────────────────────────────────────────────────

class MainFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="Split-It 文档拆分工具", size=(800, 270))
        self._set_icon()
        self._init_ui()
        self.Centre()
        self.SetMinSize((800, 270))
        self.SetMaxSize((800, 270))

    def _set_icon(self):
        icon_path = get_resource_path("split-it.ico")
        if os.path.exists(icon_path):
            icon = wx.Icon()
            icon.LoadFile(icon_path, wx.BITMAP_TYPE_ICO)
            self.SetIcon(icon)

    def _init_ui(self):
        panel = wx.Panel(self)
        main_vbox = wx.BoxSizer(wx.VERTICAL)

        # ── 文件选择行 ────────────────────────────────────────────
        file_box = wx.StaticBox(panel, label="Word 文件（仅支持 .docx）")
        file_sizer = wx.StaticBoxSizer(file_box, wx.HORIZONTAL)
        self.file_path_ctrl = wx.TextCtrl(panel, style=wx.TE_READONLY)
        browse_btn = wx.Button(panel, label="浏览...")
        file_sizer.Add(self.file_path_ctrl, 1, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 4)
        file_sizer.Add(browse_btn, 0, wx.ALL, 4)
        main_vbox.Add(file_sizer, 0, wx.EXPAND | wx.ALL, 8)

        # ── 操作 + 设置 ──────────────────────────────────────────
        mid_hbox = wx.BoxSizer(wx.HORIZONTAL)

        action_box = wx.StaticBox(panel, label="操作")
        action_sizer = wx.StaticBoxSizer(action_box, wx.VERTICAL)
        self.split_btn = wx.Button(panel, label="拆分文档 QA Pairs")
        self.prompt_btn = wx.Button(panel, label="设置提示词")
        self.api_btn = wx.Button(panel, label="API 设置")
        action_sizer.Add(self.split_btn, 0, wx.EXPAND | wx.ALL, 4)
        action_sizer.Add(self.prompt_btn, 0, wx.EXPAND | wx.ALL, 4)
        action_sizer.Add(self.api_btn, 0, wx.EXPAND | wx.ALL, 4)
        mid_hbox.Add(action_sizer, 0, wx.EXPAND | wx.RIGHT, 8)

        settings_box = wx.StaticBox(panel, label="设置")
        settings_sizer = wx.StaticBoxSizer(settings_box, wx.VERTICAL)

        depth_hbox = wx.BoxSizer(wx.HORIZONTAL)
        depth_hbox.Add(
            wx.StaticText(panel, label="拆分深度（级）："),
            0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 4,
        )
        self.depth_spin = wx.SpinCtrl(panel, value="3", min=1, max=6, size=(65, -1))
        depth_hbox.Add(self.depth_spin, 0)
        settings_sizer.Add(depth_hbox, 0, wx.ALL, 4)

        test_hbox = wx.BoxSizer(wx.HORIZONTAL)
        self.test_mode_cb = wx.CheckBox(panel, label="测试模式，只处理前")
        self.test_n_spin = wx.SpinCtrl(panel, value="3", min=1, max=999, size=(65, -1))
        self.test_n_spin.Enable(False)
        test_hbox.Add(self.test_mode_cb, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 4)
        test_hbox.Add(self.test_n_spin, 0, wx.RIGHT, 4)
        test_hbox.Add(wx.StaticText(panel, label="个块"), 0, wx.ALIGN_CENTER_VERTICAL)
        settings_sizer.Add(test_hbox, 0, wx.ALL, 4)

        self.overwrite_cb = wx.CheckBox(panel, label="覆盖已存在的 qa_pairs_xx.md")
        self.overwrite_cb.SetValue(True)
        settings_sizer.Add(self.overwrite_cb, 0, wx.ALL, 4)

        mid_hbox.Add(settings_sizer, 1, wx.EXPAND)
        main_vbox.Add(mid_hbox, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

        # ── 底部：打开目录 ────────────────────────────────────────
        open_btn = wx.Button(panel, label="打开输出目录")
        main_vbox.Add(open_btn, 0, wx.ALIGN_RIGHT | wx.RIGHT | wx.BOTTOM, 8)

        panel.SetSizer(main_vbox)

        browse_btn.Bind(wx.EVT_BUTTON, self.on_browse)
        self.split_btn.Bind(wx.EVT_BUTTON, self.on_split)
        self.prompt_btn.Bind(wx.EVT_BUTTON, self.on_set_prompt)
        self.api_btn.Bind(wx.EVT_BUTTON, self.on_api_settings)
        open_btn.Bind(wx.EVT_BUTTON, self.on_open_folder)
        self.test_mode_cb.Bind(wx.EVT_CHECKBOX, self._on_test_mode_toggle)

    # ── 事件 ──────────────────────────────────────────────────────

    def _on_test_mode_toggle(self, event):
        self.test_n_spin.Enable(self.test_mode_cb.IsChecked())

    def on_browse(self, event):
        with wx.FileDialog(
            self,
            "选择 Word 文档",
            wildcard="Word 文档 (*.docx)|*.docx",
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
        ) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                self.file_path_ctrl.SetValue(dlg.GetPath())

    def on_api_settings(self, event):
        cfg = load_api_config()
        with ApiSettingsDialog(self, cfg) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                new_cfg = dlg.get_config()
                save_api_config(
                    new_cfg["api_key"], new_cfg["base_url"], new_cfg["model"]
                )
                wx.MessageBox(
                    "API 配置已加密保存！", "保存成功", wx.OK | wx.ICON_INFORMATION, self
                )

    def on_set_prompt(self, event):
        prompt_text = ""
        if os.path.exists(PROMPT_FILE):
            with open(PROMPT_FILE, "r", encoding="utf-8") as f:
                prompt_text = f.read()

        with PromptDialog(self, prompt_text) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                with open(PROMPT_FILE, "w", encoding="utf-8") as f:
                    f.write(dlg.get_prompt())
                wx.MessageBox(
                    "提示词已保存！", "保存成功", wx.OK | wx.ICON_INFORMATION, self
                )

    def on_open_folder(self, event):
        folder = MARKDOWN_DOC_DIR
        os.makedirs(folder, exist_ok=True)
        if sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", folder])
        elif sys.platform == "win32":
            os.startfile(folder)
        else:
            import subprocess
            subprocess.Popen(["xdg-open", folder])

    def on_split(self, event):
        docx_path = self.file_path_ctrl.GetValue().strip()
        if not docx_path:
            wx.MessageBox("请先选择一个 Word 文件！", "提示", wx.OK | wx.ICON_WARNING, self)
            return
        if not os.path.exists(docx_path):
            wx.MessageBox("所选文件不存在！", "错误", wx.OK | wx.ICON_ERROR, self)
            return

        api_cfg = load_api_config()
        if not api_cfg or not api_cfg.get("api_key"):
            if wx.MessageBox(
                "尚未配置 API Key，是否现在设置？",
                "提示",
                wx.YES_NO | wx.ICON_QUESTION,
                self,
            ) == wx.YES:
                self.on_api_settings(None)
                api_cfg = load_api_config()
            if not api_cfg or not api_cfg.get("api_key"):
                return

        if not os.path.exists(PROMPT_FILE):
            if wx.MessageBox(
                "提示词文件不存在，是否现在设置提示词？",
                "提示",
                wx.YES_NO | wx.ICON_QUESTION,
                self,
            ) == wx.YES:
                self.on_set_prompt(None)
            return

        depth = self.depth_spin.GetValue()
        test_mode = self.test_mode_cb.IsChecked()
        test_n = self.test_n_spin.GetValue() if test_mode else None
        overwrite = self.overwrite_cb.IsChecked()

        self.split_btn.Enable(False)
        progress_dlg = ProgressDialog(self)

        threading.Thread(
            target=self._do_split_thread,
            args=(docx_path, depth, test_n, overwrite, api_cfg,
                  progress_dlg, progress_dlg.stop_event),
            daemon=True,
        ).start()

        progress_dlg.ShowModal()
        progress_dlg.Destroy()
        self.split_btn.Enable(True)

    # ── 后台线程 ───────────────────────────────────────────────────

    def _do_split_thread(
        self,
        docx_path: str,
        depth: int,
        test_n: int | None,
        overwrite: bool,
        api_cfg: dict,
        dlg: ProgressDialog,
        stop_event: threading.Event,
    ):
        success = False
        cancelled = False
        try:
            from main import docx_to_markdown
            from paser import markdown_split_by_level
            import openai

            output_dir = MARKDOWN_DOC_DIR
            os.makedirs(output_dir, exist_ok=True)

            # 步骤 1：复制原文件
            dlg.append_log("【步骤 1/4】复制文件到 markdown_doc 目录...")
            dlg.set_progress(5, "复制文件...")
            dest_docx = os.path.join(output_dir, Path(docx_path).name)
            shutil.copy2(docx_path, dest_docx)
            dlg.append_log(f"  ✓ 已复制：{Path(dest_docx).name}")

            # 步骤 2：DOCX → Markdown
            dlg.append_log("【步骤 2/4】将 DOCX 转换为 Markdown...")
            dlg.set_progress(10, "转换 DOCX→Markdown...")
            output_md = os.path.join(output_dir, "_output.md")
            md_text = docx_to_markdown(dest_docx, output_dir=output_dir)
            with open(output_md, "w", encoding="utf-8") as f:
                f.write(md_text)
            dlg.append_log("  ✓ 已输出：_output.md（图片目录：_images）")
            dlg.set_progress(25, "DOCX 转换完成")

            # 步骤 3：切分 + LLM 提取 QA Pairs
            dlg.append_log("【步骤 3/4】提取 QA Pairs（调用 LLM）...")
            with open(output_md, "r", encoding="utf-8") as f:
                md_content = f.read()
            blocks = markdown_split_by_level(md_content, depth=depth)

            with open(PROMPT_FILE, "r", encoding="utf-8") as f:
                prompt_template = f.read()

            if test_n is not None:
                blocks = blocks[:test_n]
                dlg.append_log(f"  [测试模式] 只处理前 {len(blocks)} 个块")

            total = len(blocks)
            dlg.append_log(f"  共 {total} 个块需要处理")

            llm_client = openai.OpenAI(
                api_key=api_cfg["api_key"],
                base_url=api_cfg.get("base_url", _DEFAULT_BASE_URL),
            )
            model = api_cfg.get("model", _DEFAULT_MODEL)

            qa_files: list[str] = []
            for i, block in enumerate(blocks):
                first_line = block.to_markdown().split("\n")[0][:60]
                qa_file = os.path.join(output_dir, f"qa_pairs_{i}.md")
                dlg.set_progress(25 + int((i + 1) / total * 65), f"处理第 {i+1}/{total} 个块...")
                print(f"qa_file: {qa_file}, overwrite: {overwrite}")
                if stop_event.is_set():
                    dlg.append_log(f"\n⚠️  收到终止指令，已在第 {i + 1}/{total} 块处中断。")
                    cancelled = True
                    break

                if not overwrite and os.path.exists(qa_file):
                    dlg.append_log(f"  [{i + 1}/{total}] 跳过（文件已存在）：qa_pairs_{i}.md")
                    qa_files.append(qa_file)
                    continue

                dlg.append_log(f"  [{i + 1}/{total}] {first_line}...")

                prompt = prompt_template.replace(
                    "$DOC_CONTENT$", block.to_markdown(include_parent=True)
                )
                completion = llm_client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                )
                response = completion.choices[0].message.content
                response = re.search(r"```markdown(.*?)\n```", response, re.DOTALL).group(1)
                with open(qa_file, "w", encoding="utf-8") as f:
                    f.write(response)
                qa_files.append(qa_file)

            # 步骤 4：合并子文档
            dlg.append_log("【步骤 4/5】合并 QA Pairs 子文档...")
            dlg.set_progress(90, "合并文件...")
            merged_file = os.path.join(output_dir, "qa_pairs.md")
            with open(merged_file, "w", encoding="utf-8") as out_f:
                for qa_file in qa_files:
                    with open(qa_file, "r", encoding="utf-8") as in_f:
                        out_f.write(in_f.read())
                    out_f.write("\n\n")
            dlg.append_log("  ✓ 已合并到：qa_pairs.md")

            # 步骤 5：Markdown → DOCX
            if not cancelled:
                dlg.append_log("【步骤 5/5】将 qa_pairs.md 转换为 DOCX（含图片）...")
                dlg.set_progress(94, "转换 Markdown→DOCX...")
                docx_output = os.path.join(output_dir, "qa_pairs.docx")
                from paser import markdown_to_docx
                with open(merged_file, "r", encoding="utf-8") as f:
                    merged_content = f.read()
                docx_bytes = markdown_to_docx(merged_content, image_dir=output_dir)
                with open(docx_output, "wb") as f:
                    f.write(docx_bytes)
                dlg.append_log("  ✓ 已输出：qa_pairs.docx")
                dlg.set_progress(100, "全部完成！")
                dlg.append_log(f"\n🎉 全部完成！共生成 {total} 个块的 QA Pairs。")
                success = True

        except Exception as exc:
            import traceback
            dlg.append_log(f"\n❌ 错误：{exc}")
            dlg.append_log(traceback.format_exc())
        finally:
            dlg.set_done(success, cancelled=cancelled)


# ── 入口 ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = wx.App(False)
    frame = MainFrame()
    frame.Show()
    app.MainLoop()
