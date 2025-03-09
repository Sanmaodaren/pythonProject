import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import threading
import queue
import xlrd
import xlwt
import re
import itertools

# 常量集中管理
class Constants:
    WEEKS_RANGE = range(1, 21)
    WEEKDAYS = ['周一', '周二', '周三', '周四', '周五']
    PERIODS = ['1-2', '3-4', '5-6', '7-8']
    COLUMN_COUNT = 6
    COLUMN_WIDTH = 256 * 15  # 15字符宽度
    TITLE_ROW_HEIGHT = 500  # 标题行高度
    START_ROW = 2  # 数据起始行


class StyleFactory:
    @staticmethod
    def create_base_style():
        style = xlwt.XFStyle()
        alignment = xlwt.Alignment()
        alignment.horz = xlwt.Alignment.HORZ_CENTER
        alignment.vert = xlwt.Alignment.VERT_CENTER
        style.alignment = alignment
        return style

    @staticmethod
    def create_title_style():
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.bold = True
        style.font = font
        alignment = xlwt.Alignment()
        alignment.horz = xlwt.Alignment.HORZ_CENTER
        alignment.vert = xlwt.Alignment.VERT_CENTER
        style.alignment = alignment
        return style


class ExcelGenerator:
    def __init__(self, config, queue):
        self.config = config
        self.queue = queue
        self.cell_style = StyleFactory.create_base_style()
        self.title_style = StyleFactory.create_title_style()

    def _process_merged_cells(self, sheet, row, col):
        """处理合并单元格逻辑"""
        for (r1, r2, c1, c2) in sheet.merged_cells:
            if r1 <= row < r2 and c1 <= col < c2:
                return sheet.cell_value(r1, c1)
        return sheet.cell_value(row, col)

    def get_teacher_name(self, filename):
        """从Excel文件中提取教师姓名"""
        try:
            with xlrd.open_workbook(filename, formatting_info=True) as workbook:
                sheet = workbook.sheet_by_index(0)
                raw_str = self._process_merged_cells(sheet, 2, 0)
                if match := re.search(r'学生姓名[:：]\s*([^\s]+)', raw_str):
                    return match.group(1).strip()
                return "未知姓名"
        except Exception as e:
            self.queue.put(("error", f"读取学生姓名失败: {str(e)}"))
            return "读取失败"

    def _parse_week_segment(self, seg):
        """解析单周区间段"""
        if single_double := re.match(r'^(\d+)-(\d+)([单双])$', seg):
            start, end, flag = map(single_double.group, [1, 2, 3])
            start, end = int(start), int(end)
            step = 2
            start += (flag == '双' and start % 2 != 0) or (flag == '单' and start % 2 == 0)
            return range(start, end + 1, step)
        elif '-' in seg:
            start, end = map(int, seg.split('-'))
            return range(start, end + 1)
        elif seg.isdigit():
            return [int(seg)]
        return []

    def analyze_schedule(self, cell_value):
        """解析课时安排数据"""
        weeks = set()
        if week_info := re.search(r'n\(([^ )]+)', str(cell_value)):
            for seg in week_info.group(1).split(','):
                weeks.update(self._parse_week_segment(seg.strip()))
        return weeks

    def _setup_sheet_columns(self, sheet):
        """初始化表格列样式"""
        for col in range(Constants.COLUMN_COUNT):
            sheet.col(col).width = Constants.COLUMN_WIDTH
            sheet.col(col).set_style(self.cell_style)

    def create_template_sheet(self, wb, week_num, student_count):
        """创建单个周次模板"""
        sheet = wb.add_sheet(f'sheet{week_num}')
        self._setup_sheet_columns(sheet)

        sheet.write_merge(0, 0, 0, 5, f'第{week_num}周 无课表', self.title_style)
        sheet.row(0).height = Constants.TITLE_ROW_HEIGHT

        for col, weekday in enumerate(Constants.WEEKDAYS, 1):
            sheet.write(1, col, weekday, self.cell_style)

        for idx, period in enumerate(Constants.PERIODS):
            start_row = Constants.START_ROW + idx * student_count
            sheet.write_merge(
                start_row,
                start_row + student_count - 1,
                0, 0, period, self.cell_style
            )
        return sheet

    def generate_template(self, wb, student_count):
        """生成全部周次模板"""
        for week in Constants.WEEKS_RANGE:
            self.create_template_sheet(wb, week, student_count)

    def process_teacher_schedule(self, wb, file_path, teacher_name, student_index, student_count):
        try:
            with xlrd.open_workbook(file_path, formatting_info=True) as workbook:
                sheet = workbook.sheet_by_index(0)
                schedule_data = {
                    col: {
                        period_idx: self.analyze_schedule(sheet.cell(row, col))
                        for period_idx, row in enumerate(range(4, 11, 2))
                    } for col in range(1, 6)
                }

                for weekday in range(1, 6):
                    for period_idx in range(4):
                        teaching_weeks = schedule_data[weekday].get(period_idx, set())
                        free_weeks = set(Constants.WEEKS_RANGE) - teaching_weeks

                        base_row = Constants.START_ROW + period_idx * student_count
                        target_row = base_row + student_index

                        for week in free_weeks:
                            target_sheet = wb.get_sheet(f'sheet{week}')
                            target_sheet.write(target_row, weekday, teacher_name, self.cell_style)
        except Exception as e:
            self.queue.put(("error", f"处理文件失败: {os.path.basename(file_path)}\n{str(e)}"))

    def execute_generation(self):
        """执行生成流程"""
        try:
            workbook = xlwt.Workbook(encoding='utf-8')
            student_count = len(self.config['teachers'])
            self.generate_template(workbook, student_count)

            for idx, teacher in enumerate(self.config['teachers']):
                if not os.path.exists(teacher['file']):
                    self.queue.put(("error", f"文件不存在: {os.path.basename(teacher['file'])}"))
                    continue

                name = self.get_teacher_name(teacher['file'])
                self.process_teacher_schedule(workbook, teacher['file'], name, idx, student_count)
                self.queue.put(("progress", (idx + 1) / student_count * 100))

            workbook.save(self.config['output_file'])
            self.queue.put(("success", "无课表生成成功！"))
        except Exception as e:
            self.queue.put(("error", f"生成失败: {str(e)}"))


class ApplicationGUI:
    def __init__(self, master):
        self.master = master
        master.title("无课表生成器 v1.0")
        master.geometry("800x700")
        self._init_ui()
        self._load_config()
        self._setup_event_loop()

    def _init_ui(self):
        """初始化用户界面"""
        self._create_output_section()
        self._create_teacher_list()
        self._create_controls()
        self._create_progress_bar()
        self._create_notes_section()

    def _create_notes_section(self):
        """创建注意事项模块"""
        notes_frame = ttk.LabelFrame(self.master, text="注意事项")
        notes_frame.pack(pady=10, padx=10, fill="x", expand=False)

        notes_text = (
            "1. 可添加任意数量的excel文件\n"
            "2. 仅支持.xls格式的2003版Excel文件"
        )

        note_label = ttk.Label(notes_frame, text=notes_text, wraplength=750, justify="left")
        note_label.pack(padx=5, pady=5, fill="x")

    def _create_output_section(self):
        """创建输出路径组件"""
        frame = ttk.LabelFrame(self.master, text="输出设置")
        frame.pack(pady=10, padx=10, fill="x")
        self.output_entry = ttk.Entry(frame, width=50)
        self.output_entry.pack(side="left", padx=5)
        ttk.Button(frame, text="浏览...", command=self._select_output).pack(side="left")

    def _create_teacher_list(self):
        """创建学生列表组件"""
        frame = ttk.LabelFrame(self.master, text="学生列表")
        frame.pack(pady=10, padx=10, fill="both", expand=True)

        self.teacher_tree = ttk.Treeview(frame, columns=("file", "offset"), show="headings")
        self.teacher_tree.heading("file", text="Excel文件")
        self.teacher_tree.heading("offset", text="行偏移量")
        self.teacher_tree.column("file", width=500)
        self.teacher_tree.column("offset", width=100)
        self.teacher_tree.pack(fill="both", expand=True)

    def _create_controls(self):
        """创建操作按钮"""
        frame = ttk.Frame(self.master)
        frame.pack(pady=10, fill="x")

        controls = [
            ("添加学生", self._add_teacher),
            ("删除选中", self._remove_teacher),
            ("生成无课表", self._start_generation)]

        for text, cmd in controls[:2]:
            ttk.Button(frame, text=text, command=cmd).pack(side="left", padx=5)

        ttk.Button(frame, text=controls[2][0], command=controls[2][1]).pack(side="right", padx=5)

    def _create_progress_bar(self):
        """创建进度条"""
        self.progress = ttk.Progressbar(self.master, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=5)

    def _setup_event_loop(self):
        """启动事件处理循环"""
        self.running = False
        self.queue = queue.Queue()
        self.master.after(100, self._process_messages)

    def _select_output(self):
        """选择输出文件路径"""
        if path := filedialog.asksaveasfilename(
                defaultextension=".xls",
                filetypes=[("Excel文件", "*.xls")]
        ):
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, path)

    def _validate_inputs(self):
        """输入验证"""
        errors = []
        if not self.output_entry.get():
            errors.append("请设置输出文件路径")
        if not self.teacher_tree.get_children():
            errors.append("请至少添加一个学生")
        return errors

    def _save_config(self):
        """保存配置"""
        self.config = {
            "output_file": self.output_entry.get(),
            "students": [
                {"file": self.teacher_tree.item(item, "values")[0]}
                for item in self.teacher_tree.get_children()
            ]
        }

    def _load_config(self):
        """加载配置"""
        try:
            with open("config.json", "r") as f:
                config = json.load(f)
                self.output_entry.delete(0, tk.END)
                self.output_entry.insert(0, config["output_file"])

                for item in list(self.teacher_tree.get_children()):
                    self.teacher_tree.delete(item)

                for teacher in config["students"]:
                    self.teacher_tree.insert("", "end", values=(
                        teacher["file"],
                        "自动分配"
                    ))
        except Exception as e:
            print(f"配置加载异常: {str(e)}")

    def _process_messages(self):
        """处理线程消息"""
        try:
            while True:
                msg_type, content = self.queue.get_nowait()
                if msg_type == "progress":
                    self.progress["value"] = content
                elif msg_type == "success":
                    messagebox.showinfo("成功", content)
                    self.running = False
                    self._toggle_controls(True)
                elif msg_type == "error":
                    messagebox.showerror("错误", content)
                    self.running = False
                    self._toggle_controls(True)
        except queue.Empty:
            pass
        finally:
            self.master.after(100, self._process_messages)

    def _add_teacher(self):
        """添加学生"""
        paths = filedialog.askopenfilenames(
            title="选择学生课表文件",
            filetypes=[("Excel文件", "*.xls")]
        )
        for path in paths:
            existing = [self.teacher_tree.item(i, "values")[0]
                        for i in self.teacher_tree.get_children()]
            if path not in existing:
                self.teacher_tree.insert("", "end", values=(path, "自动分配"))

    def _remove_teacher(self):
        """移除选中学生"""
        for item in self.teacher_tree.selection():
            self.teacher_tree.delete(item)

    def _toggle_controls(self, state):
        """切换控件状态"""
        state_str = "normal" if state else "disabled"
        for child in self.master.winfo_children():
            if isinstance(child, ttk.Button):
                child["state"] = state_str

    def _start_generation(self):
        """启动生成流程"""
        errors = self._validate_inputs()
        if errors:
            messagebox.showerror("错误", "\n".join(errors))
            return

        self._toggle_controls(False)
        self.progress["value"] = 0

        config = {
            "output_file": self.output_entry.get(),
            "teachers": [
                {
                    "file": self.teacher_tree.item(item, "values")[0],
                    "offset": idx
                }
                for idx, item in enumerate(self.teacher_tree.get_children())
            ]
        }

        generator = ExcelGenerator(config, self.queue)
        thread = threading.Thread(target=generator.execute_generation)
        thread.start()


if __name__ == "__main__":
    root = tk.Tk()  # 使用标准Tk实例
    app = ApplicationGUI(root)
    root.mainloop()