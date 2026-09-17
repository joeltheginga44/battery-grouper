import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import math
import re
import csv
import os
import json
import datetime

try:
    from PIL import Image, ImageDraw, ImageFont, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


HISTORY_FILE = "build_history.json"


class BatteryGrouperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Universal Battery Pack Grouper")
        self.root.geometry("980x900")

        self.last_groups = None
        self.last_discarded = None
        self.last_config = None
        self.last_raw_data = None
        self.current_map_image = None
        self.current_map_photo = None

        # --- Configuration Frame ---
        config_frame = tk.LabelFrame(root, text=" 1. Enter Pack Configuration ", padx=10, pady=10)
        config_frame.pack(fill="x", padx=10, pady=5)

        tk.Label(config_frame, text="Cells in Series (S):").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.series_entry = tk.Entry(config_frame, width=10)
        self.series_entry.insert(0, "13")
        self.series_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(config_frame, text="Cells in Parallel (P):").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.parallel_entry = tk.Entry(config_frame, width=10)
        self.parallel_entry.insert(0, "5")
        self.parallel_entry.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(config_frame, text="Min Capacity (mAh) [Optional]:").grid(row=0, column=2, sticky="e", padx=5, pady=5)
        self.min_cap_entry = tk.Entry(config_frame, width=10)
        self.min_cap_entry.insert(0, "2000")
        self.min_cap_entry.grid(row=0, column=3, padx=5, pady=5)

        tk.Label(config_frame, text="Max IR (mOhm) [Optional]:").grid(row=1, column=2, sticky="e", padx=5, pady=5)
        self.max_ir_entry = tk.Entry(config_frame, width=10)
        self.max_ir_entry.insert(0, "80")
        self.max_ir_entry.grid(row=1, column=3, padx=5, pady=5)

        tk.Label(config_frame, text="Min Voltage (V) [Optional]:").grid(row=2, column=2, sticky="e", padx=5, pady=5)
        self.min_volt_entry = tk.Entry(config_frame, width=10)
        self.min_volt_entry.insert(0, "3.0")
        self.min_volt_entry.grid(row=2, column=3, padx=5, pady=5)

        tk.Label(config_frame, text="(for voltage safety check)", fg="gray").grid(row=2, column=4, sticky="w", padx=5)

        # --- Data Input Frame ---
        data_frame = tk.LabelFrame(root, text=" 2. Cell Data ", padx=10, pady=10)
        data_frame.pack(fill="x", padx=10, pady=5)

        top_row = tk.Frame(data_frame)
        top_row.pack(fill="x", pady=(0, 5))

        tk.Label(top_row, text="Format: (ID, Capacity, IR) or (ID, Capacity, IR, Voltage)").pack(side="left")

        self.import_btn = tk.Button(top_row, text="📂 Import CSV", command=self.import_csv, bg="#2196F3", fg="white")
        self.import_btn.pack(side="right", padx=5)

        self.clear_btn = tk.Button(top_row, text="🗑️ Clear", command=self.clear_data, bg="#9E9E9E", fg="white")
        self.clear_btn.pack(side="right", padx=5)

        self.data_text = scrolledtext.ScrolledText(data_frame, height=8)
        self.data_text.pack(fill="x", pady=5)

        sample = """(1, 2742, 48), (2, 2338, 58), (3, 2825, 50), (4, 2345, 56), (5, 2468, 61),
(6, 2374, 55), (7, 2355, 66), (8, 2513, 48), (9, 2386, 59), (10, 2365, 54),
(11, 2477, 57), (12, 2407, 55), (13, 2417, 48), (14, 2289, 56), (15, 2338, 64),
(16, 2811, 56), (17, 2297, 59), (18, 2777, 53), (19, 2374, 105), (20, 2193, 54),
(21, 2760, 54), (22, 2346, 56), (23, 2355, 187), (24, 2327, 54), (25, 1603, 68),
(26, 2399, 53), (27, 2400, 64), (28, 2360, 54), (29, 2385, 71), (30, 2416, 79),
(31, 2500, 50), (32, 2402, 54), (33, 2078, 67), (34, 2475, 69), (35, 2522, 57),
(36, 2728, 52), (37, 2747, 52), (38, 2346, 92), (39, 2858, 56), (40, 2378, 82),
(41, 2263, 51), (42, 2437, 55), (43, 2849, 54), (44, 2426, 55), (45, 2438, 53),
(46, 2872, 51), (47, 2398, 54), (48, 2490, 50), (49, 2060, 57), (50, 2472, 54),
(51, 2358, 55), (52, 2741, 52), (53, 2386, 72), (55, 2767, 51), (56, 2721, 184),
(57, 2736, 56), (58, 2425, 54), (59, 2869, 50), (60, 2727, 52), (61, 2754, 57),
(62, 2425, 53), (63, 2339, 54), (64, 2382, 60), (65, 2384, 52), (66, 2371, 81),
(67, 2379, 52), (68, 1859, 56), (69, 2369, 54), (70, 2453, 52), (71, 2342, 56),
(72, 2769, 61), (73, 2996, 51), (74, 2861, 60), (75, 2539, 50), (76, 2387, 55),
(77, 2374, 63), (78, 2400, 54), (79, 2488, 55), (80, 2848, 52), (81, 2507, 52),
(82, 2790, 49), (83, 2377, 60), (84, 2787, 52), (85, 2371, 57), (86, 2378, 46),
(87, 2505, 51), (88, 2500, 51), (89, 2414, 50), (90, 2859, 51), (91, 2881, 51),
(92, 2811, 52), (93, 2295, 58), (94, 2392, 61), (95, 2600, 61), (96, 2380, 63),
(97, 2754, 55), (98, 2867, 49), (99, 2386, 51), (100, 2534, 50), (101, 2422, 61),
(102, 2451, 54), (103, 2383, 59), (104, 2459, 54)"""
        self.data_text.insert("1.0", sample)

        # --- Action Buttons Row 1 ---
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=5)

        self.run_btn = tk.Button(btn_frame, text="⚡ GENERATE PACK GROUPINGS", command=self.run_calculation,
                                 bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), padx=15, pady=5)
        self.run_btn.pack(side="left", padx=3)

        self.map_btn = tk.Button(btn_frame, text="🗺️ View Cell Map", command=self.generate_cell_map,
                                 bg="#9C27B0", fg="white", font=("Arial", 11, "bold"), padx=15, pady=5)
        self.map_btn.pack(side="left", padx=3)

        self.save_btn = tk.Button(btn_frame, text="💾 Save Results", command=self.save_results,
                                  bg="#FF9800", fg="white", font=("Arial", 11, "bold"), padx=15, pady=5)
        self.save_btn.pack(side="left", padx=3)

        self.print_btn = tk.Button(btn_frame, text="🖨️ Print Mode", command=self.print_results,
                                   bg="#795548", fg="white", font=("Arial", 11, "bold"), padx=15, pady=5)
        self.print_btn.pack(side="left", padx=3)

        self.excel_btn = tk.Button(btn_frame, text="📊 Export Excel", command=self.export_excel,
                                   bg="#009688", fg="white", font=("Arial", 11, "bold"), padx=15, pady=5)
        self.excel_btn.pack(side="left", padx=3)

        self.history_btn = tk.Button(btn_frame, text="📜 View History", command=self.view_history,
                                     bg="#607D8B", fg="white", font=("Arial", 11, "bold"), padx=15, pady=5)
        self.history_btn.pack(side="left", padx=3)

        # --- Heatmap toggle ---
        heat_frame = tk.Frame(root)
        heat_frame.pack(pady=2)
        self.heatmap_var = tk.BooleanVar(value=False)
        self.heatmap_check = tk.Checkbutton(heat_frame, text="🌡️ Use Heatmap (smooth colour gradient)",
                                            variable=self.heatmap_var,
                                            command=self.refresh_map_if_visible,
                                            font=("Arial", 10))
        self.heatmap_check.pack()

        # --- Output Frame with Tabs ---
        out_frame = tk.LabelFrame(root, text=" 3. Results ", padx=10, pady=10)
        out_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.notebook = ttk.Notebook(out_frame)
        self.notebook.pack(fill="both", expand=True)

        tab_text = tk.Frame(self.notebook)
        self.notebook.add(tab_text, text="  📝 Results Text  ")

        self.output_text = scrolledtext.ScrolledText(tab_text, height=15, state='disabled')
        self.output_text.pack(fill="both", expand=True)

        tab_map = tk.Frame(self.notebook)
        self.notebook.add(tab_map, text="  🗺️ Cell Map  ")

        map_toolbar = tk.Frame(tab_map)
        map_toolbar.pack(fill="x", pady=(0, 5))

        self.save_map_btn = tk.Button(map_toolbar, text="💾 Save Map as PNG",
                                      command=self.save_map_png,
                                      bg="#607D8B", fg="white", font=("Arial", 10, "bold"))
        self.save_map_btn.pack(side="left", padx=5)

        self.map_status = tk.Label(map_toolbar, text="Click 'View Cell Map' to generate the map.",
                                   fg="gray", font=("Arial", 10))
        self.map_status.pack(side="left", padx=10)

        canvas_frame = tk.Frame(tab_map)
        canvas_frame.pack(fill="both", expand=True)

        self.map_canvas = tk.Canvas(canvas_frame, bg="#f0f0f0")
        self.map_canvas.grid(row=0, column=0, sticky="nsew")

        h_scroll = tk.Scrollbar(canvas_frame, orient="horizontal", command=self.map_canvas.xview)
        h_scroll.grid(row=1, column=0, sticky="ew")

        v_scroll = tk.Scrollbar(canvas_frame, orient="vertical", command=self.map_canvas.yview)
        v_scroll.grid(row=0, column=1, sticky="ns")

        self.map_canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)

    # ==========================================
    # IMPORT CSV
    # ==========================================
    def import_csv(self):
        filepath = filedialog.askopenfilename(
            title="Select CSV file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not filepath:
            return

        try:
            cells = []
            with open(filepath, 'r', newline='', encoding='utf-8-sig') as f:
                sample = f.read(2048)
                f.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=',;\t')
                except csv.Error:
                    dialect = csv.excel

                reader = csv.reader(f, dialect)
                header_skipped = False

                for row in reader:
                    if not row or len(row) < 3:
                        continue
                    if not header_skipped:
                        try:
                            int(row[0])
                            float(row[1])
                            float(row[2])
                        except ValueError:
                            header_skipped = True
                            continue
                        header_skipped = True

                    try:
                        cell_id = int(row[0].strip())
                        cap = float(row[1].strip())
                        ir = float(row[2].strip())
                        if len(row) >= 4:
                            volt = float(row[3].strip())
                            cells.append((cell_id, cap, ir, volt))
                        else:
                            cells.append((cell_id, cap, ir))
                    except ValueError:
                        continue

            if not cells:
                messagebox.showerror("CSV Error", "No valid data found in the CSV file.\n\nExpected format: ID, Capacity, IR[, Voltage]")
                return

            formatted = ", ".join(
                f"({c[0]}, {c[1]}, {c[2]}, {c[3]})" if len(c) == 4 else f"({c[0]}, {c[1]}, {c[2]})"
                for c in cells
            )
            self.data_text.delete("1.0", tk.END)
            self.data_text.insert("1.0", formatted)

            messagebox.showinfo("Import Successful", f"Loaded {len(cells)} cells from:\n{os.path.basename(filepath)}")

        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to read CSV:\n{e}")

    # ==========================================
    # CLEAR DATA
    # ==========================================
    def clear_data(self):
        if messagebox.askyesno("Clear Data", "Are you sure you want to clear all cell data?"):
            self.data_text.delete("1.0", tk.END)

    # ==========================================
    # SAVE RESULTS
    # ==========================================
    def save_results(self):
        content = self.output_text.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("No Results", "Please generate pack groupings first before saving.")
            return

        filepath = filedialog.asksaveasfilename(
            title="Save Results As",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile="battery_groupings.txt"
        )
        if not filepath:
            return

        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            messagebox.showinfo("Saved", f"Results saved to:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save file:\n{e}")

    # ==========================================
    # PRINT MODE
    # ==========================================
    def print_results(self):
        if not self.last_groups:
            messagebox.showwarning("No Results", "Please generate pack groupings first.")
            return

        filepath = filedialog.asksaveasfilename(
            title="Save Printable Summary",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")],
            initialfile="pack_printout.txt"
        )
        if not filepath:
            return

        series, parallel = self.last_config

        lines = []
        lines.append("=" * 60)
        lines.append(f"  {series}S{parallel}P BATTERY PACK BUILD SHEET".center(60))
        lines.append(f"  Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}".center(60))
        lines.append("=" * 60)
        lines.append("")

        for p_idx, group in enumerate(self.last_groups, 1):
            group_sorted = sorted(group, key=lambda c: c[1], reverse=True)
            avg_cap = sum(c[1] for c in group_sorted) / len(group_sorted)
            avg_ir = sum(c[2] for c in group_sorted) / len(group_sorted)
            lines.append(f"┌── PARALLEL GROUP P{p_idx} ──────────────────────────┐")
            lines.append(f"│  Avg: {avg_cap:.0f} mAh | {avg_ir:.0f} mΩ                 │")
            for s_idx, cell in enumerate(group_sorted, 1):
                line = f"│   Cell {s_idx}: ID #{cell[0]:<4} | {int(cell[1])} mAh | {int(cell[2])} mΩ"
                line = line.ljust(53) + " │"
                lines.append(line)
            lines.append(f"└{'─' * 51}┘")
            lines.append("")

        if self.last_discarded:
            lines.append("=" * 60)
            lines.append(f"  DISCARDED CELLS ({len(self.last_discarded)})".center(60))
            lines.append("=" * 60)
            for cell in sorted(self.last_discarded, key=lambda c: c[0]):
                lines.append(f"  ID #{cell[0]:<4} | {int(cell[1])} mAh | {int(cell[2])} mΩ")

        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write("\n".join(lines))
            messagebox.showinfo("Saved", f"Printout saved to:\n{filepath}\n\nYou can now open it and send it to a printer.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save printout:\n{e}")

    # ==========================================
    # EXPORT EXCEL
    # ==========================================
    def export_excel(self):
        if not EXCEL_AVAILABLE:
            messagebox.showerror("Missing Library",
                                 "openpyxl is not installed.\n\nOpen Command Prompt and run:\n\npip install openpyxl")
            return

        if not self.last_groups:
            messagebox.showwarning("No Results", "Please generate pack groupings first.")
            return

        filepath = filedialog.asksaveasfilename(
            title="Export to Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="battery_pack.xlsx"
        )
        if not filepath:
            return

        try:
            wb = openpyxl.Workbook()

            # --- Sheet 1: Groups ---
            ws1 = wb.active
            ws1.title = "Parallel Groups"

            header_fill = PatternFill(start_color="4CAF50", end_color="4CAF50", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")

            headers = ["Parallel Group", "Cell Slot", "Physical ID", "Capacity (mAh)", "IR (mΩ)"]
            for col, h in enumerate(headers, 1):
                c = ws1.cell(row=1, column=col, value=h)
                c.fill = header_fill
                c.font = header_font
                c.alignment = Alignment(horizontal="center")

            row = 2
            for p_idx, group in enumerate(self.last_groups, 1):
                group_sorted = sorted(group, key=lambda c: c[1], reverse=True)
                for s_idx, cell in enumerate(group_sorted, 1):
                    ws1.cell(row=row, column=1, value=f"P{p_idx}")
                    ws1.cell(row=row, column=2, value=s_idx)
                    ws1.cell(row=row, column=3, value=cell[0])
                    ws1.cell(row=row, column=4, value=cell[1])
                    ws1.cell(row=row, column=5, value=cell[2])
                    row += 1

            for col in range(1, 6):
                ws1.column_dimensions[chr(64 + col)].width = 18

            # --- Sheet 2: Summary ---
            ws2 = wb.create_sheet("Summary")
            ws2.cell(row=1, column=1, value="Parallel Group").font = header_font
            ws2.cell(row=1, column=1).fill = header_fill
            ws2.cell(row=1, column=2, value="Avg Capacity").font = header_font
            ws2.cell(row=1, column=2).fill = header_fill
            ws2.cell(row=1, column=3, value="Avg IR").font = header_font
            ws2.cell(row=1, column=3).fill = header_fill

            for p_idx, group in enumerate(self.last_groups, 1):
                avg_cap = sum(c[1] for c in group) / len(group)
                avg_ir = sum(c[2] for c in group) / len(group)
                ws2.cell(row=p_idx + 1, column=1, value=f"P{p_idx}")
                ws2.cell(row=p_idx + 1, column=2, value=round(avg_cap, 1))
                ws2.cell(row=p_idx + 1, column=3, value=round(avg_ir, 1))

            ws2.column_dimensions['A'].width = 18
            ws2.column_dimensions['B'].width = 18
            ws2.column_dimensions['C'].width = 18

            # --- Sheet 3: Discarded ---
            if self.last_discarded:
                ws3 = wb.create_sheet("Discarded")
                ws3.cell(row=1, column=1, value="Physical ID").font = header_font
                ws3.cell(row=1, column=1).fill = PatternFill(start_color="E74C3C", end_color="E74C3C", fill_type="solid")
                ws3.cell(row=1, column=2, value="Capacity (mAh)").font = header_font
                ws3.cell(row=1, column=2).fill = PatternFill(start_color="E74C3C", end_color="E74C3C", fill_type="solid")
                ws3.cell(row=1, column=3, value="IR (mΩ)").font = header_font
                ws3.cell(row=1, column=3).fill = PatternFill(start_color="E74C3C", end_color="E74C3C", fill_type="solid")

                for i, cell in enumerate(sorted(self.last_discarded, key=lambda c: c[0]), 2):
                    ws3.cell(row=i, column=1, value=cell[0])
                    ws3.cell(row=i, column=2, value=cell[1])
                    ws3.cell(row=i, column=3, value=cell[2])

                for col in ['A', 'B', 'C']:
                    ws3.column_dimensions[col].width = 18

            wb.save(filepath)
            messagebox.showinfo("Exported", f"Excel file saved to:\n{filepath}")

        except Exception as e:
            messagebox.showerror("Excel Error", f"Failed to export:\n{e}")

    # ==========================================
    # VIEW HISTORY
    # ==========================================
    def view_history(self):
        if not os.path.exists(HISTORY_FILE):
            messagebox.showinfo("No History", "No build history yet.\n\nGenerate at least one pack to start recording history.")
            return

        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                history = json.load(f)
        except Exception as e:
            messagebox.showerror("History Error", f"Could not read history:\n{e}")
            return

        win = tk.Toplevel(self.root)
        win.title("Build History")
        win.geometry("700x500")

        txt = scrolledtext.ScrolledText(win, height=25, state="normal")
        txt.pack(fill="both", expand=True, padx=10, pady=10)

        txt.insert(tk.END, f"Total Builds Recorded: {len(history)}\n")
        txt.insert(tk.END, "=" * 60 + "\n\n")

        for i, entry in enumerate(reversed(history), 1):
            txt.insert(tk.END, f"--- Build #{len(history) - i + 1} ---\n")
            txt.insert(tk.END, f"Date:     {entry.get('date', 'unknown')}\n")
            txt.insert(tk.END, f"Config:   {entry.get('config', 'unknown')}\n")
            txt.insert(tk.END, f"Cells:    {entry.get('used_count', '?')} used, {entry.get('discard_count', '?')} discarded\n")
            txt.insert(tk.END, f"Avg Cap:  {entry.get('avg_cap', '?')} mAh\n")
            txt.insert(tk.END, f"Avg IR:   {entry.get('avg_ir', '?')} mΩ\n")
            txt.insert(tk.END, "\n")

        txt.config(state="disabled")

    def save_build_history(self, series, parallel, groups, discarded, raw_data):
        try:
            all_used = [c for g in groups for c in g]
            avg_cap = sum(c[1] for c in all_used) / len(all_used)
            avg_ir = sum(c[2] for c in all_used) / len(all_used)

            entry = {
                "date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "config": f"{series}S{parallel}P",
                "used_count": len(all_used),
                "discard_count": len(discarded),
                "avg_cap": round(avg_cap, 1),
                "avg_ir": round(avg_ir, 1),
            }

            history = []
            if os.path.exists(HISTORY_FILE):
                try:
                    with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                        history = json.load(f)
                except Exception:
                    history = []

            history.append(entry)

            with open(HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(history, f, indent=2)

        except Exception:
            pass  # Never block the user for a history write failure

    # ==========================================
    # SAVE MAP AS PNG
    # ==========================================
    def save_map_png(self):
        if self.current_map_image is None:
            messagebox.showwarning("No Map", "Please generate the cell map first.")
            return

        filepath = filedialog.asksaveasfilename(
            title="Save Cell Map As",
            defaultextension=".png",
            filetypes=[("PNG image", "*.png")],
            initialfile="battery_cell_map.png"
        )
        if not filepath:
            return

        try:
            self.current_map_image.save(filepath)
            messagebox.showinfo("Saved", f"Map saved to:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save image:\n{e}")

    # ==========================================
    # PARSE DATA
    # ==========================================
    def parse_data(self):
        raw_data = []
        text_content = self.data_text.get("1.0", tk.END).strip()

        pattern = r'\(\s*(\d+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*(?:,\s*([\d.]+)\s*)?\)'
        matches = re.findall(pattern, text_content)

        if matches:
            for match in matches:
                try:
                    cell_id = int(match[0])
                    cap = float(match[1])
                    ir = float(match[2])
                    if match[3]:
                        volt = float(match[3])
                        raw_data.append((cell_id, cap, ir, volt))
                    else:
                        raw_data.append((cell_id, cap, ir))
                except ValueError:
                    continue
        else:
            for line in text_content.split('\n'):
                if not line.strip():
                    continue
                try:
                    parts = line.replace('\t', ',').split(',')
                    if len(parts) >= 3:
                        cell_id = int(parts[0].strip())
                        cap = float(parts[1].strip())
                        ir = float(parts[2].strip())
                        if len(parts) >= 4:
                            volt = float(parts[3].strip())
                            raw_data.append((cell_id, cap, ir, volt))
                        else:
                            raw_data.append((cell_id, cap, ir))
                except ValueError:
                    continue

        return raw_data

    # ==========================================
    # DUPLICATE ID DETECTION
    # ==========================================
    def check_duplicate_ids(self, raw_data):
        seen = {}
        duplicates = []
        for cell in raw_data:
            cid = cell[0]
            if cid in seen:
                duplicates.append(cid)
            else:
                seen[cid] = True
        return duplicates

    # ==========================================
    # VOLTAGE SANITY CHECK
    # ==========================================
    def check_voltages(self, raw_data, min_volt):
        warnings = []
        for cell in raw_data:
            if len(cell) >= 4 and cell[3] < min_volt:
                warnings.append(cell)
        return warnings

    # ==========================================
    # RUN CALCULATION
    # ==========================================
    def run_calculation(self):
        try:
            series = int(self.series_entry.get())
            parallel = int(self.parallel_entry.get())
            min_cap = float(self.min_cap_entry.get()) if self.min_cap_entry.get().strip() else 0.0
            max_ir = float(self.max_ir_entry.get()) if self.max_ir_entry.get().strip() else 9999.0
            min_volt = float(self.min_volt_entry.get()) if self.min_volt_entry.get().strip() else 0.0
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numbers for all configuration fields.")
            return

        raw_data = self.parse_data()

        if not raw_data:
            messagebox.showerror("Data Error", "No valid cell data found. Please check your input format.")
            return

        # --- Duplicate ID check ---
        duplicates = self.check_duplicate_ids(raw_data)
        if duplicates:
            msg = "Duplicate Physical IDs detected:\n\n"
            msg += ", ".join(str(d) for d in sorted(set(duplicates)))
            msg += "\n\nThis usually means you have a copy/paste error. Continue anyway?"
            if not messagebox.askyesno("Duplicate IDs", msg):
                return

        # --- Voltage sanity check ---
        if min_volt > 0:
            voltage_warnings = self.check_voltages(raw_data, min_volt)
            if voltage_warnings:
                msg = f"{len(voltage_warnings)} cell(s) have voltage below {min_volt}V:\n\n"
                for cell in voltage_warnings[:10]:
                    msg += f"  ID #{cell[0]}: {cell[3]}V\n"
                if len(voltage_warnings) > 10:
                    msg += f"  ... and {len(voltage_warnings) - 10} more\n"
                msg += "\nThese cells may be self-discharging and unsafe. Continue anyway?"
                if not messagebox.askyesno("Voltage Warning", msg):
                    return

        valid_cells = [c for c in raw_data if c[1] >= min_cap and c[2] <= max_ir]
        total_needed = series * parallel

        if len(valid_cells) < total_needed:
            messagebox.showerror("Not Enough Cells", f"You need {total_needed} cells, but only {len(valid_cells)} passed your limits.")
            return

        sorted_cells = sorted(valid_cells, key=lambda x: (x[2], x[1]))
        used_cells = sorted_cells[:total_needed]
        discarded_cells = sorted_cells[total_needed:]
        rejected_by_filter = [c for c in raw_data if c[1] < min_cap or c[2] > max_ir]

        groups = [[] for _ in range(series)]
        for i in range(total_needed):
            cycle_position = i % (series * 2)
            if cycle_position < series:
                group_index = cycle_position
            else:
                group_index = (series * 2 - 1) - cycle_position
            groups[group_index].append(used_cells[i])

        self.last_groups = groups
        self.last_discarded = rejected_by_filter + discarded_cells
        self.last_config = (series, parallel)
        self.last_raw_data = raw_data

        # Save build to history
        self.save_build_history(series, parallel, groups, self.last_discarded, raw_data)

        # Build output text
        output_str = "="*60 + "\n"
        output_str += f" RECOMMENDED {series}S{parallel}P BATTERY PACK GROUPING \n"
        output_str += "="*60 + "\n"
        output_str += f"Note: You need {series} Parallel Groups (P1 to P{series}).\n"
        output_str += f"      Each group must contain {parallel} cells wired in parallel.\n\n"

        for p_idx, group in enumerate(groups, 1):
            group.sort(key=lambda x: x[1], reverse=True)
            avg_cap = sum(c[1] for c in group) / len(group)
            avg_ir = sum(c[2] for c in group) / len(group)

            output_str += f"--- PARALLEL GROUP P{p_idx} ({parallel} Cells) ---\n"
            output_str += f"    Avg Capacity: {avg_cap:.1f} mAh | Avg IR: {avg_ir:.1f} mOhm\n"
            for s_idx, cell in enumerate(group, 1):
                output_str += f"      Cell {s_idx}: [Physical ID: {cell[0]:3}] -> Capacity={cell[1]} mAh, IR={cell[2]} mOhm\n"
            output_str += "\n"

        output_str += "="*60 + "\n PACK BALANCE SUMMARY \n" + "="*60 + "\n"
        for p_idx, group in enumerate(groups, 1):
            avg_cap = sum(c[1] for c in group) / len(group)
            avg_ir = sum(c[2] for c in group) / len(group)
            output_str += f"P{p_idx:2} Average -> Capacity: {avg_cap:.1f} mAh | IR: {avg_ir:.1f} mOhm\n"

        if rejected_by_filter or discarded_cells:
            output_str += "\n" + "="*60 + "\n 🗑️ DISCARDED CELLS 🗑️ \n" + "="*60 + "\n"
            if rejected_by_filter:
                output_str += "\n--- Failed your custom limits ---\n"
                for cell in rejected_by_filter:
                    output_str += f"  [ID: {cell[0]:3}] -> Cap: {cell[1]} mAh | IR: {cell[2]} mOhm\n"
            if discarded_cells:
                output_str += "\n--- Extras (Didn't make the final cut) ---\n"
                discarded_cells.sort(key=lambda x: x[0])
                for cell in discarded_cells:
                    output_str += f"  [ID: {cell[0]:3}] -> Cap: {cell[1]} mAh | IR: {cell[2]} mOhm\n"
            output_str += f"\nTotal Discarded: {len(rejected_by_filter) + len(discarded_cells)} cells\n"

        self.output_text.config(state='normal')
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert(tk.END, output_str)
        self.output_text.config(state='disabled')

        self.notebook.select(1)
        self.root.after(50, self.generate_cell_map)

    # ==========================================
    # GENERATE CELL MAP (in-app preview)
    # ==========================================
    def generate_cell_map(self):
        if not PIL_AVAILABLE:
            messagebox.showerror("Missing Library",
                                 "Pillow is not installed.\n\nOpen Command Prompt and run:\n\npip install pillow")
            return

        if self.last_groups is None:
            messagebox.showwarning("No Data", "Please click 'Generate Pack Groupings' first.")
            return

        try:
            self.current_map_image = self._build_cell_map_image()

            canvas_w = self.map_canvas.winfo_width() or 850
            canvas_h = self.map_canvas.winfo_height() or 400

            img_w, img_h = self.current_map_image.size
            scale = min(1.0, (canvas_w - 20) / img_w, (canvas_h - 20) / img_h)
            if scale < 1.0:
                new_size = (int(img_w * scale), int(img_h * scale))
                display_img = self.current_map_image.resize(new_size, Image.LANCZOS)
            else:
                display_img = self.current_map_image

            self.current_map_photo = ImageTk.PhotoImage(display_img)

            self.map_canvas.delete("all")
            self.map_canvas.create_image(10, 10, anchor="nw", image=self.current_map_photo)
            self.map_canvas.configure(scrollregion=(0, 0, display_img.size[0] + 20, display_img.size[1] + 20))

            self.map_status.config(text=f"Map generated ({img_w}x{img_h}px). Use 'Save Map as PNG' to export.",
                                   fg="green")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to create map:\n{e}")

    def refresh_map_if_visible(self):
        if self.last_groups is not None:
            self.generate_cell_map()

    # ==========================================
    # COLOUR HELPERS
    # ==========================================
    def _get_color_4cat(self, cell, min_cap, max_cap, min_ir, max_ir):
        cap_range = max(max_cap - min_cap, 1)
        ir_range = max(max_ir - min_ir, 1)

        cap_score = (cell[1] - min_cap) / cap_range
        ir_score = 1.0 - ((cell[2] - min_ir) / ir_range)
        score = (cap_score + ir_score) / 2

        if score >= 0.75:
            return (34, 177, 76), "white"
        elif score >= 0.5:
            return (146, 208, 80), "black"
        elif score >= 0.25:
            return (255, 217, 102), "black"
        else:
            return (231, 76, 60), "white"

    def _heatmap_color(self, score):
        score = max(0.0, min(1.0, score))
        stops = [
            (0.00, (215, 48, 39)),
            (0.25, (253, 174, 97)),
            (0.50, (255, 255, 191)),
            (0.75, (166, 217, 106)),
            (1.00, (26, 152, 80)),
        ]
        for i in range(len(stops) - 1):
            low_pos, low_col = stops[i]
            high_pos, high_col = stops[i + 1]
            if low_pos <= score <= high_pos:
                t = (score - low_pos) / (high_pos - low_pos)
                r = int(low_col[0] + t * (high_col[0] - low_col[0]))
                g = int(low_col[1] + t * (high_col[1] - low_col[1]))
                b = int(low_col[2] + t * (high_col[2] - low_col[2]))
                return (r, g, b)
        return stops[-1][1]

    def _get_text_color(self, bg):
        brightness = (bg[0] * 299 + bg[1] * 587 + bg[2] * 114) / 1000
        return "black" if brightness > 140 else "white"

    def _get_cell_color(self, cell, min_cap, max_cap, min_ir, max_ir, use_heatmap):
        if use_heatmap:
            cap_range = max(max_cap - min_cap, 1)
            ir_range = max(max_ir - min_ir, 1)
            cap_score = (cell[1] - min_cap) / cap_range
            ir_score = 1.0 - ((cell[2] - min_ir) / ir_range)
            score = (cap_score + ir_score) / 2
            bg = self._heatmap_color(score)
            txt = self._get_text_color(bg)
            return bg, txt
        else:
            return self._get_color_4cat(cell, min_cap, max_cap, min_ir, max_ir)

    # ==========================================
    # BUILD MAP IMAGE (in memory)
    # ==========================================
    def _build_cell_map_image(self):
        use_heatmap = self.heatmap_var.get()
        groups = self.last_groups
        discarded = self.last_discarded
        series, parallel = self.last_config

        all_used = [c for g in groups for c in g]
        min_cap = min(c[1] for c in all_used)
        max_cap = max(c[1] for c in all_used)
        min_ir = min(c[2] for c in all_used)
        max_ir = max(c[2] for c in all_used)

        cell_w = 70
        cell_h = 70
        padding = 15
        top_margin = 140
        left_margin = 30

        cols = series
        rows = parallel

        img_w = left_margin * 2 + cols * (cell_w + padding)
        img_h = top_margin + rows * (cell_h + padding) + 220

        img = Image.new("RGB", (img_w, img_h), (255, 255, 255))
        draw = ImageDraw.Draw(img)

        try:
            font_big = ImageFont.truetype("arial.ttf", 22)
            font_med = ImageFont.truetype("arial.ttf", 14)
            font_small = ImageFont.truetype("arial.ttf", 11)
        except IOError:
            font_big = ImageFont.load_default()
            font_med = ImageFont.load_default()
            font_small = ImageFont.load_default()

        mode_label = "Heatmap Mode" if use_heatmap else "Category Mode"
        title = f"{series}S{parallel}P Cell Map ({mode_label})"
        draw.text((left_margin, 20), title, fill=(0, 0, 0), font=font_big)
        draw.text((left_margin, 55),
                  f"Min Cap: {min_cap:.0f} | Max Cap: {max_cap:.0f} | Min IR: {min_ir:.0f} | Max IR: {max_ir:.0f}",
                  fill=(80, 80, 80), font=font_small)

        legend_y = 90
        if use_heatmap:
            bar_x = left_margin
            bar_w = 400
            bar_h = 20
            for i in range(bar_w):
                score = i / (bar_w - 1)
                color = self._heatmap_color(score)
                draw.line([(bar_x + i, legend_y), (bar_x + i, legend_y + bar_h)], fill=color)
            draw.rectangle([bar_x, legend_y, bar_x + bar_w, legend_y + bar_h], outline=(0, 0, 0))
            draw.text((bar_x, legend_y + bar_h + 3), "Weak / Bad", fill=(0, 0, 0), font=font_small)
            draw.text((bar_x + bar_w - 70, legend_y + bar_h + 3), "Strong / Good", fill=(0, 0, 0), font=font_small)
        else:
            legend_items = [
                ((34, 177, 76), "Strong"),
                ((146, 208, 80), "Good"),
                ((255, 217, 102), "Weak"),
                ((231, 76, 60), "Discarded"),
            ]
            legend_x = left_margin
            for color, label in legend_items:
                draw.rectangle([legend_x, legend_y, legend_x + 20, legend_y + 20], fill=color, outline=(0, 0, 0))
                draw.text((legend_x + 25, legend_y + 3), label, fill=(0, 0, 0), font=font_small)
                legend_x += 100

        for p_idx in range(series):
            x = left_margin + p_idx * (cell_w + padding) + cell_w // 2
            header = f"P{p_idx + 1}"
            bbox = draw.textbbox((0, 0), header, font=font_med)
            tw = bbox[2] - bbox[0]
            draw.text((x - tw // 2, top_margin - 25), header, fill=(0, 0, 0), font=font_med)

        for p_idx, group in enumerate(groups):
            group_sorted = sorted(group, key=lambda c: c[1], reverse=True)
            for s_idx, cell in enumerate(group_sorted):
                x1 = left_margin + p_idx * (cell_w + padding)
                y1 = top_margin + s_idx * (cell_h + padding)
                x2 = x1 + cell_w
                y2 = y1 + cell_h

                bg, txt_color = self._get_cell_color(cell, min_cap, max_cap, min_ir, max_ir, use_heatmap)
                draw.rectangle([x1, y1, x2, y2], fill=bg, outline=(0, 0, 0), width=2)

                id_str = str(cell[0])
                bbox = draw.textbbox((0, 0), id_str, font=font_big)
                tw = bbox[2] - bbox[0]
                draw.text((x1 + (cell_w - tw) // 2, y1 + 8), id_str, fill=txt_color, font=font_big)

                info = f"{int(cell[1])}mAh"
                bbox2 = draw.textbbox((0, 0), info, font=font_small)
                tw2 = bbox2[2] - bbox2[0]
                draw.text((x1 + (cell_w - tw2) // 2, y1 + 38), info, fill=txt_color, font=font_small)

                info2 = f"{int(cell[2])}mΩ"
                bbox3 = draw.textbbox((0, 0), info2, font=font_small)
                tw3 = bbox3[2] - bbox3[0]
                draw.text((x1 + (cell_w - tw3) // 2, y1 + 52), info2, fill=txt_color, font=font_small)

        if discarded:
            discard_y = top_margin + rows * (cell_h + padding) + 30
            draw.text((left_margin, discard_y - 25),
                      f"Discarded Cells ({len(discarded)})", fill=(0, 0, 0), font=font_med)

            small_w = 55
            small_h = 55
            small_pad = 8
            max_per_row = (img_w - left_margin * 2) // (small_w + small_pad)

            discarded_sorted = sorted(discarded, key=lambda c: c[0])
            for i, cell in enumerate(discarded_sorted):
                row = i // max_per_row
                col = i % max_per_row
                x1 = left_margin + col * (small_w + small_pad)
                y1 = discard_y + row * (small_h + small_pad)
                x2 = x1 + small_w
                y2 = y1 + small_h

                draw.rectangle([x1, y1, x2, y2], fill=(231, 76, 60), outline=(0, 0, 0))
                id_str = str(cell[0])
                bbox = draw.textbbox((0, 0), id_str, font=font_med)
                tw = bbox[2] - bbox[0]
                draw.text((x1 + (small_w - tw) // 2, y1 + 6), id_str, fill="white", font=font_med)

                cap_str = f"{int(cell[1])}"
                bbox2 = draw.textbbox((0, 0), cap_str, font=font_small)
                tw2 = bbox2[2] - bbox2[0]
                draw.text((x1 + (small_w - tw2) // 2, y1 + 24), cap_str, fill="white", font=font_small)

                ir_str = f"{int(cell[2])}Ω"
                bbox3 = draw.textbbox((0, 0), ir_str, font=font_small)
                tw3 = bbox3[2] - bbox3[0]
                draw.text((x1 + (small_w - tw3) // 2, y1 + 38), ir_str, fill="white", font=font_small)

        return img


if __name__ == "__main__":
    root = tk.Tk()
    app = BatteryGrouperGUI(root)
    root.mainloop()