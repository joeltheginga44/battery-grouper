import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import math
import re
import csv
import os

try:
    from PIL import Image, ImageDraw, ImageFont
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


class BatteryGrouperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Universal Battery Pack Grouper")
        self.root.geometry("820x780")

        # Store last calculation results so the Cell Map button can use them
        self.last_groups = None
        self.last_discarded = None
        self.last_config = None

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

        # --- Data Input Frame ---
        data_frame = tk.LabelFrame(root, text=" 2. Cell Data ", padx=10, pady=10)
        data_frame.pack(fill="both", expand=True, padx=10, pady=5)

        top_row = tk.Frame(data_frame)
        top_row.pack(fill="x", pady=(0, 5))

        tk.Label(top_row, text="Format: (ID, Capacity, IR), (ID, Capacity, IR)...").pack(side="left")

        self.import_btn = tk.Button(top_row, text="📂 Import CSV", command=self.import_csv, bg="#2196F3", fg="white")
        self.import_btn.pack(side="right", padx=5)

        self.clear_btn = tk.Button(top_row, text="🗑️ Clear", command=self.clear_data, bg="#9E9E9E", fg="white")
        self.clear_btn.pack(side="right", padx=5)

        self.data_text = scrolledtext.ScrolledText(data_frame, height=10)
        self.data_text.pack(fill="both", expand=True, pady=5)

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

        # --- Action Buttons ---
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=5)

        self.run_btn = tk.Button(btn_frame, text="⚡ GENERATE PACK GROUPINGS", command=self.run_calculation,
                                 bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), padx=15, pady=5)
        self.run_btn.pack(side="left", padx=5)

        self.map_btn = tk.Button(btn_frame, text="📊 Generate Cell Map", command=self.generate_cell_map,
                                 bg="#9C27B0", fg="white", font=("Arial", 11, "bold"), padx=15, pady=5)
        self.map_btn.pack(side="left", padx=5)

        self.save_btn = tk.Button(btn_frame, text="💾 Save Results", command=self.save_results,
                                  bg="#FF9800", fg="white", font=("Arial", 11, "bold"), padx=15, pady=5)
        self.save_btn.pack(side="left", padx=5)

        # --- Output Frame ---
        out_frame = tk.LabelFrame(root, text=" 3. Results ", padx=10, pady=10)
        out_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.output_text = scrolledtext.ScrolledText(out_frame, height=15, state='disabled')
        self.output_text.pack(fill="both", expand=True)

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
                        cells.append((cell_id, cap, ir))
                    except ValueError:
                        continue

            if not cells:
                messagebox.showerror("CSV Error", "No valid data found in the CSV file.\n\nExpected format: ID, Capacity, IR")
                return

            formatted = ", ".join(f"({c[0]}, {c[1]}, {c[2]})" for c in cells)
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
    # PARSE DATA
    # ==========================================
    def parse_data(self):
        raw_data = []
        text_content = self.data_text.get("1.0", tk.END).strip()

        pattern = r'\(\s*(\d+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*\)'
        matches = re.findall(pattern, text_content)

        if matches:
            for match in matches:
                try:
                    raw_data.append((int(match[0]), float(match[1]), float(match[2])))
                except ValueError:
                    continue
        else:
            for line in text_content.split('\n'):
                if not line.strip():
                    continue
                try:
                    parts = line.replace('\t', ',').split(',')
                    if len(parts) >= 3:
                        raw_data.append((int(parts[0].strip()), float(parts[1].strip()), float(parts[2].strip())))
                except ValueError:
                    continue

        return raw_data

    # ==========================================
    # RUN CALCULATION
    # ==========================================
    def run_calculation(self):
        try:
            series = int(self.series_entry.get())
            parallel = int(self.parallel_entry.get())
            min_cap = float(self.min_cap_entry.get()) if self.min_cap_entry.get().strip() else 0.0
            max_ir = float(self.max_ir_entry.get()) if self.max_ir_entry.get().strip() else 9999.0
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numbers for Series, Parallel, Capacity, and IR.")
            return

        raw_data = self.parse_data()

        if not raw_data:
            messagebox.showerror("Data Error", "No valid cell data found. Please check your input format.")
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

        # Store for later use by Cell Map button
        self.last_groups = groups
        self.last_discarded = rejected_by_filter + discarded_cells
        self.last_config = (series, parallel)

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

    # ==========================================
    # GENERATE CELL MAP IMAGE
    # ==========================================
    def generate_cell_map(self):
        if not PIL_AVAILABLE:
            messagebox.showerror("Missing Library",
                                 "Pillow is not installed.\n\nOpen Command Prompt and run:\n\npip install pillow")
            return

        if self.last_groups is None:
            messagebox.showwarning("No Data", "Please click 'Generate Pack Groupings' first.")
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
            self._draw_cell_map(filepath)
            messagebox.showinfo("Saved", f"Cell map saved to:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create map:\n{e}")

    def _get_color(self, cell, min_cap, max_cap, min_ir, max_ir):
        """Returns (bg, text) colour tuple based on quality."""
        cap_range = max(max_cap - min_cap, 1)
        ir_range = max(max_ir - min_ir, 1)

        cap_score = (cell[1] - min_cap) / cap_range   # 0.0 = worst, 1.0 = best
        ir_score = 1.0 - ((cell[2] - min_ir) / ir_range)  # 0.0 = worst, 1.0 = best

        score = (cap_score + ir_score) / 2

        if score >= 0.75:
            return (34, 177, 76), "white"    # strong green
        elif score >= 0.5:
            return (146, 208, 80), "black"   # light green
        elif score >= 0.25:
            return (255, 217, 102), "black"  # yellow
        else:
            return (231, 76, 60), "white"    # red

    def _draw_cell_map(self, filepath):
        groups = self.last_groups
        discarded = self.last_discarded
        series, parallel = self.last_config

        # Gather all used cells to determine scale
        all_used = [c for g in groups for c in g]
        min_cap = min(c[1] for c in all_used)
        max_cap = max(c[1] for c in all_used)
        min_ir = min(c[2] for c in all_used)
        max_ir = max(c[2] for c in all_used)

        # Layout constants
        cell_w = 70
        cell_h = 70
        padding = 15
        top_margin = 140
        left_margin = 30

        cols = series
        rows = parallel

        img_w = left_margin * 2 + cols * (cell_w + padding)
        img_h = top_margin + rows * (cell_h + padding) + 200

        img = Image.new("RGB", (img_w, img_h), (255, 255, 255))
        draw = ImageDraw.Draw(img)

        # Try to load a truetype font, fallback to default
        try:
            font_big = ImageFont.truetype("arial.ttf", 22)
            font_med = ImageFont.truetype("arial.ttf", 14)
            font_small = ImageFont.truetype("arial.ttf", 11)
        except IOError:
            font_big = ImageFont.load_default()
            font_med = ImageFont.load_default()
            font_small = ImageFont.load_default()

        # --- Title ---
        title = f"{series}S{parallel}P Cell Map"
        draw.text((left_margin, 20), title, fill=(0, 0, 0), font=font_big)
        draw.text((left_margin, 55),
                  f"Min Cap: {min_cap:.0f} | Max Cap: {max_cap:.0f} | Min IR: {min_ir:.0f} | Max IR: {max_ir:.0f}",
                  fill=(80, 80, 80), font=font_small)

        # --- Legend ---
        legend_y = 90
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

        # --- Column Headers (P1..P13) ---
        for p_idx in range(series):
            x = left_margin + p_idx * (cell_w + padding) + cell_w // 2
            header = f"P{p_idx + 1}"
            # Center the header text
            bbox = draw.textbbox((0, 0), header, font=font_med)
            tw = bbox[2] - bbox[0]
            draw.text((x - tw // 2, top_margin - 25), header, fill=(0, 0, 0), font=font_med)

        # --- Cell Grids ---
        for p_idx, group in enumerate(groups):
            # Sort so strongest is at the top of each column
            group_sorted = sorted(group, key=lambda c: c[1], reverse=True)

            for s_idx, cell in enumerate(group_sorted):
                x1 = left_margin + p_idx * (cell_w + padding)
                y1 = top_margin + s_idx * (cell_h + padding)
                x2 = x1 + cell_w
                y2 = y1 + cell_h

                bg, txt_color = self._get_color(cell, min_cap, max_cap, min_ir, max_ir)
                draw.rectangle([x1, y1, x2, y2], fill=bg, outline=(0, 0, 0), width=2)

                # Big ID number centered
                id_str = str(cell[0])
                bbox = draw.textbbox((0, 0), id_str, font=font_big)
                tw = bbox[2] - bbox[0]
                th = bbox[3] - bbox[1]
                draw.text((x1 + (cell_w - tw) // 2, y1 + 8),
                          id_str, fill=txt_color, font=font_big)

                # Small capacity + IR below
                info = f"{int(cell[1])}mAh"
                bbox2 = draw.textbbox((0, 0), info, font=font_small)
                tw2 = bbox2[2] - bbox2[0]
                draw.text((x1 + (cell_w - tw2) // 2, y1 + 38),
                          info, fill=txt_color, font=font_small)

                info2 = f"{int(cell[2])}mΩ"
                bbox3 = draw.textbbox((0, 0), info2, font=font_small)
                tw3 = bbox3[2] - bbox3[0]
                draw.text((x1 + (cell_w - tw3) // 2, y1 + 52),
                          info2, fill=txt_color, font=font_small)

        # --- Discarded Cells Section ---
        if discarded:
            discard_y = top_margin + rows * (cell_h + padding) + 30
            draw.text((left_margin, discard_y - 25),
                      f"Discarded Cells ({len(discarded)})",
                      fill=(0, 0, 0), font=font_med)

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
                draw.text((x1 + (small_w - tw) // 2, y1 + 6),
                          id_str, fill="white", font=font_med)

                cap_str = f"{int(cell[1])}"
                bbox2 = draw.textbbox((0, 0), cap_str, font=font_small)
                tw2 = bbox2[2] - bbox2[0]
                draw.text((x1 + (small_w - tw2) // 2, y1 + 24),
                          cap_str, fill="white", font=font_small)

                ir_str = f"{int(cell[2])}Ω"
                bbox3 = draw.textbbox((0, 0), ir_str, font=font_small)
                tw3 = bbox3[2] - bbox3[0]
                draw.text((x1 + (small_w - tw3) // 2, y1 + 38),
                          ir_str, fill="white", font=font_small)

        img.save(filepath)


if __name__ == "__main__":
    root = tk.Tk()
    app = BatteryGrouperGUI(root)
    root.mainloop()