# Universal Battery Pack Grouper
![Screenshot](battery_grouper.png)
A Python-based GUI tool designed to help battery builders create perfectly balanced lithium-ion battery packs (e.g., 18650, 21700). 

## 📖 Description

Building a custom battery pack (like a 13S5P) requires careful matching of cell capacities and Internal Resistance (IR). If cells are mismatched, the pack will become unbalanced, lose capacity quickly, and potentially become a fire hazard.

This tool solves that problem by:
1. Taking your tested cell data (Capacity and IR).
2. Sorting the cells to find the best matches.
3. Distributing them into perfectly balanced Parallel Groups (P1, P2, P3, etc.) using a "snake" algorithm.
4. Providing a clear list of cells to use, and a separate list of cells to discard.

It works for **any configuration** (e.g., 3S2P, 10S4P, 13S5P, 20S10P) and includes optional safety filters to automatically reject bad cells before grouping.

---

## 🚀 Installation & Usage

### 🪟 Windows

**Prerequisites:**
You must have Python installed. If you don't, download it from [python.org](https://www.python.org/downloads/). 
* **CRITICAL:** During installation, check the box that says **"Add Python to PATH"**.

**Step 1: Install the Script**
1. Download the `battery_gui.py` file and place it in a folder (e.g., `C:\BatteryTools`).
2. Open the Command Prompt (search "cmd" in the Start menu).
3. Navigate to your folder by typing: `cd C:\BatteryTools`
4. Run the script by typing:
   ```bash
   python battery_gui.py
