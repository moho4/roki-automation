import tkinter as tk
from tkinter import messagebox
import subprocess
import sys
import os
import re

SCRIPT_NAME = "12.py"
DEFAULT_CONFIG = "config_roki.json"
REPORT_DIR = "poročila"


def get_base_dir():
    return os.path.dirname(os.path.abspath(__file__))


def find_latest_input_excel():
    """
    Samodejno najde Excel datoteko z največjim vX.
    Primer datotek:
        izpitni_roki_25_26_prilagojen_v5.xlsx
        izpitni_roki_25_26_prilagojen_v7.xlsx
    Vrne pot do največjega vX ali None.
    """
    base_dir = get_base_dir()
    files = os.listdir(base_dir)

    pattern = re.compile(r"izpitni_roki_25_26_prilagojen_v(\d+)\.xlsx$")
    best_file = None
    best_version = -1

    for f in files:
        m = pattern.match(f)
        if m:
            version = int(m.group(1))
            if version > best_version:
                best_version = version
                best_file = f

    if best_file:
        return os.path.join(base_dir, best_file)
    return None


def run_analysis_auto(log_widget, run_button):
    base_dir = get_base_dir()

    # najdi zadnjo verzijo Excel tabele
    input_path = find_latest_input_excel()
    if not input_path:
        messagebox.showerror(
            "Napaka",
            "Ni mogoče najti Excel tabele 'izpitni_roki_25_26_prilagojen_vX.xlsx' v isti mapi."
        )
        return

    config_path = os.path.join(base_dir, DEFAULT_CONFIG)
    script_path = os.path.join(base_dir, SCRIPT_NAME)

    if not os.path.exists(script_path):
        messagebox.showerror("Napaka", f"Glavne skripte {SCRIPT_NAME} ni v mapi.")
        return

    if not os.path.exists(config_path):
        messagebox.showerror("Napaka", f"Config datoteka {DEFAULT_CONFIG} ni v mapi.")
        return

    cmd = [
        sys.executable, script_path,
        "--input", input_path,
        "--config", config_path,
    ]

    # UI: Log + disable button
    run_button.config(state="disabled")
    log_widget.config(state="normal")
    log_widget.delete("1.0", tk.END)
    log_widget.insert(tk.END, f"Najdena tabela: {os.path.basename(input_path)}\n")
    log_widget.insert(tk.END, "Zaganjam analizo...\n")
    log_widget.update()

    try:
        result = subprocess.run(
            cmd,
            cwd=base_dir,
            capture_output=True,
            text=True
        )

        if result.stdout:
            log_widget.insert(tk.END, "\n--- OUTPUT ---\n" + result.stdout)

        if result.stderr:
            log_widget.insert(tk.END, "\n--- ERRORS ---\n" + result.stderr)

        if result.returncode == 0:
            log_widget.insert(tk.END, "\n\n✅ Analiza uspešno zaključena.\n")
            reports_dir = os.path.join(base_dir, REPORT_DIR)
            if os.path.isdir(reports_dir):
                try:
                    os.startfile(reports_dir)
                    log_widget.insert(tk.END, f"Odprta mapa poročil: {reports_dir}\n")
                except Exception as e:
                    log_widget.insert(
                        tk.END,
                        f"Mape poročil ni bilo mogoče odpreti: {e}\n"
                    )
        else:
            log_widget.insert(
                tk.END,
                f"\n❌ Napaka (returncode={result.returncode}) – preveri log.\n"
            )

    except Exception as e:
        log_widget.insert(tk.END, "\n❌ Izjema: " + str(e) + "\n")

    log_widget.config(state="disabled")
    run_button.config(state="normal")


def main():
    base_dir = get_base_dir()

    root = tk.Tk()
    root.title("Analiza izpitnih rokov – One-Click UI")
    root.geometry("700x450")

    tk.Label(root, text="One-click zaganjalnik analize izpitnih rokov").pack(pady=10)

    # LOG okno
    log_text = tk.Text(root, height=20, state="disabled")
    log_text.pack(fill="both", expand=True, padx=10, pady=10)

    # Gumbi
    frame_btn = tk.Frame(root)
    frame_btn.pack(pady=5)

    run_button = tk.Button(
        frame_btn,
        text="Zaženi analizo (samodejno)",
        width=25,
        command=lambda: run_analysis_auto(log_text, run_button)
    )
    run_button.pack(side="left", padx=5)

    def open_reports():
        path = os.path.join(base_dir, REPORT_DIR)
        if not os.path.isdir(path):
            messagebox.showwarning("Ni mape", "Mapa poročila še ne obstaja.")
            return
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("Napaka", f"Mape poročil ni bilo mogoče odpreti: {e}")

    tk.Button(
        frame_btn,
        text="Odpri mapo poročil",
        width=18,
        command=open_reports
    ).pack(side="left", padx=5)

    root.mainloop()


if __name__ == "__main__":
    main()
