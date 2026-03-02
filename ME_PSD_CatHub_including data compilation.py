import numpy as np
import matplotlib.pyplot as plt
from scipy import integrate
import tkinter as tk
from tkinter import filedialog, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import pandas as pd
import os
import subprocess

class PSDApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PSD Processing Interface")
        self.root.configure(bg="#dcdcdc")
        self.create_widgets()
        self.data_loaded = False
        self.root.mainloop()

    def create_widgets(self):
        self.entries = {}
        label_font = ("Helvetica", 10)
        entry_font = ("Helvetica", 10)
        title_font = ("Helvetica", 12, "bold")

        main_frame = tk.Frame(self.root, bg="#dcdcdc")
        main_frame.grid(row=0, column=0, sticky="nsew")

        # PSD Parameters Section
        psd_frame = tk.LabelFrame(main_frame, text="PSD Parameters", bg="#e6e6e6", padx=10, pady=10, font=title_font)
        psd_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        psd_labels = [
            "Modulation Period (s):",
            "Dummy Periods:",
            "Sample Periods:",
            "Harmonic Order (k):",
            "Start Phase (degrees):",
            "End Phase (degrees):",
            "Phase Step (degrees):",
            "Wavenumber Range Start:",
            "Wavenumber Range End:"
        ]

        for i, label in enumerate(psd_labels):
            tk.Label(psd_frame, text=label, bg="#e6e6e6", font=label_font).grid(row=i, column=0, sticky=tk.W, pady=2)
            entry = tk.Entry(psd_frame, font=entry_font, width=15)
            entry.grid(row=i, column=1, pady=2, sticky="w")
            self.entries[label] = entry

        button_frame = tk.Frame(psd_frame, bg="#e6e6e6")
        button_frame.grid(row=len(psd_labels), column=0, columnspan=2, pady=5)
        tk.Button(button_frame, text="Load Data", font=label_font, command=self.load_data).pack(side="left", padx=5)
        tk.Button(button_frame, text="Process & Plot PSD", font=label_font, command=self.process_and_plot).pack(side="left", padx=5)
        tk.Button(button_frame, text="Save Data", font=label_font, command=self.save_data).pack(side="left", padx=5)

        # --- Shared right panel for time and compilation ---
        right_panel = tk.Frame(main_frame, bg="#f0f0f0")
        right_panel.grid(row=0, column=1, rowspan=2, sticky="nsew", padx=5, pady=5)

        # Time Domain Section
        time_frame = tk.LabelFrame(right_panel, text="Intensity as a function of time", bg="#f0f0f0", padx=10, pady=8, font=title_font)
        time_frame.pack(fill="x", pady=(0, 8))  # Slight gap after

        tk.Label(time_frame, text="Total Time (s):", bg="#f0f0f0", font=label_font).grid(row=0, column=0, sticky=tk.W, pady=2)
        total_time_entry = tk.Entry(time_frame, font=entry_font, width=10)
        total_time_entry.grid(row=0, column=1, pady=2)
        self.entries["Total Time (s):"] = total_time_entry

        tk.Label(time_frame, text="Plot Wavenumber:", bg="#f0f0f0", font=label_font).grid(row=1, column=0, sticky=tk.W, pady=2)
        wn1_entry = tk.Entry(time_frame, font=entry_font, width=10)
        wn1_entry.grid(row=1, column=1, pady=2)
        wn2_entry = tk.Entry(time_frame, font=entry_font, width=10)
        wn2_entry.grid(row=1, column=2, pady=2)
        self.entries["Plot Wavenumber 1:"] = wn1_entry
        self.entries["Plot Wavenumber 2:"] = wn2_entry

        tk.Label(time_frame, text="Time Range Start / End (s):", bg="#f0f0f0", font=label_font).grid(row=2, column=0, sticky=tk.W, pady=2)
        t_start_entry = tk.Entry(time_frame, font=entry_font, width=10)
        t_start_entry.grid(row=2, column=1, pady=2)
        t_end_entry = tk.Entry(time_frame, font=entry_font, width=10)
        t_end_entry.grid(row=2, column=2, pady=2)
        self.entries["Time Range Start (s):"] = t_start_entry
        self.entries["Time Range End (s):"] = t_end_entry

        button_row = tk.Frame(time_frame, bg="#f0f0f0")
        button_row.grid(row=3, column=0, columnspan=3, pady=5)
        tk.Button(button_row, text="Plot Time Profiles", font=label_font, command=self.plot_time_profiles).pack(side="left", padx=5)
        tk.Button(button_row, text="Export Time Plot", font=label_font, command=self.export_time_plot).pack(side="left", padx=5)
        tk.Button(button_row, text="Plot Raw Data", font=label_font, command=self.plot_raw_data).pack(side="left", padx=5)

        # --- UV-Vis Compilation Section ---
        compile_frame = tk.LabelFrame(right_panel, text="Data Compilation", bg="#f0f0f0", padx=10, pady=8, font=title_font)
        compile_frame.pack(fill="x")
        
        tk.Button(compile_frame, text="UV-Vis Data", font=label_font, command=self.run_uvvis_compilation).pack(pady=5, fill="x")
        
        # PSD Angle Section
        angle_frame = tk.LabelFrame(main_frame, text="PSD angle", bg="#f5f5f5", padx=10, pady=10, font=title_font)
        angle_frame.grid(row=0, column=2, sticky="nsew", padx=5, pady=5)

        for i in range(4):
            r = (i // 2)
            c = (i % 2) * 2
            label = f"Wavenumber {i+1}:"
            tk.Label(angle_frame, text=label, bg="#f5f5f5", font=label_font).grid(row=r, column=c, sticky=tk.W, pady=2)
            entry = tk.Entry(angle_frame, font=entry_font, width=10)
            entry.grid(row=r, column=c+1, pady=2)
            self.entries[label] = entry

        bottom_controls = tk.Frame(angle_frame, bg="#f5f5f5")
        bottom_controls.grid(row=2, column=0, columnspan=4, pady=5)
        self.normalize_flag = tk.BooleanVar()
        tk.Checkbutton(bottom_controls, text="Normalize All", variable=self.normalize_flag, bg="#f5f5f5", font=label_font).pack(side="left", padx=5)
        tk.Button(bottom_controls, text="Plot Phase vs Intensity", font=label_font, command=self.plot_angle_dependence).pack(side="left", padx=5)

        tk.Label(angle_frame, text="PSD Heatmap", bg="#f5f5f5", font=("Helvetica", 12, "bold")).grid(row=3, column=0, columnspan=4, pady=(10, 2), sticky="w")

        # --- Range Selection for Heatmap ---
        range_frame = tk.Frame(angle_frame, bg="#f5f5f5")
        range_frame.grid(row=4, column=0, columnspan=4, pady=(5, 5))
        
        tk.Label(range_frame, text="Wavenumber Range:", bg="#f5f5f5", font=label_font).grid(row=0, column=0, sticky="w")
        self.entries["Heatmap WN Start"] = tk.Entry(range_frame, font=entry_font, width=10)
        self.entries["Heatmap WN Start"].grid(row=0, column=1)
        self.entries["Heatmap WN End"] = tk.Entry(range_frame, font=entry_font, width=10)
        self.entries["Heatmap WN End"].grid(row=0, column=2)
        
        tk.Label(range_frame, text="Phase Range (°):", bg="#f5f5f5", font=label_font).grid(row=1, column=0, sticky="w")
        self.entries["Heatmap Phase Start"] = tk.Entry(range_frame, font=entry_font, width=10)
        self.entries["Heatmap Phase Start"].grid(row=1, column=1)
        self.entries["Heatmap Phase End"] = tk.Entry(range_frame, font=entry_font, width=10)
        self.entries["Heatmap Phase End"].grid(row=1, column=2)
        tk.Button(angle_frame, text="Show PSD Heatmap", font=label_font, command=self.show_psd_heatmap).grid(row=6, column=0, columnspan=3, pady=5)
        tk.Button(angle_frame, text="Save Heatmap", font=label_font, command=self.save_psd_heatmap).grid(row=6, column=2, columnspan=2, pady=5)

        self.fig, self.ax = plt.subplots(figsize=(10, 3.2))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.root)
        self.canvas.get_tk_widget().grid(row=1, column=0, columnspan=3, sticky="nsew")

        toolbar = NavigationToolbar2Tk(self.canvas, self.root, pack_toolbar=False)
        toolbar.grid(row=2, column=0, columnspan=3, sticky='ew')
        toolbar.update()

        # Documentation and Help Section
        doc_frame = tk.LabelFrame(main_frame, text="Help & Theory", bg="#eeeeee", padx=10, pady=10, font=title_font)
        doc_frame.grid(row=0, column=3, sticky="nsew", padx=5, pady=5)

        tk.Button(doc_frame, text="📘 Open User Guide", font=label_font, command=self.open_user_guide).pack(pady=5, fill="x")
        tk.Button(doc_frame, text="🧠 Theory & Assumptions", font=label_font, command=self.open_theory_doc).pack(pady=5, fill="x")

        footer_label = tk.Label(doc_frame,
            text="A software developed by Leonardo Sousa with the help of ChatGPT\nFor enquiries, contact: l.sousa@ucl.ac.uk",
            font=("Helvetica", 9, "italic"),
            bg="#dcdcdc", justify="left", anchor="w")
        footer_label.pack(side="bottom", pady=10, padx=5)

    def load_data(self):
        file_path = filedialog.askopenfilename(title="Select Data File", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if file_path:
            try:
                self.raw_data = np.genfromtxt(file_path, delimiter="\t")
                self.wavenumbers = self.raw_data[:, 0]
                self.spectra = self.raw_data[:, 1:]
                self.data_loaded = True
                messagebox.showinfo("Data Loaded", "Data successfully loaded.")
            except Exception as e:
                messagebox.showerror("Load Error", f"Could not load file:\n{e}")

    def plot_time_profiles(self):
        if not self.data_loaded:
            messagebox.showwarning("Warning", "Load data first!")
            return
        try:
            wn1 = float(self.entries["Plot Wavenumber 1:"].get())
            wn2 = float(self.entries["Plot Wavenumber 2:"].get())
            t_start = float(self.entries["Time Range Start (s):"].get())
            t_end = float(self.entries["Time Range End (s):"].get())
            total_time = float(self.entries["Total Time (s):"].get())
            total_points = self.spectra.shape[1]
            time_axis = np.linspace(0, total_time, total_points)

            idx1 = np.argmin(np.abs(self.wavenumbers - wn1))
            idx2 = np.argmin(np.abs(self.wavenumbers - wn2))

            mask = (time_axis >= t_start) & (time_axis <= t_end)

            self.ax.clear()
            self.ax.plot(time_axis[mask], self.spectra[idx1, mask], label=f"{wn1:.1f} cm⁻¹")
            self.ax.plot(time_axis[mask], self.spectra[idx2, mask], label=f"{wn2:.1f} cm⁻¹", color='red')

            # Store time and intensity data for export
            self.spectrum_out = pd.DataFrame({
                "Time (s)": time_axis[mask],
                f"Intensity ({wn1:.1f} cm⁻¹)": self.spectra[idx1, mask],
                f"Intensity ({wn2:.1f} cm⁻¹)": self.spectra[idx2, mask]
            })

            self.ax.set_xlabel("Time (s)")
            self.ax.set_ylabel("Intensity")
            self.ax.set_title("Time Profile of Selected Wavenumbers")
            self.ax.legend()
            self.canvas.draw()

        except Exception as e:
            messagebox.showerror("Plot Error", str(e))


    def export_time_plot(self):
        if not hasattr(self, 'spectrum_out'): return
        filename = filedialog.asksaveasfilename(defaultextension=".csv")
        if filename:
            try:
                self.spectrum_out.to_csv(filename, index=False)
                messagebox.showinfo("Success", "Time profile exported successfully.")
            except Exception as e:
                messagebox.showerror("Save Error", str(e))

    def plot_angle_dependence(self):
        if not hasattr(self, 'spectra') or not hasattr(self, 'wavenumbers'):
            messagebox.showwarning("Warning", "Load data and process PSD first!")
            return

        try:
            modulation_period = float(self.entries["Modulation Period (s):"].get())
            sample_periods = int(self.entries["Sample Periods:"].get())
            start_phase = float(self.entries["Start Phase (degrees):"].get())
            end_phase = float(self.entries["End Phase (degrees):"].get())
            phase_step = float(self.entries["Phase Step (degrees):"].get())

            phase_angles = np.arange(start_phase, end_phase + phase_step, phase_step)
            omega = 2 * np.pi / modulation_period
            total_points = self.spectra.shape[1]
            time_axis = np.linspace(0, modulation_period * sample_periods, total_points)

            psd_spectra = np.zeros((self.spectra.shape[0], len(phase_angles)))
            for i, phi in enumerate(phase_angles):
                for j in range(self.spectra.shape[0]):
                    psd_spectra[j, i] = integrate.trapezoid(
                        self.spectra[j, :] * np.sin(omega * time_axis + np.deg2rad(phi)),
                        time_axis
                    )

            self.ax.clear()
            colors = ['blue', 'red', 'green', 'purple']

            for i in range(4):
                key = f"Wavenumber {i+1}:"
                entry = self.entries[key].get()
                if entry.strip() == "":
                    continue
                wn = float(entry)
                idx = np.argmin(np.abs(self.wavenumbers - wn))
                intensity = psd_spectra[idx, :]

                if self.normalize_flag.get():
                    intensity = intensity / np.max(np.abs(intensity))

                self.ax.plot(phase_angles, intensity, label=f"{wn:.1f} cm⁻¹", color=colors[i % len(colors)])

            self.ax.set_xlabel("Phase Angle (degrees)")
            self.ax.set_ylabel("Intensity")
            self.ax.set_title("PSD Intensity vs Phase Angle")
            self.ax.legend()
            self.canvas.draw()

        except Exception as e:
            messagebox.showerror("Plot Error", str(e))
            
    def show_psd_heatmap(self):
        if not hasattr(self, 'psd_output'):
            messagebox.showwarning("Warning", "No PSD data available. Run 'Process & Plot PSD' first.")
            return
    
        try:
            import matplotlib.pyplot as plt
            
            # Extract ranges
            wn_start = float(self.entries["Heatmap WN Start"].get())
            wn_end = float(self.entries["Heatmap WN End"].get())
            phase_start = int(self.entries["Heatmap Phase Start"].get())
            phase_end = int(self.entries["Heatmap Phase End"].get())
            
            # Filter PSD data
            heatmap_data = self.psd_output.loc[
                (self.psd_output.index >= wn_start) & (self.psd_output.index <= wn_end),
                [col for col in self.psd_output.columns if phase_start <= int(col.strip("°")) <= phase_end]
            ]
            
            # Transpose to have wavenumbers on x-axis and phases on y-axis
            heatmap_array = heatmap_data.T.values  # shape: (phases, wavenumbers)
            wn_values = heatmap_data.index.values  # x-axis
            phase_values = [int(col.strip("°")) for col in heatmap_data.columns]  # y-axis
            
            # Plot
            fig, ax = plt.subplots(figsize=(8, 5))
            cax = ax.imshow(
                heatmap_array,
                aspect='auto',
                origin='lower',
                extent=[wn_values[0], wn_values[-1], phase_values[0], phase_values[-1]],
                cmap='viridis'
            )
            
            ax.set_xlabel("Wavenumber (cm⁻¹)")
            ax.set_ylabel("Phase Angle (°)")
            ax.set_title("PSD Heatmap")
            fig.colorbar(cax, label='Intensity')
            plt.tight_layout()
            plt.show()

        except Exception as e:
            messagebox.showerror("Plot Error", f"Could not display heatmap:\n{str(e)}")

    def save_psd_heatmap(self):
        try:
            wn_start = float(self.entries["Heatmap WN Start"].get())
            wn_end = float(self.entries["Heatmap WN End"].get())
            phase_start = float(self.entries["Heatmap Phase Start"].get())
            phase_end = float(self.entries["Heatmap Phase End"].get())
    
            # Extract columns within the phase range, removing "°" for comparison
            cols_to_export = [
                col for col in self.psd_output.columns
                if phase_start <= float(col.replace("°", "")) <= phase_end
            ]
    
            # Extract wavenumber range and selected columns, and clean column names
            heatmap_data = self.psd_output.loc[
                (self.psd_output.index >= wn_start) & (self.psd_output.index <= wn_end),
                cols_to_export
            ].copy()
    
            heatmap_data.columns = [col.replace("°", "") for col in heatmap_data.columns]
            heatmap_data.index.name = "Wavenumber (cm⁻¹)"
    
            # Save file
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("Text files", "*.txt")]
            )
    
            if file_path:
                heatmap_data.to_csv(file_path)
                messagebox.showinfo("Export Successful", "Heatmap saved successfully.")
    
        except Exception as e:
            messagebox.showerror("Export Error", str(e))


    def plot_raw_data(self):
        if not self.data_loaded:
            messagebox.showwarning("Warning", "Load data first!")
            return
        try:
            self.ax.clear()
            for i in range(0, self.spectra.shape[1], max(1, self.spectra.shape[1] // 100)):
                self.ax.plot(self.wavenumbers, self.spectra[:, i], label=f"t{i}")
            self.ax.set_xlabel("Wavenumber (cm⁻¹)")
            self.ax.set_ylabel("Intensity")
            self.ax.set_title("Raw Spectral Data")
            self.ax.legend()
            self.canvas.draw()
        except Exception as e:
            messagebox.showerror("Plot Error", str(e))

    def save_data(self):
        if not hasattr(self, 'psd_output'):
            messagebox.showerror("Error", "No PSD data to save. Run 'Process & Plot PSD' first.")
            return
    
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if file_path:
            try:
                self.psd_output.to_csv(file_path, sep='\t')
                messagebox.showinfo("Success", f"PSD data saved to:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Save Error", f"Could not save file:\n{e}")


    def open_user_guide(self):
        self._open_external_file("user_guide.pdf")  # Replace with actual path if needed
    
    def open_theory_doc(self):
        self._open_external_file("theory_and_assumptions.pdf")  # Replace with actual path if needed
    
    def _open_external_file(self, filepath):
        try:
            if os.name == 'nt':  # Windows
                os.startfile(filepath)
            elif os.name == 'posix':  # macOS, Linux
                subprocess.call(['xdg-open', filepath])
            else:
                messagebox.showerror("Error", "Unsupported operating system.")
        except Exception as e:
            messagebox.showerror("File Open Error", f"Could not open file:\n{e}")


    def process_and_plot(self):
        if not self.data_loaded:
            messagebox.showwarning("Warning", "Load data first!")
            return
    
        try:
            modulation_period = float(self.entries["Modulation Period (s):"].get())
            dummy_periods = int(self.entries["Dummy Periods:"].get())
            sample_periods = int(self.entries["Sample Periods:"].get())
            start_phase = float(self.entries["Start Phase (degrees):"].get())
            end_phase = float(self.entries["End Phase (degrees):"].get())
            phase_step = float(self.entries["Phase Step (degrees):"].get())
            wn_start = float(self.entries["Wavenumber Range Start:"].get())
            wn_end = float(self.entries["Wavenumber Range End:"].get())
            total_time = float(self.entries["Total Time (s):"].get())
    
            total_points = self.spectra.shape[1]
            time_axis = np.linspace(0, total_time, total_points)
            dt = time_axis[1] - time_axis[0]
            points_per_period = int(round(modulation_period / dt))
    
            # Determine usable time and data after trimming
            start_idx = dummy_periods * points_per_period
            end_idx = (dummy_periods + sample_periods) * points_per_period
            if end_idx > total_points:
                messagebox.showerror("Error", "Requested sample + dummy periods exceed dataset duration.")
                return
    
            wn_mask = (self.wavenumbers >= wn_start) & (self.wavenumbers <= wn_end)
            wn_selected = self.wavenumbers[wn_mask]
            spec_selected = self.spectra[wn_mask, start_idx:end_idx]
            time_trimmed = time_axis[start_idx:end_idx]
    
            # Determine number of usable full cycles
            usable_points = spec_selected.shape[1]
            usable_cycles = usable_points // points_per_period
            usable_points = usable_cycles * points_per_period
    
            # Truncate to exact number of usable points
            spec_selected = spec_selected[:, :usable_points]
            time_trimmed = time_trimmed[:usable_points]
    
            # Reshape spectra and time to (num_wavenumbers, cycles, points_per_period)
            reshaped_spec = spec_selected.reshape(len(wn_selected), usable_cycles, points_per_period)
            reshaped_time = time_trimmed.reshape(usable_cycles, points_per_period)
    
            k_harmonic = int(self.entries["Harmonic Order (k):"].get())
            omega = k_harmonic * 2 * np.pi / modulation_period

            phase_angles = np.arange(start_phase, end_phase + phase_step, phase_step)
            psd_spectra = np.zeros((len(wn_selected), len(phase_angles)))
    
            for i, phi in enumerate(phase_angles):
                phi_rad = np.deg2rad(phi)
                sine_ref = np.sin(omega * reshaped_time + phi_rad)  # shape: (cycles, points_per_period)
                sine_ref /= np.max(np.abs(sine_ref), axis=1, keepdims=True)  # normalize
    
                for j in range(len(wn_selected)):
                    signal = reshaped_spec[j, :, :]  # shape: (cycles, points_per_period)
                    # Optionally subtract mean per cycle
                    signal -= np.mean(signal, axis=1, keepdims=True)
                    product = signal * sine_ref
                    integrals = integrate.trapezoid(product, reshaped_time, axis=1)  # integrate over time for each cycle
                    psd_spectra[j, i] = np.mean(integrals)
    
            self.psd_output = pd.DataFrame(psd_spectra, index=wn_selected, columns=[f"{int(p)}°" for p in phase_angles])
            self.psd_output.index.name = "Wavenumber (cm⁻¹)"
    
            self.ax.clear()
            for i, phi in enumerate(phase_angles):
                self.ax.plot(wn_selected, psd_spectra[:, i], label=f"{int(phi)}°")
            self.ax.set_xlabel("Wavenumber (cm⁻¹)")
            self.ax.set_ylabel("PSD Intensity")
            self.ax.set_title("Phase-Resolved PSD Spectra")
            self.ax.legend()
            self.canvas.draw()
    
        except Exception as e:
            messagebox.showerror("PSD Error", str(e))

    def run_uvvis_compilation(self):
        import subprocess
        import tempfile
        import textwrap
    
        try:
            uvvis_script = textwrap.dedent("""\
                import sys
                from pathlib import Path
                from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox
                import numpy as np
    
                def _is_float_token(tok: str) -> bool:
                    try:
                        float(tok)
                        return True
                    except Exception:
                        return False
    
                def _find_first_numeric_line(lines, min_cols=2):
                    for i, ln in enumerate(lines):
                        ln = ln.strip()
                        if not ln:
                            continue
                        ln = ln.replace(",", " ").replace(";", " ")
                        toks = [t for t in ln.split() if t]
                        if len(toks) >= min_cols and _is_float_token(toks[0]) and _is_float_token(toks[1]):
                            return i
                    return None
    
                def _parse_xy_from_lines(lines, start_idx):
                    xs, ys = [], []
                    for ln in lines[start_idx:]:
                        s = ln.strip()
                        if not s:
                            continue
                        s = s.replace(",", " ").replace(";", " ")
                        arr = np.fromstring(s, sep=" ")
                        if arr.size < 2:
                            continue
                        x, y = arr[0], arr[-1]
                        if np.isfinite(x) and np.isfinite(y):
                            xs.append(x); ys.append(y)
                    if not xs:
                        return None, None
                    x = np.asarray(xs, dtype=float)
                    y = np.asarray(ys, dtype=float)
                    order = np.argsort(x)
                    x = x[order]; y = y[order]
                    uniq, idx, counts = np.unique(x, return_index=True, return_counts=True)
                    y_sum = np.add.reduceat(y, idx)
                    y_avg = y_sum / counts
                    return uniq, y_avg
    
                def read_spectrum_xy(path: Path):
                    try:
                        text = path.read_text(encoding="utf-8", errors="ignore")
                    except UnicodeDecodeError:
                        text = path.read_text(encoding="latin-1", errors="ignore")
                    lines = text.splitlines()
                    start = _find_first_numeric_line(lines, min_cols=2)
                    if start is None:
                        return None, None
                    return _parse_xy_from_lines(lines, start)
    
                def compile_folder_fast(folder: Path):
                    exts = {".txt", ".dat", ".csv", ".asc"}
                    files = sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in exts])
                    if not files:
                        raise FileNotFoundError("No spectra found.")
                    x_list, y_list = [], []
                    for f in files:
                        X, Y = read_spectrum_xy(f)
                        if X is None or X.size == 0:
                            continue
                        x_list.append(X); y_list.append(Y)
                    if not x_list:
                        raise ValueError("No readable spectra.")
                    X_common = x_list[0]
                    for X in x_list[1:]:
                        X_common = np.intersect1d(X_common, X)
                        if X_common.size == 0:
                            raise ValueError("No common X values.")
                    Y_stack = []
                    for X, Y in zip(x_list, y_list):
                        idx = np.searchsorted(X, X_common)
                        valid = (idx >= 0) & (idx < X.size) & (X[idx] == X_common)
                        col = np.full(X_common.shape, np.nan, dtype=float)
                        col[valid] = Y[idx[valid]]
                        Y_stack.append(col)
                    Y_stack = np.column_stack(Y_stack)
                    return X_common, Y_stack
    
                def choose_folder() -> Path:
                    app = QApplication(sys.argv)
                    folder = QFileDialog.getExistingDirectory(None, "Select Folder with UV-Vis Spectra")
                    return Path(folder) if folder else None
    
                def alert(title: str, text: str, critical=False):
                    app = QApplication.instance() or QApplication(sys.argv)
                    box = QMessageBox()
                    box.setWindowTitle(title)
                    box.setText(text)
                    box.setIcon(QMessageBox.Critical if critical else QMessageBox.Information)
                    box.exec_()
    
                def main():
                    folder = choose_folder()
                    if not folder:
                        return
                    try:
                        X, Y_stack = compile_folder_fast(folder)
                        n_cols = Y_stack.shape[1]
                        header = ["X"] + [f"Y{i+1}" for i in range(n_cols)]
                        out_path = folder / "compiled_spectra.txt"
                        out = np.column_stack([X, Y_stack])
                        np.savetxt(out_path, out, delimiter="\\t", header="\\t".join(header), comments="", fmt="%.10g")
                        alert("Success", f"Compiled spectra saved to:\\n{out_path}")
                    except Exception as e:
                        alert("Error", f"{type(e).__name__}: {e}", critical=True)
                        raise
    
                if __name__ == "__main__":
                    main()
            """)
    
            with tempfile.NamedTemporaryFile("w", delete=False, suffix=".py") as f:
                f.write(uvvis_script)
                temp_path = f.name
    
            subprocess.Popen(["python", temp_path])
    
        except Exception as e:
            messagebox.showerror("UV-Vis Compilation Error", str(e))


if __name__ == "__main__":
    PSDApp()

