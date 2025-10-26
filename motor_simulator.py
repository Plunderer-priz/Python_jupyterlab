"""
Advanced Motor Simulator - Professional Edition
Comprehensive motor performance analysis tool for DC and AC motors
Author: Motor Simulation System
Version: 1.0
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import math
import json
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import numpy as np

# Optional imports
try:
    from fpdf import FPDF
    FPDF_AVAILABLE = True
except ImportError:
    FPDF_AVAILABLE = False
    print("Warning: FPDF not installed. PDF export will be unavailable.")

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False
    print("Warning: Pandas not installed. Excel export will be unavailable.")


class MotorSimulator:
    """Advanced Motor Simulation Program with accurate calculations and cyberpunk aesthetics"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("⚡ Advanced Motor Simulator - Professional Edition")
        self.root.geometry("1500x950")
        
        # Modern dark theme (Cyberpunk aesthetics) colors
        self.colors = {
            'bg_primary': '#0a0e27',      # Deep space blue
            'bg_secondary': '#1a1f3a',    # Dark navy
            'bg_tertiary': '#252b48',     # Lighter navy
            'accent_cyan': '#00d9ff',     # Electric cyan
            'accent_purple': '#b794f6',   # Soft purple
            'accent_green': '#00ff88',    # Neon green
            'accent_orange': '#ff6b35',   # Vibrant orange
            'text_primary': '#ffffff',    # White
            'text_secondary': '#a0aec0',  # Light gray
            'success': '#10b981',         # Green
            'warning': '#fbbf24',         # Yellow
            'error': '#ef4444',           # Red
            'chart_bg': '#0f1729',        # Darker for charts
        }
        
        self.root.configure(bg=self.colors['bg_primary'])
        
        # Motor parameters storage
        self.motor_type = tk.StringVar(value="DC")
        self.simulation_history = []
        
        # Style configuration
        self.setup_styles()
        
        # Create main interface
        self.create_interface()
        
    def setup_styles(self):
        """Configure ttk styles with modern aesthetics"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Frame styles
        style.configure('TFrame', background=self.colors['bg_primary'])
        style.configure('Card.TFrame', background=self.colors['bg_secondary'], relief=tk.RAISED)
        
        # Label styles
        style.configure('TLabel', 
                       background=self.colors['bg_primary'], 
                       foreground=self.colors['text_primary'], 
                       font=('Helvetica', 10))
        
        style.configure('Title.TLabel', 
                       font=('Helvetica', 20, 'bold'), 
                       foreground=self.colors['accent_cyan'],
                       background=self.colors['bg_primary'])
        
        style.configure('Subtitle.TLabel',
                       font=('Helvetica', 11),
                       foreground=self.colors['text_secondary'],
                       background=self.colors['bg_primary'])
        
        style.configure('Header.TLabel',
                       font=('Helvetica', 12, 'bold'),
                       foreground=self.colors['accent_purple'],
                       background=self.colors['bg_secondary'])
        
        # LabelFrame styles
        style.configure('TLabelframe', 
                       background=self.colors['bg_secondary'],
                       bordercolor=self.colors['accent_cyan'],
                       borderwidth=2)
        style.configure('TLabelframe.Label',
                       background=self.colors['bg_secondary'],
                       foreground=self.colors['accent_cyan'],
                       font=('Helvetica', 11, 'bold'))
        
        # Radiobutton styles
        style.configure('TRadiobutton', 
                       background=self.colors['bg_secondary'], 
                       foreground=self.colors['text_primary'], 
                       font=('Helvetica', 10))
        
        # Entry styles
        style.configure('TEntry',
                       fieldbackground=self.colors['bg_tertiary'],
                       foreground=self.colors['text_primary'],
                       bordercolor=self.colors['accent_cyan'])
        
        # Scrollbar styles
        style.configure('Vertical.TScrollbar',
                       background=self.colors['bg_tertiary'],
                       troughcolor=self.colors['bg_primary'],
                       bordercolor=self.colors['bg_secondary'],
                       arrowcolor=self.colors['accent_cyan'])
        
    def create_interface(self):
        """Create main user interface with stunning aesthetics"""
        # Header with gradient effect simulation
        header_frame = tk.Frame(self.root, bg=self.colors['bg_primary'], height=100)
        header_frame.pack(fill=tk.X, padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        # Title with glow effect
        title_label = tk.Label(header_frame, 
                              text="⚡ ADVANCED MOTOR SIMULATOR ⚡", 
                              bg=self.colors['bg_primary'],
                              fg=self.colors['accent_cyan'],
                              font=('Helvetica', 24, 'bold'))
        title_label.pack(pady=(20, 5))
        
        subtitle_label = tk.Label(header_frame, 
                                 text="Professional Motor Performance Analysis & Virtual Testing Platform",
                                 bg=self.colors['bg_primary'],
                                 fg=self.colors['text_secondary'],
                                 font=('Helvetica', 11, 'italic'))
        subtitle_label.pack()
        
        # Separator line
        separator = tk.Frame(self.root, bg=self.colors['accent_cyan'], height=3)
        separator.pack(fill=tk.X)
        
        # Control panel at top
        self.create_control_panel()
        
        # Main container with three columns
        main_container = tk.Frame(self.root, bg=self.colors['bg_primary'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        # Configure grid weights
        main_container.columnconfigure(0, weight=1, minsize=320)
        main_container.columnconfigure(1, weight=3, minsize=650)
        main_container.columnconfigure(2, weight=1, minsize=320)
        main_container.rowconfigure(0, weight=1)
        
        # Left panel - Input parameters
        self.create_input_panel(main_container)
        
        # Center panel - Visualization
        self.create_visualization_panel(main_container)
        
        # Right panel - Results
        self.create_results_panel(main_container)
        
        # Status bar at bottom
        self.create_status_bar()

    def create_input_panel(self, parent):
        """Create stunning input parameters panel"""
        input_frame = ttk.LabelFrame(parent, text="⚙ MOTOR CONFIGURATION", padding=15)
        input_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 8))
        
        # Motor Type Selection with card design
        type_card = tk.Frame(input_frame, bg=self.colors['bg_tertiary'], relief=tk.RAISED, bd=2)
        type_card.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(type_card, text="Select Motor Type:", 
                 font=('Helvetica', 11, 'bold'),
                 background=self.colors['bg_tertiary'],
                 foreground=self.colors['accent_cyan']).pack(anchor=tk.W, padx=10, pady=(10, 5))
        
        dc_radio = ttk.Radiobutton(type_card, text="⚡ DC Motor", variable=self.motor_type, 
                                   value="DC", command=self.update_parameter_visibility)
        dc_radio.pack(anchor=tk.W, padx=25, pady=3)
        
        ac_radio = ttk.Radiobutton(type_card, text="⚡ AC Motor (3-Phase)", variable=self.motor_type, 
                                   value="AC", command=self.update_parameter_visibility)
        ac_radio.pack(anchor=tk.W, padx=25, pady=(3, 10))
        
        # Scrollable frame for parameters
        canvas = tk.Canvas(input_frame, bg=self.colors['bg_secondary'], 
                          highlightthickness=0, height=500)
        scrollbar = ttk.Scrollbar(input_frame, orient="vertical", command=canvas.yview)
        self.params_frame = tk.Frame(canvas, bg=self.colors['bg_secondary'])
        
        self.params_frame.bind("<Configure>", 
                              lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.params_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True, pady=10)
        scrollbar.pack(side="right", fill="y")
        
        # Create parameter inputs
        self.create_parameter_inputs()

    def create_parameter_inputs(self):
        """Create all parameter input fields"""
        self.param_vars = {}
        
        # Common parameters for both DC and AC
        common_params = [
            ("Electrical Parameters", [
                ("rated_voltage", "Rated Voltage (V)", "220", "Nominal operating voltage"),
                ("rated_current", "Rated Current (A)", "10", "Nominal operating current"),
                ("rated_power", "Rated Power (W)", "2200", "Rated output power"),
            ]),
            ("Mechanical Parameters", [
                ("rated_speed", "Rated Speed (RPM)", "1500", "Nominal operating speed"),
                ("rated_torque", "Rated Torque (Nm)", "14.0", "Nominal torque at rated speed"),
                ("load_torque", "Load Torque (Nm)", "10.0", "External load torque"),
            ]),
            ("Physical Parameters", [
                ("frame_size", "Frame Size (mm)", "100", "Motor frame dimensions"),
                ("rotor_inertia", "Rotor Inertia (kg·m²)", "0.001", "Moment of inertia"),
                ("weight", "Weight (kg)", "15", "Total motor weight"),
            ]),
            ("Thermal Parameters", [
                ("insulation_class", "Insulation Class", "F", "Temperature class (A/B/F/H)"),
                ("duty_cycle", "Duty Cycle (%)", "100", "Continuous operation percentage"),
                ("ambient_temp", "Ambient Temp (°C)", "40", "Operating environment temperature"),
            ])
        ]
        
        # DC-specific parameters
        dc_params = [
            ("DC Motor Parameters", [
                ("armature_resistance", "Armature Resistance (Ω)", "1.5", "Ra - Armature winding resistance"),
                ("field_resistance", "Field Resistance (Ω)", "100", "Rf - Field winding resistance"),
                ("back_emf_constant", "Back EMF Constant (V/rad/s)", "0.8", "Ke - Voltage constant"),
                ("torque_constant", "Torque Constant (Nm/A)", "0.8", "Kt - Torque per ampere"),
            ])
        ]
        
        # AC-specific parameters
        ac_params = [
            ("AC Motor Parameters", [
                ("frequency", "Frequency (Hz)", "50", "Supply frequency"),
                ("poles", "Number of Poles", "4", "Pole count (even number)"),
                ("stator_resistance", "Stator Resistance (Ω)", "1.2", "Per phase resistance"),
                ("rotor_resistance", "Rotor Resistance (Ω)", "0.8", "Referred to stator"),
                ("magnetizing_current", "Magnetizing Current (A)", "3.0", "No-load current"),
                ("slip", "Slip (%)", "3.0", "Speed difference percentage"),
            ])
        ]
        
        # Create input fields
        row = 0
        for section_title, params in common_params:
            self.add_parameter_section(section_title, params, row)
            row += len(params) + 1
        
        # Store DC parameters
        self.dc_param_widgets = []
        for section_title, params in dc_params:
            widgets = self.add_parameter_section(section_title, params, row)
            self.dc_param_widgets.extend(widgets)
            row += len(params) + 1
        
        # Store AC parameters
        self.ac_param_widgets = []
        for section_title, params in ac_params:
            widgets = self.add_parameter_section(section_title, params, row)
            self.ac_param_widgets.extend(widgets)
            row += len(params) + 1
        
        self.update_parameter_visibility()

    def add_parameter_section(self, title, params, start_row):
        """Add a section of parameters with modern styling"""
        widgets = []
        
        # Section header with accent color
        header = tk.Label(self.params_frame, text=title, 
                         font=('Helvetica', 11, 'bold'),
                         bg=self.colors['bg_secondary'],
                         fg=self.colors['accent_orange'],
                         anchor=tk.W)
        header.grid(row=start_row, column=0, columnspan=2, sticky=tk.W+tk.E, 
                   pady=(12, 8), padx=5)
        widgets.append(header)
        
        # Add separator line
        sep = tk.Frame(self.params_frame, bg=self.colors['accent_orange'], height=2)
        sep.grid(row=start_row, column=0, columnspan=2, sticky=tk.W+tk.E, pady=(0, 8))
        widgets.append(sep)
        
        # Parameters
        for i, (key, label, default, tooltip) in enumerate(params, start=1):
            row = start_row + i
            
            lbl = tk.Label(self.params_frame, text=label + ":", 
                          bg=self.colors['bg_secondary'],
                          fg=self.colors['text_primary'],
                          font=('Helvetica', 9),
                          anchor=tk.W)
            lbl.grid(row=row, column=0, sticky=tk.W, padx=8, pady=4)
            widgets.append(lbl)
            
            var = tk.StringVar(value=default)
            self.param_vars[key] = var
            
            entry = tk.Entry(self.params_frame, textvariable=var, width=12,
                            bg=self.colors['bg_tertiary'], 
                            fg=self.colors['text_primary'],
                            font=('Helvetica', 10),
                            insertbackground=self.colors['accent_cyan'],
                            relief=tk.FLAT,
                            bd=2,
                            highlightthickness=1,
                            highlightcolor=self.colors['accent_cyan'],
                            highlightbackground=self.colors['bg_tertiary'])
            entry.grid(row=row, column=1, sticky=tk.EW, padx=8, pady=4)
            widgets.append(entry)
            
            # Tooltip
            self.create_tooltip(entry, tooltip)
        
        return widgets

    def create_tooltip(self, widget, text):
        """Create tooltip for widget"""
        def on_enter(e):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{e.x_root+10}+{e.y_root+10}")
            label = tk.Label(tooltip, text=text, background="#ffffe0", relief=tk.SOLID, 
                           borderwidth=1, font=('Segoe UI', 8))
            label.pack()
            widget.tooltip = tooltip
            
        def on_leave(e):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                
        widget.bind('<Enter>', on_enter)
        widget.bind('<Leave>', on_leave)

    def update_parameter_visibility(self):
        """Show/hide parameters based on motor type"""
        if self.motor_type.get() == "DC":
            for widget in self.dc_param_widgets:
                widget.grid()
            for widget in self.ac_param_widgets:
                widget.grid_remove()
        else:
            for widget in self.dc_param_widgets:
                widget.grid_remove()
            for widget in self.ac_param_widgets:
                widget.grid()

    def create_visualization_panel(self, parent):
        """Create visualization panel with graphs and cyberpunk styling"""
        viz_frame = ttk.LabelFrame(parent, text="📊 PERFORMANCE VISUALIZATION", padding=15)
        viz_frame.grid(row=0, column=1, sticky='nsew', padx=8)
        
        # Create matplotlib figure with dark theme
        self.fig = Figure(figsize=(10, 10), facecolor=self.colors['chart_bg'], 
                         edgecolor=self.colors['accent_cyan'], linewidth=2)
        
        # Create subplots
        self.ax1 = self.fig.add_subplot(3, 2, 1)
        self.ax2 = self.fig.add_subplot(3, 2, 2)
        self.ax3 = self.fig.add_subplot(3, 2, 3)
        self.ax4 = self.fig.add_subplot(3, 2, 4)
        self.ax5 = self.fig.add_subplot(3, 2, 5)
        self.ax6 = self.fig.add_subplot(3, 2, 6)
        
        self.axes = [self.ax1, self.ax2, self.ax3, self.ax4, self.ax5, self.ax6]
        
        # Style all axes with cyberpunk aesthetics
        for ax in self.axes:
            ax.set_facecolor(self.colors['bg_primary'])
            ax.tick_params(colors=self.colors['text_secondary'], labelsize=8)
            
            # Neon borders
            for spine in ax.spines.values():
                spine.set_color(self.colors['accent_cyan'])
                spine.set_linewidth(1.5)
            ax.xaxis.label.set_color(self.colors['text_primary'])
            ax.yaxis.label.set_color(self.colors['text_primary'])
            ax.title.set_color(self.colors['accent_purple'])
            ax.title.set_fontsize(10)
            ax.title.set_fontweight('bold')
        
        self.fig.tight_layout(pad=2.5)
        
        # Embed in tkinter
        self.canvas = FigureCanvasTkAgg(self.fig, master=viz_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def create_results_panel(self, parent):
        """Create stunning results display panel"""
        results_frame = ttk.LabelFrame(parent, text="📈 SIMULATION RESULTS", padding=15)
        results_frame.grid(row=0, column=2, sticky='nsew', padx=(8, 0))
        
        # Create text widget with scrollbar
        text_frame = tk.Frame(results_frame, bg=self.colors['bg_secondary'])
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.results_text = tk.Text(text_frame, wrap=tk.WORD, 
                                   bg=self.colors['bg_primary'], 
                                   fg=self.colors['accent_green'],
                                   font=('Courier New', 9),
                                   yscrollcommand=scrollbar.set,
                                   relief=tk.FLAT,
                                   bd=2,
                                   highlightthickness=2,
                                   highlightcolor=self.colors['accent_cyan'],
                                   highlightbackground=self.colors['bg_tertiary'],
                                   insertbackground=self.colors['accent_cyan'])
        self.results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.results_text.yview)
        
        # Configure tags for colored output with neon colors
        self.results_text.tag_config('header', 
                                    foreground=self.colors['accent_cyan'], 
                                    font=('Courier New', 10, 'bold'))
        self.results_text.tag_config('subheader', 
                                    foreground=self.colors['accent_purple'], 
                                    font=('Courier New', 9, 'bold'))
        self.results_text.tag_config('value', 
                                    foreground=self.colors['accent_green'])
        self.results_text.tag_config('warning', 
                                    foreground=self.colors['warning'])
        self.results_text.tag_config('error', 
                                    foreground=self.colors['error'])

    def create_status_bar(self):
        """Create status bar at bottom"""
        status_frame = tk.Frame(self.root, bg=self.colors['bg_secondary'], height=30)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        status_frame.pack_propagate(False)
        
        self.status_label = tk.Label(status_frame, 
                                     text="⚡ Ready to simulate | Powered by Python",
                                     bg=self.colors['bg_secondary'],
                                     fg=self.colors['text_secondary'],
                                     font=('Helvetica', 9),
                                     anchor=tk.W)
        self.status_label.pack(side=tk.LEFT, padx=15)
        
        # Version label
        version_label = tk.Label(status_frame,
                                text="v1.0 Professional",
                                bg=self.colors['bg_secondary'],
                                fg=self.colors['accent_purple'],
                                font=('Helvetica', 9, 'italic'),
                                anchor=tk.E)
        version_label.pack(side=tk.RIGHT, padx=15)

    def create_control_panel(self):
        """Create stunning control buttons panel"""
        control_frame = tk.Frame(self.root, bg=self.colors['bg_secondary'], height=100)
        control_frame.pack(fill=tk.X, padx=15, pady=10)
        control_frame.pack_propagate(False)
        
        # Title for control panel
        title = tk.Label(control_frame, 
                        text="⚙ CONTROL PANEL", 
                        bg=self.colors['bg_secondary'], 
                        fg=self.colors['accent_purple'], 
                        font=('Helvetica', 12, 'bold'))
        title.pack(pady=(8, 5))
        
        # Create a inner frame for buttons
        button_container = tk.Frame(control_frame, bg=self.colors['bg_secondary'])
        button_container.pack()
        
        # SIMULATE Button - Neon green with glow effect
        simulate_btn = tk.Button(button_container, 
                                text="▶ RUN SIMULATION", 
                                command=self.run_simulation,
                                bg=self.colors['accent_green'], 
                                fg='#000000', 
                                font=('Helvetica', 13, 'bold'),
                                activebackground='#00cc70',
                                activeforeground='#000000',
                                cursor='hand2',
                                width=22, 
                                height=2, 
                                relief=tk.RAISED, 
                                bd=4,
                                highlightthickness=2,
                                highlightbackground=self.colors['accent_green'],
                                highlightcolor=self.colors['accent_green'])
        simulate_btn.grid(row=0, column=0, padx=12, pady=5)
        
        save_btn = tk.Button(button_container, 
                            text="💾 SAVE", 
                            command=self.save_simulation,
                            bg=self.colors['bg_tertiary'], 
                            fg=self.colors['accent_cyan'], 
                            font=('Helvetica', 11, 'bold'),
                            activebackground=self.colors['accent_cyan'],
                            activeforeground='#000000',
                            cursor='hand2',
                            width=13, 
                            height=2, 
                            relief=tk.RAISED, 
                            bd=3,
                            highlightthickness=1,
                            highlightbackground=self.colors['accent_cyan'])
        save_btn.grid(row=0, column=1, padx=8, pady=5)
        
        load_btn = tk.Button(button_container, 
                            text="📂 LOAD", 
                            command=self.load_simulation,
                            bg=self.colors['bg_tertiary'], 
                            fg=self.colors['accent_purple'], 
                            font=('Helvetica', 11, 'bold'),
                            activebackground=self.colors['accent_purple'],
                            activeforeground='#000000',
                            cursor='hand2',
                            width=13, 
                            height=2, 
                            relief=tk.RAISED, 
                            bd=3,
                            highlightthickness=1,
                            highlightbackground=self.colors['accent_purple'])
        load_btn.grid(row=0, column=2, padx=8, pady=5)
        
        col_idx = 3
        if PANDAS_AVAILABLE:
            excel_btn = tk.Button(button_container, 
                                text="📊 EXCEL", 
                                command=self.export_excel,
                                bg=self.colors['bg_tertiary'], 
                                fg=self.colors['success'], 
                                font=('Helvetica', 11, 'bold'),
                                activebackground=self.colors['success'],
                                activeforeground='#000000',
                                cursor='hand2',
                                width=13, 
                                height=2, 
                                relief=tk.RAISED, 
                                bd=3,
                                highlightthickness=1,
                                highlightbackground=self.colors['success'])
            excel_btn.grid(row=0, column=col_idx, padx=8, pady=5)
            col_idx += 1
        
        if FPDF_AVAILABLE:
            pdf_btn = tk.Button(button_container, 
                               text="📄 PDF", 
                               command=self.export_pdf,
                               bg=self.colors['bg_tertiary'], 
                               fg=self.colors['error'], 
                               font=('Helvetica', 11, 'bold'),
                               activebackground=self.colors['error'],
                               activeforeground='#000000',
                               cursor='hand2',
                               width=13, 
                               height=2, 
                               relief=tk.RAISED, 
                               bd=3,
                               highlightthickness=1,
                               highlightbackground=self.colors['error'])
            pdf_btn.grid(row=0, column=col_idx, padx=8, pady=5)
            col_idx += 1
        
        reset_btn = tk.Button(button_container, 
                             text="🔄 RESET", 
                             command=self.reset_parameters,
                             bg=self.colors['bg_tertiary'], 
                             fg=self.colors['warning'], 
                             font=('Helvetica', 11, 'bold'),
                             activebackground=self.colors['warning'],
                             activeforeground='#000000',
                             cursor='hand2',
                             width=13, 
                             height=2, 
                             relief=tk.RAISED, 
                             bd=3,
                             highlightthickness=1,
                             highlightbackground=self.colors['warning'])
        reset_btn.grid(row=0, column=col_idx, padx=8, pady=5)
        
    def get_parameter_value(self, key, default=0):
        """Get parameter value with error handling"""
        try:
            return float(self.param_vars[key].get())
        except (ValueError, KeyError):
            return default
    
    def run_simulation(self):
        """Run motor simulation with accurate calculations"""
        try:
            # Clear previous results
            self.results_text.delete(1.0, tk.END)
            
            # Get parameters
            motor_type = self.motor_type.get()
            
            # Calculate based on motor type
            if motor_type == "DC":
                results = self.simulate_dc_motor()
            else:
                results = self.simulate_ac_motor()
            
            # Display results
            self.display_results(results)
            
            # Update visualizations
            self.update_visualizations(results)
            
            # Store in history
            results['timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            results['motor_type'] = motor_type
            self.simulation_history.append(results)
            
        except Exception as e:
            messagebox.showerror("Simulation Error", f"Error during simulation:\n{str(e)}")
    
    def simulate_dc_motor(self):
        """Simulate DC motor with accurate calculations"""
        # Get parameters
        V = self.get_parameter_value('rated_voltage', 220)
        I_rated = self.get_parameter_value('rated_current', 10)
        Ra = self.get_parameter_value('armature_resistance', 1.5)
        Rf = self.get_parameter_value('field_resistance', 100)
        Ke = self.get_parameter_value('back_emf_constant', 0.8)
        Kt = self.get_parameter_value('torque_constant', 0.8)
        T_load = self.get_parameter_value('load_torque', 10)
        N_rated = self.get_parameter_value('rated_speed', 1500)
        J = self.get_parameter_value('rotor_inertia', 0.001)
        
        # Calculate field current
        I_field = V / Rf
        
        # Calculate armature voltage
        V_armature = V
        
        # Calculate back EMF at rated speed
        omega_rated = N_rated * 2 * math.pi / 60  # rad/s
        Eb_rated = Ke * omega_rated
        
        # Calculate armature current at rated condition
        I_armature = (V_armature - Eb_rated) / Ra
        
        # Calculate developed torque
        T_developed = Kt * I_armature
        
        # Calculate actual speed under load
        # T_developed = T_load + friction losses (simplified)
        # Assuming 5% mechanical losses
        T_friction = T_developed * 0.05
        T_shaft = T_developed - T_friction
        
        # Speed calculation considering load
        speed_reduction_factor = T_load / T_developed if T_developed > 0 else 0
        N_actual = N_rated * (1 - speed_reduction_factor * 0.1)
        omega_actual = N_actual * 2 * math.pi / 60
        
        # Power calculations
        P_input = V * I_armature
        P_armature_loss = I_armature ** 2 * Ra
        P_field_loss = I_field ** 2 * Rf
        P_mechanical = T_shaft * omega_actual
        P_output = T_load * omega_actual
        
        # Efficiency
        efficiency = (P_output / P_input * 100) if P_input > 0 else 0
        
        # Core losses (approximation based on speed)
        P_core_loss = 0.02 * P_input  # 2% approximation
        
        # Total losses
        P_total_loss = P_armature_loss + P_field_loss + P_core_loss + (T_friction * omega_actual)
        
        # Temperature rise estimation (simplified)
        thermal_resistance = 5  # °C/W (typical for small motors)
        temp_rise = P_total_loss * thermal_resistance
        ambient_temp = self.get_parameter_value('ambient_temp', 40)
        operating_temp = ambient_temp + temp_rise
        
        # Current density
        # Assuming typical wire gauge
        conductor_area = 0.0001  # m² (example)
        current_density = I_armature / conductor_area if conductor_area > 0 else 0
        
        # Performance curves data (speed vs torque)
        speed_range = np.linspace(0, N_rated * 1.2, 50)
        torque_curve = []
        current_curve = []
        power_curve = []
        efficiency_curve = []
        
        for n in speed_range:
            omega = n * 2 * math.pi / 60
            Eb = Ke * omega
            Ia = max(0, (V_armature - Eb) / Ra)
            T = Kt * Ia
            P_out = T * omega
            P_in = V * Ia
            eff = (P_out / P_in * 100) if P_in > 0 else 0
            
            torque_curve.append(T)
            current_curve.append(Ia)
            power_curve.append(P_out)
            efficiency_curve.append(min(eff, 100))
        
        results = {
            'motor_type': 'DC Motor',
            'input_voltage': V,
            'input_current': I_armature,
            'input_power': P_input,
            'output_power': P_output,
            'mechanical_power': P_mechanical,
            'efficiency': efficiency,
            'speed_rpm': N_actual,
            'speed_rad_s': omega_actual,
            'torque': T_shaft,
            'load_torque': T_load,
            'developed_torque': T_developed,
            'armature_current': I_armature,
            'field_current': I_field,
            'back_emf': Ke * omega_actual,
            'armature_loss': P_armature_loss,
            'field_loss': P_field_loss,
            'core_loss': P_core_loss,
            'mechanical_loss': T_friction * omega_actual,
            'total_loss': P_total_loss,
            'operating_temp': operating_temp,
            'temp_rise': temp_rise,
            'current_density': current_density,
            'power_factor': 1.0,  # DC motor has unity power factor
            'speed_range': speed_range,
            'torque_curve': torque_curve,
            'current_curve': current_curve,
            'power_curve': power_curve,
            'efficiency_curve': efficiency_curve,
        }
        
        return results

    def simulate_ac_motor(self):
        """Simulate 3-phase AC induction motor with accurate calculations"""
        # Get parameters
        V = self.get_parameter_value('rated_voltage', 220)
        I_rated = self.get_parameter_value('rated_current', 10)
        f = self.get_parameter_value('frequency', 50)
        P_poles = int(self.get_parameter_value('poles', 4))
        Rs = self.get_parameter_value('stator_resistance', 1.2)
        Rr = self.get_parameter_value('rotor_resistance', 0.8)
        Im = self.get_parameter_value('magnetizing_current', 3.0)
        slip_percent = self.get_parameter_value('slip', 3.0)
        T_load = self.get_parameter_value('load_torque', 10)
        
        # Synchronous speed
        Ns = (120 * f) / P_poles
        omega_sync = Ns * 2 * math.pi / 60
        
        # Slip
        s = slip_percent / 100
        
        # Rotor speed
        Nr = Ns * (1 - s)
        omega_r = Nr * 2 * math.pi / 60
        
        # Simplified equivalent circuit calculations
        # Assuming Xm (magnetizing reactance) and X1, X2 (leakage reactances)
        Xm = V / (Im * math.sqrt(3))  # Approximate
        X1 = 0.1 * Xm  # Stator leakage reactance
        X2 = 0.1 * Xm  # Rotor leakage reactance (referred to stator)
        
        # Rotor current calculation (simplified)
        Rr_prime = Rr / s if s > 0 else float('inf')
        
        # Total impedance (simplified per-phase)
        Z = math.sqrt((Rs + Rr_prime)**2 + (X1 + X2)**2)
        
        # Phase voltage for Y-connected motor
        Vph = V / math.sqrt(3)
        
        # Stator current (simplified)
        I1 = Vph / Z if Z > 0 else 0
        
        # Power factor
        power_factor = (Rs + Rr_prime) / Z if Z > 0 else 0
        
        # Input power (3-phase)
        P_input = math.sqrt(3) * V * I1 * power_factor
        
        # Air gap power
        P_ag = P_input - 3 * I1**2 * Rs
        
        # Rotor copper loss
        P_rcl = s * P_ag
        
        # Mechanical power developed
        P_mech = P_ag - P_rcl
        
        # Torque developed
        T_developed = P_mech / omega_r if omega_r > 0 else 0
        
        # Core losses (approximation)
        P_core = 0.03 * P_input
        
        # Friction and windage losses
        P_fw = 0.01 * P_input
        
        # Output power
        P_output = P_mech - P_fw
        
        # Output torque
        T_output = P_output / omega_r if omega_r > 0 else 0
        
        # Total losses
        P_loss = P_input - P_output
        
        # Efficiency
        efficiency = (P_output / P_input * 100) if P_input > 0 else 0
        
        # Temperature calculations
        thermal_resistance = 6  # °C/W
        temp_rise = P_loss * thermal_resistance
        ambient_temp = self.get_parameter_value('ambient_temp', 40)
        operating_temp = ambient_temp + temp_rise
        
        # Performance curves - Torque-Slip characteristics
        slip_range = np.linspace(0.001, 1.0, 50)
        speed_range = Ns * (1 - slip_range)
        torque_curve = []
        current_curve = []
        power_curve = []
        efficiency_curve = []
        
        for s_val in slip_range:
            # Torque calculation using simplified formula
            Rr_s = Rr / s_val if s_val > 0 else float('inf')
            Z_s = math.sqrt((Rs + Rr_s)**2 + (X1 + X2)**2)
            I_s = Vph / Z_s if Z_s > 0 else 0
            
            # Torque (approximate)
            T_s = (3 * Vph**2 * Rr / s_val) / (omega_sync * ((Rs + Rr/s_val)**2 + (X1 + X2)**2)) if s_val > 0 else 0
            
            P_in_s = math.sqrt(3) * V * I_s * ((Rs + Rr_s) / Z_s)
            P_out_s = T_s * Ns * (1 - s_val) * 2 * math.pi / 60
            eff_s = (P_out_s / P_in_s * 100) if P_in_s > 0 else 0
            
            torque_curve.append(T_s)
            current_curve.append(I_s * math.sqrt(3))  # Line current
            power_curve.append(P_out_s)
            efficiency_curve.append(min(eff_s, 100))
        
        results = {
            'motor_type': 'AC Induction Motor (3-Phase)',
            'input_voltage': V,
            'input_current': I1 * math.sqrt(3),  # Line current
            'input_power': P_input,
            'output_power': P_output,
            'mechanical_power': P_mech,
            'efficiency': efficiency,
            'speed_rpm': Nr,
            'speed_rad_s': omega_r,
            'synchronous_speed': Ns,
            'slip': s * 100,
            'torque': T_output,
            'load_torque': T_load,
            'developed_torque': T_developed,
            'frequency': f,
            'poles': P_poles,
            'power_factor': power_factor,
            'stator_current': I1,
            'stator_loss': 3 * I1**2 * Rs,
            'rotor_loss': P_rcl,
            'core_loss': P_core,
            'mechanical_loss': P_fw,
            'total_loss': P_loss,
            'operating_temp': operating_temp,
            'temp_rise': temp_rise,
            'air_gap_power': P_ag,
            'speed_range': speed_range,
            'torque_curve': torque_curve,
            'current_curve': current_curve,
            'power_curve': power_curve,
            'efficiency_curve': efficiency_curve,
        }
        
        return results

    def display_results(self, results):
        """Display simulation results in formatted text"""
        self.results_text.delete(1.0, tk.END)
        
        # Header
        self.results_text.insert(tk.END, "╔" + "═" * 48 + "╗\n", 'header')
        self.results_text.insert(tk.END, "║" + "  MOTOR SIMULATION RESULTS".center(48) + "║\n", 'header')
        self.results_text.insert(tk.END, "╚" + "═" * 48 + "╝\n\n", 'header')
        
        # Motor Type
        self.results_text.insert(tk.END, f"Motor Type: {results['motor_type']}\n\n", 'subheader')
        
        # Electrical Parameters
        self.results_text.insert(tk.END, "━━━ ELECTRICAL PARAMETERS ━━━\n", 'subheader')
        self.results_text.insert(tk.END, f"Input Voltage:        {results['input_voltage']:.2f} V\n", 'value')
        self.results_text.insert(tk.END, f"Input Current:        {results['input_current']:.2f} A\n", 'value')
        self.results_text.insert(tk.END, f"Input Power:          {results['input_power']:.2f} W\n", 'value')
        self.results_text.insert(tk.END, f"Power Factor:         {results['power_factor']:.3f}\n", 'value')
        
        if results['motor_type'] == 'DC Motor':
            self.results_text.insert(tk.END, f"Armature Current:     {results['armature_current']:.2f} A\n", 'value')
            self.results_text.insert(tk.END, f"Field Current:        {results['field_current']:.2f} A\n", 'value')
            self.results_text.insert(tk.END, f"Back EMF:             {results['back_emf']:.2f} V\n", 'value')
        else:
            self.results_text.insert(tk.END, f"Frequency:            {results['frequency']:.1f} Hz\n", 'value')
            self.results_text.insert(tk.END, f"Synchronous Speed:    {results['synchronous_speed']:.1f} RPM\n", 'value')
            self.results_text.insert(tk.END, f"Slip:                 {results['slip']:.2f} %\n", 'value')
            self.results_text.insert(tk.END, f"Air Gap Power:        {results['air_gap_power']:.2f} W\n", 'value')
        
        # Mechanical Parameters
        self.results_text.insert(tk.END, "\n━━━ MECHANICAL PARAMETERS ━━━\n", 'subheader')
        self.results_text.insert(tk.END, f"Operating Speed:      {results['speed_rpm']:.2f} RPM\n", 'value')
        self.results_text.insert(tk.END, f"Angular Velocity:     {results['speed_rad_s']:.2f} rad/s\n", 'value')
        self.results_text.insert(tk.END, f"Output Torque:        {results['torque']:.2f} Nm\n", 'value')
        self.results_text.insert(tk.END, f"Load Torque:          {results['load_torque']:.2f} Nm\n", 'value')
        self.results_text.insert(tk.END, f"Developed Torque:     {results['developed_torque']:.2f} Nm\n", 'value')
        
        # Power Analysis
        self.results_text.insert(tk.END, "\n━━━ POWER ANALYSIS ━━━\n", 'subheader')
        self.results_text.insert(tk.END, f"Mechanical Power:     {results['mechanical_power']:.2f} W\n", 'value')
        self.results_text.insert(tk.END, f"Output Power:         {results['output_power']:.2f} W\n", 'value')
        self.results_text.insert(tk.END, f"Efficiency:           {results['efficiency']:.2f} %\n", 'value')
        
        # Check efficiency status
        if results['efficiency'] > 85:
            self.results_text.insert(tk.END, "Status:               ✓ Excellent\n", 'value')
        elif results['efficiency'] > 70:
            self.results_text.insert(tk.END, "Status:               ⚠ Good\n", 'warning')
        else:
            self.results_text.insert(tk.END, "Status:               ✗ Poor - Check Parameters\n", 'error')
        
        # Losses Breakdown
        self.results_text.insert(tk.END, "\n━━━ LOSSES BREAKDOWN ━━━\n", 'subheader')
        
        if results['motor_type'] == 'DC Motor':
            self.results_text.insert(tk.END, f"Armature Loss:        {results['armature_loss']:.2f} W\n", 'value')
            self.results_text.insert(tk.END, f"Field Loss:           {results['field_loss']:.2f} W\n", 'value')
        else:
            self.results_text.insert(tk.END, f"Stator Loss:          {results['stator_loss']:.2f} W\n", 'value')
            self.results_text.insert(tk.END, f"Rotor Loss:           {results['rotor_loss']:.2f} W\n", 'value')
        
        self.results_text.insert(tk.END, f"Core Loss:            {results['core_loss']:.2f} W\n", 'value')
        self.results_text.insert(tk.END, f"Mechanical Loss:      {results['mechanical_loss']:.2f} W\n", 'value')
        self.results_text.insert(tk.END, f"Total Losses:         {results['total_loss']:.2f} W\n", 'value')
        loss_percentage = (results['total_loss'] / results['input_power'] * 100) if results['input_power'] > 0 else 0
        self.results_text.insert(tk.END, f"Loss Percentage:      {loss_percentage:.2f} %\n", 'value')
        
        # Thermal Analysis
        self.results_text.insert(tk.END, "\n━━━ THERMAL ANALYSIS ━━━\n", 'subheader')
        self.results_text.insert(tk.END, f"Temperature Rise:     {results['temp_rise']:.2f} °C\n", 'value')
        self.results_text.insert(tk.END, f"Operating Temp:       {results['operating_temp']:.2f} °C\n", 'value')
        
        # Temperature warning
        insulation_class = self.param_vars['insulation_class'].get().upper()
        temp_limits = {'A': 105, 'B': 130, 'F': 155, 'H': 180}
        limit = temp_limits.get(insulation_class, 155)
        
        if results['operating_temp'] < limit * 0.8:
            self.results_text.insert(tk.END, f"Status:               ✓ Safe (Limit: {limit}°C)\n", 'value')
        elif results['operating_temp'] < limit:
            self.results_text.insert(tk.END, f"Status:               ⚠ Caution (Limit: {limit}°C)\n", 'warning')
        else:
            self.results_text.insert(tk.END, f"Status:               ✗ Overheating! (Limit: {limit}°C)\n", 'error')
        
        # Performance Rating
        self.results_text.insert(tk.END, "\n━━━ PERFORMANCE RATING ━━━\n", 'subheader')
        rating = self.calculate_performance_rating(results)
        self.results_text.insert(tk.END, f"Overall Rating:       {rating['score']:.1f}/100\n", 'value')
        self.results_text.insert(tk.END, f"Grade:                {rating['grade']}\n", 'value')
        self.results_text.insert(tk.END, f"Assessment:           {rating['assessment']}\n", 'value')
        
        self.results_text.insert(tk.END, "\n" + "═" * 50 + "\n", 'header')
        self.results_text.insert(tk.END, f"Simulation completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n", 'value')

    def calculate_performance_rating(self, results):
        """Calculate overall performance rating"""
        score = 0
        
        # Efficiency (40 points)
        eff = results['efficiency']
        if eff >= 90:
            score += 40
        elif eff >= 80:
            score += 35
        elif eff >= 70:
            score += 25
        elif eff >= 60:
            score += 15
        else:
            score += 5
        
        # Power Factor (30 points)
        pf = results['power_factor']
        if pf >= 0.95:
            score += 30
        elif pf >= 0.85:
            score += 25
        elif pf >= 0.75:
            score += 15
        else:
            score += 5
        
        # Temperature (30 points)
        temp = results['operating_temp']
        insulation_class = self.param_vars['insulation_class'].get().upper()
        temp_limits = {'A': 105, 'B': 130, 'F': 155, 'H': 180}
        limit = temp_limits.get(insulation_class, 155)
        
        temp_ratio = temp / limit
        if temp_ratio < 0.7:
            score += 30
        elif temp_ratio < 0.85:
            score += 20
        elif temp_ratio < 1.0:
            score += 10
        else:
            score += 0
        
        # Determine grade
        if score >= 90:
            grade = 'A+ (Excellent)'
            assessment = 'Outstanding performance'
        elif score >= 80:
            grade = 'A (Very Good)'
            assessment = 'Very good performance'
        elif score >= 70:
            grade = 'B (Good)'
            assessment = 'Good performance'
        elif score >= 60:
            grade = 'C (Acceptable)'
            assessment = 'Acceptable performance'
        else:
            grade = 'D (Poor)'
            assessment = 'Needs improvement'
        
        return {'score': score, 'grade': grade, 'assessment': assessment}

    def update_visualizations(self, results):
        """Update all visualization graphs with stunning aesthetics"""
        # Clear all axes
        for ax in self.axes:
            ax.clear()
            ax.set_facecolor(self.colors['bg_primary'])
            ax.tick_params(colors=self.colors['text_secondary'], labelsize=8)
            for spine in ax.spines.values():
                spine.set_color(self.colors['accent_cyan'])
                spine.set_linewidth(1.5)
            ax.xaxis.label.set_color(self.colors['text_primary'])
            ax.yaxis.label.set_color(self.colors['text_primary'])
            ax.title.set_color(self.colors['accent_purple'])
            ax.grid(True, alpha=0.2, color=self.colors['text_secondary'], linestyle='--')
        
        speed_range = results['speed_range']
        
        # 1. Torque vs Speed - Neon cyan
        self.ax1.plot(speed_range, results['torque_curve'], 
                     color=self.colors['accent_cyan'], linewidth=2.5, label='Torque')
        self.ax1.axhline(y=results['load_torque'], color=self.colors['error'], 
                        linestyle='--', linewidth=2, label='Load Torque')
        self.ax1.axvline(x=results['speed_rpm'], color=self.colors['warning'], 
                        linestyle='--', linewidth=1.5, alpha=0.7, label='Operating Point')
        self.ax1.set_xlabel('Speed (RPM)', fontsize=9, fontweight='bold')
        self.ax1.set_ylabel('Torque (Nm)', fontsize=9, fontweight='bold')
        self.ax1.set_title('⚡ Torque-Speed Characteristic', fontsize=10, fontweight='bold')
        self.ax1.legend(fontsize=7, loc='best', facecolor=self.colors['bg_tertiary'], 
                       edgecolor=self.colors['accent_cyan'])
        
        # 2. Efficiency vs Speed - Neon green
        self.ax2.plot(speed_range, results['efficiency_curve'], 
                     color=self.colors['accent_green'], linewidth=2.5)
        self.ax2.axhline(y=results['efficiency'], color=self.colors['warning'], 
                        linestyle='--', linewidth=2, 
                        label=f'Operating: {results["efficiency"]:.1f}%')
        self.ax2.axvline(x=results['speed_rpm'], color=self.colors['warning'], 
                        linestyle='--', linewidth=1.5, alpha=0.7)
        self.ax2.set_xlabel('Speed (RPM)', fontsize=9, fontweight='bold')
        self.ax2.set_ylabel('Efficiency (%)', fontsize=9, fontweight='bold')
        self.ax2.set_title('📈 Efficiency Curve', fontsize=10, fontweight='bold')
        self.ax2.legend(fontsize=7, loc='best', facecolor=self.colors['bg_tertiary'],
                       edgecolor=self.colors['accent_green'])
        self.ax2.set_ylim([0, 105])
        
        # 3. Current vs Speed - Purple
        self.ax3.plot(speed_range, results['current_curve'], 
                     color=self.colors['accent_purple'], linewidth=2.5)
        self.ax3.axhline(y=results['input_current'], color=self.colors['warning'], 
                        linestyle='--', linewidth=2, 
                        label=f'Operating: {results["input_current"]:.1f}A')
        self.ax3.axvline(x=results['speed_rpm'], color=self.colors['warning'], 
                        linestyle='--', linewidth=1.5, alpha=0.7)
        self.ax3.set_xlabel('Speed (RPM)', fontsize=9, fontweight='bold')
        self.ax3.set_ylabel('Current (A)', fontsize=9, fontweight='bold')
        self.ax3.set_title('⚡ Current vs Speed', fontsize=10, fontweight='bold')
        self.ax3.legend(fontsize=7, loc='best', facecolor=self.colors['bg_tertiary'],
                       edgecolor=self.colors['accent_purple'])
        
        # 4. Power vs Speed - Orange
        self.ax4.plot(speed_range, np.array(results['power_curve'])/1000, 
                     color=self.colors['accent_orange'], linewidth=2.5, label='Output Power')
        self.ax4.axhline(y=results['output_power']/1000, color=self.colors['accent_cyan'], 
                        linestyle='--', linewidth=2, 
                        label=f'Operating: {results["output_power"]/1000:.2f}kW')
        self.ax4.axvline(x=results['speed_rpm'], color=self.colors['warning'], 
                        linestyle='--', linewidth=1.5, alpha=0.7)
        self.ax4.set_xlabel('Speed (RPM)', fontsize=9, fontweight='bold')
        self.ax4.set_ylabel('Power (kW)', fontsize=9, fontweight='bold')
        self.ax4.set_title('⚡ Power Output Curve', fontsize=10, fontweight='bold')
        self.ax4.legend(fontsize=7, loc='best', facecolor=self.colors['bg_tertiary'],
                       edgecolor=self.colors['accent_orange'])
        
        # 5. Losses Distribution (Pie Chart) - Neon colors
        if results['motor_type'] == 'DC Motor':
            losses = [
                results['armature_loss'],
                results['field_loss'],
                results['core_loss'],
                results['mechanical_loss']
            ]
            labels = ['Armature\nLoss', 'Field\nLoss', 'Core\nLoss', 'Mechanical\nLoss']
        else:
            losses = [
                results['stator_loss'],
                results['rotor_loss'],
                results['core_loss'],
                results['mechanical_loss']
            ]
            labels = ['Stator\nLoss', 'Rotor\nLoss', 'Core\nLoss', 'Mechanical\nLoss']
        
        colors_pie = [self.colors['error'], self.colors['accent_orange'], 
                     self.colors['warning'], self.colors['accent_purple']]
        explode = (0.05, 0.05, 0.05, 0.05)
        
        wedges, texts, autotexts = self.ax5.pie(losses, labels=labels, autopct='%1.1f%%',
                                                 colors=colors_pie, explode=explode,
                                                 startangle=90, 
                                                 textprops={'color': self.colors['text_primary'], 
                                                           'fontsize': 8, 'fontweight': 'bold'})
        for autotext in autotexts:
            autotext.set_color('black')
            autotext.set_fontweight('bold')
        self.ax5.set_title('💥 Losses Distribution', fontsize=10, fontweight='bold',
                          color=self.colors['accent_purple'])
        
        # 6. Power Flow Diagram (Bar Chart) - Gradient effect
        power_stages = ['Input\nPower', 'Mechanical\nPower', 'Output\nPower']
        power_values = [results['input_power']/1000, 
                       results['mechanical_power']/1000, 
                       results['output_power']/1000]
        colors_bar = [self.colors['accent_cyan'], 
                     self.colors['accent_purple'], 
                     self.colors['accent_green']]
        
        bars = self.ax6.bar(power_stages, power_values, color=colors_bar, 
                           edgecolor=self.colors['text_primary'], linewidth=2)
        self.ax6.set_ylabel('Power (kW)', fontsize=9, fontweight='bold')
        self.ax6.set_title('⚡ Power Flow Analysis', fontsize=10, fontweight='bold')
        
        # Add value labels on bars with glow effect
        for bar, value in zip(bars, power_values):
            height = bar.get_height()
            self.ax6.text(bar.get_x() + bar.get_width()/2., height,
                         f'{value:.2f}kW',
                         ha='center', va='bottom', 
                         color=self.colors['text_primary'], 
                         fontsize=9, fontweight='bold',
                         bbox=dict(boxstyle='round,pad=0.3', 
                                  facecolor=self.colors['bg_tertiary'], 
                                  edgecolor=self.colors['accent_cyan'],
                                  alpha=0.8))
        
        self.fig.tight_layout(pad=2.5)
        self.canvas.draw()
        
        # Update status bar
        if hasattr(self, 'status_label'):
            self.status_label.config(
                text=f"✓ Simulation Complete | Efficiency: {results['efficiency']:.1f}% | Speed: {results['speed_rpm']:.0f} RPM"
            )

    def save_simulation(self):
        """Save simulation data to JSON file"""
        if not self.simulation_history:
            messagebox.showwarning("No Data", "No simulation data to save. Run a simulation first.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialfile=f"motor_sim_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        )
        
        if filename:
            try:
                # Prepare data for JSON (convert numpy arrays to lists)
                save_data = []
                for sim in self.simulation_history:
                    sim_copy = sim.copy()
                    for key in ['speed_range', 'torque_curve', 'current_curve', 'power_curve', 'efficiency_curve']:
                        if key in sim_copy and isinstance(sim_copy[key], np.ndarray):
                            sim_copy[key] = sim_copy[key].tolist()
                    save_data.append(sim_copy)
                
                with open(filename, 'w') as f:
                    json.dump(save_data, f, indent=2)
                
                messagebox.showinfo("Success", f"Simulation data saved to:\n{filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file:\n{str(e)}")

    def load_simulation(self):
        """Load simulation data from JSON file"""
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'r') as f:
                    data = json.load(f)
                
                if data and len(data) > 0:
                    # Load the most recent simulation
                    last_sim = data[-1]
                    
                    # Convert lists back to numpy arrays
                    for key in ['speed_range', 'torque_curve', 'current_curve', 'power_curve', 'efficiency_curve']:
                        if key in last_sim:
                            last_sim[key] = np.array(last_sim[key])
                    
                    # Display results
                    self.display_results(last_sim)
                    self.update_visualizations(last_sim)
                    
                    messagebox.showinfo("Success", f"Loaded simulation from:\n{filename}\n\nTotal simulations in file: {len(data)}")
                else:
                    messagebox.showwarning("Empty File", "The selected file contains no simulation data.")
                    
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file:\n{str(e)}")

    def export_excel(self):
        """Export simulation data to Excel"""
        if not PANDAS_AVAILABLE:
            messagebox.showerror("Error", "Pandas library not installed. Cannot export to Excel.")
            return
        
        if not self.simulation_history:
            messagebox.showwarning("No Data", "No simulation data to export. Run a simulation first.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"motor_sim_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if filename:
            try:
                with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                    # Summary sheet
                    summary_data = []
                    for i, sim in enumerate(self.simulation_history):
                        summary_data.append({
                            'Simulation #': i + 1,
                            'Timestamp': sim.get('timestamp', 'N/A'),
                            'Motor Type': sim.get('motor_type', 'N/A'),
                            'Input Power (W)': sim.get('input_power', 0),
                            'Output Power (W)': sim.get('output_power', 0),
                            'Efficiency (%)': sim.get('efficiency', 0),
                            'Speed (RPM)': sim.get('speed_rpm', 0),
                            'Torque (Nm)': sim.get('torque', 0),
                            'Current (A)': sim.get('input_current', 0)
                        })
                    
                    df_summary = pd.DataFrame(summary_data)
                    df_summary.to_excel(writer, sheet_name='Summary', index=False)
                    
                    # Detailed data for latest simulation
                    latest = self.simulation_history[-1]
                    
                    # Performance curves
                    if 'speed_range' in latest:
                        df_curves = pd.DataFrame({
                            'Speed (RPM)': latest['speed_range'],
                            'Torque (Nm)': latest['torque_curve'],
                            'Current (A)': latest['current_curve'],
                            'Power (W)': latest['power_curve'],
                            'Efficiency (%)': latest['efficiency_curve']
                        })
                        df_curves.to_excel(writer, sheet_name='Performance Curves', index=False)
                    
                    # Detailed parameters
                    param_data = []
                    for key, value in latest.items():
                        if not isinstance(value, (list, np.ndarray)):
                            param_data.append({'Parameter': key, 'Value': value})
                    
                    df_params = pd.DataFrame(param_data)
                    df_params.to_excel(writer, sheet_name='Detailed Parameters', index=False)
                
                messagebox.showinfo("Success", f"Data exported to Excel:\n{filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export to Excel:\n{str(e)}")

    def export_pdf(self):
        """Export simulation results to PDF"""
        if not FPDF_AVAILABLE:
            messagebox.showwarning("PDF Export", 
                                 "FPDF library not installed.\n\nInstall with: pip install fpdf\n\nFor now, use 'Save Data' to export as JSON.")
            return
        
        if not self.simulation_history:
            messagebox.showwarning("No Data", "No simulation data to export. Run a simulation first.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialfile=f"motor_sim_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        )
        
        if filename:
            try:
                results = self.simulation_history[-1]
                
                pdf = FPDF()
                pdf.add_page()
                
                # Title
                pdf.set_font('Arial', 'B', 16)
                pdf.cell(0, 10, 'Motor Simulation Report', 0, 1, 'C')
                pdf.ln(5)
                
                # Timestamp
                pdf.set_font('Arial', 'I', 10)
                pdf.cell(0, 10, f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', 0, 1, 'C')
                pdf.ln(5)
                
                # Motor Type
                pdf.set_font('Arial', 'B', 12)
                pdf.cell(0, 10, f'Motor Type: {results["motor_type"]}', 0, 1)
                pdf.ln(3)
                
                # Electrical Parameters
                pdf.set_font('Arial', 'B', 11)
                pdf.cell(0, 8, 'Electrical Parameters', 0, 1)
                pdf.set_font('Arial', '', 10)
                pdf.cell(0, 6, f'Input Voltage: {results["input_voltage"]:.2f} V', 0, 1)
                pdf.cell(0, 6, f'Input Current: {results["input_current"]:.2f} A', 0, 1)
                pdf.cell(0, 6, f'Input Power: {results["input_power"]:.2f} W', 0, 1)
                pdf.cell(0, 6, f'Power Factor: {results["power_factor"]:.3f}', 0, 1)
                pdf.ln(3)
                
                # Mechanical Parameters
                pdf.set_font('Arial', 'B', 11)
                pdf.cell(0, 8, 'Mechanical Parameters', 0, 1)
                pdf.set_font('Arial', '', 10)
                pdf.cell(0, 6, f'Speed: {results["speed_rpm"]:.2f} RPM', 0, 1)
                pdf.cell(0, 6, f'Torque: {results["torque"]:.2f} Nm', 0, 1)
                pdf.cell(0, 6, f'Output Power: {results["output_power"]:.2f} W', 0, 1)
                pdf.ln(3)
                
                # Performance
                pdf.set_font('Arial', 'B', 11)
                pdf.cell(0, 8, 'Performance', 0, 1)
                pdf.set_font('Arial', '', 10)
                pdf.cell(0, 6, f'Efficiency: {results["efficiency"]:.2f} %', 0, 1)
                pdf.cell(0, 6, f'Total Losses: {results["total_loss"]:.2f} W', 0, 1)
                pdf.cell(0, 6, f'Operating Temperature: {results["operating_temp"]:.2f} C', 0, 1)
                
                pdf.output(filename)
                messagebox.showinfo("Success", f"Report exported to PDF:\n{filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export to PDF:\n{str(e)}")

    def reset_parameters(self):
        """Reset all parameters to default values"""
        response = messagebox.askyesno("Reset Parameters", 
                                      "Are you sure you want to reset all parameters to default values?")
        if response:
            # Recreate parameter inputs
            for widget in self.params_frame.winfo_children():
                widget.destroy()
            
            self.create_parameter_inputs()
            
            # Clear results
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(tk.END, "Parameters reset to default values.\nRun simulation to see results.", 'value')
            
            # Clear graphs
            for ax in self.axes:
                ax.clear()
                ax.set_facecolor(self.colors['bg_primary'])
                ax.tick_params(colors=self.colors['text_secondary'], labelsize=8)
                for spine in ax.spines.values():
                    spine.set_color(self.colors['accent_cyan'])
                    spine.set_linewidth(1.5)
                ax.xaxis.label.set_color(self.colors['text_primary'])
                ax.yaxis.label.set_color(self.colors['text_primary'])
                ax.title.set_color(self.colors['accent_purple'])
                ax.grid(True, alpha=0.2, color=self.colors['text_secondary'], linestyle='--')
            
            self.canvas.draw()


def main():
    """Main application entry point"""
    root = tk.Tk()
    app = MotorSimulator(root)
    root.mainloop()


if __name__ == "__main__":
    main()
