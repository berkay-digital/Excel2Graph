import os
import pandas as pd
import matplotlib.pyplot as plt
from glob import glob
import numpy as np
from scipy import interpolate
import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
from matplotlib.lines import Line2D
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import webbrowser

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

FRAME_COLOR = "#2b2b2b"
INNER_FRAME_COLOR = "#232323"

root = ctk.CTk()
root.title("Excel 2 Graph")
root.geometry("1200x800")
root.minsize(900, 600)

input_folder = ctk.StringVar(value="./excel")
output_folder = ctk.StringVar(value="./graph")

num_series = ctk.IntVar(value=1)
x_label = ctk.StringVar(value="ε")
y_label = ctk.StringVar(value="σ")
legend_position = ctk.StringVar(value="upper left")

series_colors = []
series_markers = []
default_colors = ['red', 'blue', 'green', 'purple', 'orange', 'brown', 'pink', 'gray', 'olive', 'cyan', 
                 'darkred', 'navy', 'lime', 'magenta', 'gold', 'teal', 'violet', 'coral', 'darkgreen', 'skyblue']
default_markers = ['o', 's', '^', 'v', 'D', 'p', '*', 'h', '+', 'x', '>', '<', '1', '2', '3', '4', '8', 'P', 'X', 'd']

os.makedirs("./excel", exist_ok=True)
os.makedirs("./graph", exist_ok=True)

series_names = []

show_legend = ctk.BooleanVar(value=True)
show_symbols = ctk.BooleanVar(value=True)

graph_font = ctk.StringVar(value="Times New Roman")
available_fonts = ["Times New Roman", "Arial", "Helvetica", "Calibri", "Cambria", "Georgia"]

sheet_name = ctk.StringVar(value="Sheet1")

def select_input_folder():
    folder = filedialog.askdirectory()
    if folder:
        input_folder.set(folder)

def select_output_folder():
    folder = filedialog.askdirectory()
    if folder:
        output_folder.set(folder)

def update_preview():
    try:
        for widget in preview_frame.winfo_children():
            widget.destroy()
        
        frame_width = preview_frame.winfo_width()
        frame_height = preview_frame.winfo_height()
        
        preview_width = max(300, frame_width - 40)
        preview_height = max(200, frame_height - 40)
        
        dpi = 100
        
        fig_width = preview_width / dpi
        fig_height = preview_height / dpi
        
        fig = plt.Figure(figsize=(fig_width, fig_height), dpi=dpi)
        ax = fig.add_subplot(111)
        
        x = np.linspace(0, 10, 100)
        legend_elements = []
        
        for i in range(num_series.get()):
            if i == 0:
                y = np.sin(x)
            elif i == 1:
                y = np.cos(x)
            elif i == 2:
                y = x * 0.1
            else:
                y = np.sin(x + i * np.pi/4)
            
            color = series_colors[i] if i < len(series_colors) else default_colors[i % len(default_colors)]
            marker = series_markers[i] if i < len(series_markers) else default_markers[i % len(default_markers)]
            
            series_label = series_names[i] if i < len(series_names) else f'Series {i+1}'
            
            line = ax.plot(x, y, color=color, label=series_label)
            
            if show_symbols.get():
                ax.scatter(x[::10], y[::10], color=color, marker=marker)
            
            legend_elements.append(Line2D([0], [0], color=color, 
                                        marker=marker if show_symbols.get() else None,
                                        label=series_label, 
                                        markerfacecolor=color, 
                                        markersize=8))
        
        ax.set_xlabel(x_label.get(), fontsize=10, fontweight='bold', color='darkred', fontname=graph_font.get())
        ax.set_ylabel(y_label.get(), fontsize=10, fontweight='bold', color='darkred', fontname=graph_font.get())
        ax.set_title('Preview', fontsize=12, fontweight='semibold', color='navy', fontname=graph_font.get())
        
        for label in ax.get_xticklabels() + ax.get_yticklabels():
            label.set_fontname(graph_font.get())
        
        if legend_elements and show_legend.get():
            ax.legend(handles=legend_elements, loc=legend_position.get(), fontsize=8, prop={'family': graph_font.get()})
        
        ax.set_facecolor('#2b2b2b')
        fig.patch.set_facecolor('#2b2b2b')
        ax.tick_params(colors='white')
        ax.xaxis.label.set_color('white')
        ax.yaxis.label.set_color('white')
        ax.title.set_color('white')
        
        canvas = FigureCanvasTkAgg(fig, master=preview_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    except Exception as e:
        print(f"Preview update error: {e}")
        pass

def create_color_config_frame():
    global series_names
    color_config_frame = ctk.CTkFrame(settings_frame, fg_color=INNER_FRAME_COLOR, corner_radius=8, height=100)
    color_config_frame.pack(fill="x", padx=10, pady=1)
    color_config_frame.pack_propagate(False)
    
    scroll_frame = ctk.CTkScrollableFrame(color_config_frame, height=100)
    scroll_frame.pack(fill="both", expand=True, padx=5, pady=1)
    
    series_configs = []

    def change_series_name(index, label):
        global series_names
        new_name = simpledialog.askstring("Change Series Name", f"Enter new name for Series {index + 1}:")
        if new_name:
            while len(series_names) <= index:
                series_names.append(f"Series {len(series_names) + 1}")
            series_names[index] = new_name
            label.configure(text=f"{new_name}:")
            update_preview()
    
    def update_series_count(*args):
        for widget in scroll_frame.winfo_children():
            widget.destroy()
        
        for i in range(num_series.get()):
            series_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent", height=25)
            series_frame.pack(fill="x", pady=0)
            
            series_label = ctk.CTkLabel(series_frame, 
                                      text=f"{series_names[i] if i < len(series_names) else f'S{i+1}'}:",
                                      width=40,
                                      font=("Arial", 12))
            series_label.pack(side="left", padx=1)
            
            edit_button = ctk.CTkButton(series_frame, 
                                      text="✎",
                                      width=20,
                                      height=20,
                                      font=("Arial", 12))
            edit_button.configure(command=lambda i=i, label=series_label: change_series_name(i, label))
            edit_button.pack(side="left", padx=1)
            
            color_var = ctk.StringVar(value=series_colors[i] if i < len(series_colors) else default_colors[i % len(default_colors)])
            color_combo = ctk.CTkComboBox(series_frame, values=default_colors, 
                                        variable=color_var, width=80,
                                        height=25,
                                        font=("Arial", 12))
            color_combo.pack(side="left", padx=1)
            
            marker_var = ctk.StringVar(value=series_markers[i] if i < len(series_markers) else default_markers[i % len(default_markers)])
            marker_combo = ctk.CTkComboBox(series_frame, values=default_markers,
                                         variable=marker_var, width=80,
                                         height=25,
                                         font=("Arial", 12))
            marker_combo.pack(side="left", padx=1)
            
            color_preview = ctk.CTkLabel(series_frame, text="■", font=("Arial", 12),
                                       text_color=color_var.get())
            color_preview.pack(side="left", padx=1)
            
            def create_update_callback(index, color_v, marker_v, preview_label):
                def update(*args):
                    while len(series_colors) <= index:
                        series_colors.append(default_colors[len(series_colors) % len(default_colors)])
                    while len(series_markers) <= index:
                        series_markers.append(default_markers[len(series_markers) % len(default_markers)])
                    
                    series_colors[index] = color_v.get()
                    series_markers[index] = marker_v.get()
                    preview_label.configure(text_color=color_v.get())
                    update_preview()
                return update
            
            update_func = create_update_callback(i, color_var, marker_var, color_preview)
            color_var.trace_add("write", update_func)
            marker_var.trace_add("write", update_func)
            
            series_configs.append((color_var, marker_var))
    
    update_series_count()
    num_series.trace_add("write", update_series_count)
    
    return color_config_frame

def create_graphs():
    if not input_folder.get() or not output_folder.get():
        messagebox.showwarning("Missing Folders", "Please select both input and output folders.")
        return

    if len(series_colors) == 0 or len(series_markers) == 0:
        messagebox.showwarning("Missing Configuration", "Please configure series colors and markers first.")
        return

    os.makedirs(output_folder.get(), exist_ok=True)

    for filepath in glob(os.path.join(input_folder.get(), '*.xlsx')):
        filename = os.path.splitext(os.path.basename(filepath))[0]
        
        if filename.startswith('~'):
            continue
        
        try:
            data = pd.read_excel(filepath, sheet_name=sheet_name.get())
        except Exception as e:
            messagebox.showerror("Sheet Error", f"Error reading sheet '{sheet_name.get()}' from {filename}. Please verify the sheet name.")
            status_label.configure(text=f"Error processing {filename}: Invalid sheet name")
            continue
        
        x_columns = [col for col in data.columns if col.startswith('X')]
        y_columns = [col for col in data.columns if col.startswith('Y')]
        xy_pairs = [(x, y) for x in x_columns for y in y_columns if x[1:] == y[1:]]
        
        plt.figure(figsize=(12, 8), dpi=300)
        
        plt.rcParams['figure.facecolor'] = 'white'
        plt.rcParams['axes.facecolor'] = 'white'
        
        plt.subplots_adjust(left=0.12, right=0.95, top=0.95, bottom=0.12)
        
        plt.xlabel(x_label.get(), fontsize=12, fontweight='bold', fontname=graph_font.get())
        plt.ylabel(y_label.get(), fontsize=12, fontweight='bold', fontname=graph_font.get())
        plt.title(f'{filename}', fontsize=14, fontweight='semibold', fontname=graph_font.get())
        
        plt.tick_params(axis='both', which='major', labelsize=10, width=1)
        plt.grid(True, which='major', linestyle='--', linewidth=0.5, alpha=0.7, color='gray')
        
        ax = plt.gca()
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_linewidth(1)
        ax.spines['bottom'].set_linewidth(1)
        
        legend_elements = []
        
        for idx, (x_col, y_col) in enumerate(xy_pairs[:num_series.get()]):
            x = data[x_col].to_numpy()
            y = data[y_col].to_numpy()
            
            color = series_colors[idx % len(series_colors)]
            marker = series_markers[idx % len(series_markers)]
            
            series_label = series_names[idx] if idx < len(series_names) else f'Series {idx+1}'
            
            try:
                sort_idx = np.argsort(x)
                x = x[sort_idx]
                y = y[sort_idx]
                
                if len(np.unique(x)) > 3:
                    tck = interpolate.splrep(x, y, s=0)
                    x_smooth = np.linspace(min(x), max(x), 1000)
                    y_smooth = interpolate.splev(x_smooth, tck)
                    plt.plot(x_smooth, y_smooth, color=color, linewidth=1.5)
                else:
                    plt.plot(x, y, color=color, linewidth=1.5)
            except:
                plt.plot(x, y, color=color, linewidth=1.5)
            
            if show_symbols.get():
                marker_interval = max(1, len(x) // 50)
                plt.plot(x[::marker_interval], y[::marker_interval], 
                        marker=marker, color=color, 
                        markersize=6, linestyle='none')
            
            legend_elements.append(Line2D([0], [0], 
                                        marker=marker if show_symbols.get() else None, 
                                        color=color, 
                                        label=series_label,
                                        markerfacecolor=color, 
                                        markersize=6,
                                        linewidth=1.5))
        
        if show_legend.get():
            plt.legend(handles=legend_elements, 
                      title='Data Series', 
                      loc=legend_position.get(), 
                      fontsize=10, 
                      frameon=True, 
                      framealpha=0.8,
                      prop={'family': graph_font.get()})
        
        plt.tight_layout()
        
        output_filepath = os.path.join(output_folder.get(), f'{filename}.png')
        plt.savefig(output_filepath, format='png', dpi=300, bbox_inches='tight')
        plt.close()

        print(f'{filename} file has been converted to {output_filepath}!')
        status_label.configure(text=f'Processing: {filename}')
        root.update()
        
    status_label.configure(text="All graphs have been created and saved!")

def create_series_slider():
    slider_frame = ctk.CTkFrame(series_frame, fg_color="transparent")
    slider_frame.pack(fill="x", padx=10, pady=5)
    
    slider = ctk.CTkSlider(slider_frame,
                          from_=1,
                          to=10,
                          number_of_steps=9,
                          command=lambda value: num_series.set(int(value)))
    slider.set(num_series.get())
    slider.pack(side="left", fill="x", expand=True, padx=10)
    
    value_label = ctk.CTkLabel(slider_frame, text="1")
    value_label.pack(side="left", padx=10)
    
    def update_label(value):
        value_label.configure(text=str(int(value)))
        num_series.set(int(value))
        update_preview()
    
    slider.configure(command=update_label)

main_container = ctk.CTkFrame(root, fg_color="transparent")
main_container.pack(fill="both", expand=True, padx=20, pady=20)

left_frame = ctk.CTkFrame(main_container, fg_color="transparent")
left_frame.pack(side="left", fill="both", expand=True, padx=(0, 10), pady=0)
left_frame.pack_propagate(False)
left_frame.configure(width=700)

right_frame = ctk.CTkFrame(main_container, fg_color=FRAME_COLOR, corner_radius=10)
right_frame.pack(side="right", fill="both", expand=True, padx=(10, 0), pady=0)
right_frame.pack_propagate(False)
right_frame.configure(width=400)

folder_frame = ctk.CTkFrame(left_frame, fg_color=FRAME_COLOR, corner_radius=10)
folder_frame.pack(fill="x", padx=20, pady=10)

ctk.CTkLabel(folder_frame, 
             text="Folder Settings", 
             font=("Arial", 16, "bold"), 
             text_color="#3498db").pack(anchor="w", padx=15, pady=10)

input_frame = ctk.CTkFrame(folder_frame, fg_color=INNER_FRAME_COLOR, corner_radius=8)
input_frame.pack(fill="x", padx=10, pady=(5, 5))
ctk.CTkLabel(input_frame, text="Input Folder:").pack(side="left", padx=10)
ctk.CTkEntry(input_frame, textvariable=input_folder, width=400).pack(side="left", padx=10)
ctk.CTkButton(input_frame, text="Browse", command=select_input_folder, width=100).pack(side="left")

output_frame = ctk.CTkFrame(folder_frame, fg_color=INNER_FRAME_COLOR, corner_radius=8)
output_frame.pack(fill="x", padx=10, pady=(0, 10))
ctk.CTkLabel(output_frame, text="Output Folder:").pack(side="left", padx=10)
ctk.CTkEntry(output_frame, textvariable=output_folder, width=400).pack(side="left", padx=10)
ctk.CTkButton(output_frame, text="Browse", command=select_output_folder, width=100).pack(side="left")

sheet_frame = ctk.CTkFrame(folder_frame, fg_color=INNER_FRAME_COLOR, corner_radius=8)
sheet_frame.pack(fill="x", padx=10, pady=(0, 10))
ctk.CTkLabel(sheet_frame, text="Sheet Name:").pack(side="left", padx=10)
ctk.CTkEntry(sheet_frame, textvariable=sheet_name, width=200).pack(side="left", padx=10)

settings_frame = ctk.CTkFrame(left_frame, fg_color=FRAME_COLOR, corner_radius=10)
settings_frame.pack(fill="x", padx=20, pady=10)

ctk.CTkLabel(settings_frame, 
             text="Graph Settings", 
             font=("Arial", 16, "bold"), 
             text_color="#3498db").pack(anchor="w", padx=15, pady=10)

series_frame = ctk.CTkFrame(settings_frame, fg_color=INNER_FRAME_COLOR, corner_radius=8)
series_frame.pack(fill="x", padx=10, pady=5)
ctk.CTkLabel(series_frame, text="Number of Data Series:").pack(side="left", padx=10)
create_series_slider()

color_config_frame = create_color_config_frame()

labels_frame = ctk.CTkFrame(settings_frame, fg_color=INNER_FRAME_COLOR, corner_radius=8)
labels_frame.pack(fill="x", padx=10, pady=5)
ctk.CTkLabel(labels_frame, text="X-axis Label:").pack(side="left", padx=10)
ctk.CTkEntry(labels_frame, textvariable=x_label, width=100).pack(side="left", padx=10)
ctk.CTkLabel(labels_frame, text="Y-axis Label:").pack(side="left", padx=10)
ctk.CTkEntry(labels_frame, textvariable=y_label, width=100).pack(side="left", padx=10)

legend_frame = ctk.CTkFrame(settings_frame, fg_color=INNER_FRAME_COLOR, corner_radius=8)
legend_frame.pack(fill="x", padx=10, pady=(5, 10))
ctk.CTkLabel(legend_frame, text="Legend Position:").pack(side="left", padx=10)
legend_combobox = ctk.CTkComboBox(legend_frame, 
                                 values=['upper right', 'upper left', 'lower left', 'lower right',
                                        'right', 'center left', 'center right', 'lower center',
                                        'upper center', 'center'],
                                 variable=legend_position,
                                 width=200)
legend_combobox.pack(side="left", padx=10)

visibility_frame = ctk.CTkFrame(settings_frame, fg_color=INNER_FRAME_COLOR, corner_radius=8)
visibility_frame.pack(fill="x", padx=10, pady=(5, 10))

legend_switch = ctk.CTkSwitch(visibility_frame, 
                            text="Show Legend",
                            variable=show_legend,
                            command=update_preview)
legend_switch.pack(side="left", padx=10)

symbols_switch = ctk.CTkSwitch(visibility_frame, 
                             text="Show Symbols",
                             variable=show_symbols,
                             command=update_preview)
symbols_switch.pack(side="left", padx=10)

font_frame = ctk.CTkFrame(settings_frame, fg_color=INNER_FRAME_COLOR, corner_radius=8)
font_frame.pack(fill="x", padx=10, pady=(5, 10))

ctk.CTkLabel(font_frame, text="Graph Font:").pack(side="left", padx=10)
font_combobox = ctk.CTkComboBox(font_frame, 
                               values=available_fonts,
                               variable=graph_font,
                               width=200,
                               command=lambda _: update_preview())
font_combobox.pack(side="left", padx=10)

action_frame = ctk.CTkFrame(left_frame, fg_color=FRAME_COLOR, corner_radius=10)
action_frame.pack(fill="x", padx=20, pady=10)

create_button = ctk.CTkButton(action_frame,
                             text="Create Graphs",
                             command=create_graphs,
                             height=40,
                             fg_color="#2ecc71",
                             hover_color="#27ae60",
                             corner_radius=8)
create_button.pack(pady=10, padx=10)

status_frame = ctk.CTkFrame(left_frame, fg_color=FRAME_COLOR, corner_radius=10)
status_frame.pack(fill="x", padx=20, pady=10)
status_label = ctk.CTkLabel(status_frame, 
                           text="Ready to process...", 
                           font=("Arial", 12),
                           text_color="#3498db")
status_label.pack(pady=10)

preview_label = ctk.CTkLabel(right_frame, text="Graph Preview", font=("Arial", 16, "bold"), text_color="#3498db")
preview_label.pack(pady=10)

preview_frame = ctk.CTkFrame(right_frame, fg_color=INNER_FRAME_COLOR, corner_radius=8)
preview_frame.pack(fill="both", expand=True, padx=10, pady=10)

update_preview()

def on_setting_change(*args):
    update_preview()

x_label.trace_add("write", on_setting_change)
y_label.trace_add("write", on_setting_change)
legend_position.trace_add("write", on_setting_change)
num_series.trace_add("write", on_setting_change)

legend_combobox.configure(command=lambda _: update_preview())

def on_window_resize(event=None):
    if event and event.widget != root:
        return
        
    try:
        window_width = root.winfo_width()
        window_height = root.winfo_height()
        
        min_left_width = 700
        min_right_width = 400
        
        available_width = window_width - 60
        
        left_width = max(min_left_width, int(available_width * 0.6))
        right_width = max(min_right_width, int(available_width * 0.4))
        
        current_left_width = left_frame.winfo_width()
        current_right_width = right_frame.winfo_width()
        
        if abs(current_left_width - left_width) > 10:
            left_frame.configure(width=left_width)
            
        if abs(current_right_width - right_width) > 10:
            right_frame.configure(width=right_width)
        
        root.after_cancel(root.after_id) if hasattr(root, 'after_id') else None
        root.after_id = root.after(100, update_preview)
        
    except Exception as e:
        print(f"Resize error: {e}")
        pass

def delayed_resize(event):
    if hasattr(root, 'resize_after_id'):
        root.after_cancel(root.resize_after_id)
    root.resize_after_id = root.after(100, lambda: on_window_resize(event))

root.bind("<Configure>", delayed_resize)

root.after(100, on_window_resize)

def open_url(url):
    webbrowser.open(url)

footer_frame = ctk.CTkFrame(root, fg_color="transparent", height=30)
footer_frame.pack(side="bottom", fill="x", padx=10, pady=5)

center_container = ctk.CTkFrame(footer_frame, fg_color="transparent")
center_container.pack(expand=True)

url = "https://berkay.digital"
footer_label = ctk.CTkLabel(
    center_container, 
    text=f"Developed by Berkay", 
    font=("Arial", 10),
    text_color="#3498db",
    cursor="hand2"
)
footer_label.pack(padx=10)
footer_label.bind("<Button-1>", lambda e: open_url(url))

def on_enter(e):
    footer_label.configure(text_color="#2980b9")

def on_leave(e):
    footer_label.configure(text_color="#3498db")

footer_label.bind("<Enter>", on_enter)
footer_label.bind("<Leave>", on_leave)

root.mainloop()
