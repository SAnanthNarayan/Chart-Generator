import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, simpledialog
import sys

# Global variables
file_path = None
x_variable = None
y_variables = []
sheet_names = []
plot_type = None
range_count = 5  # Default range count for Pie Chart

# Function to select the Excel file
def select_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        ask_chart_type()  # Ask chart type first
    chart_path = os.path.join(os.path.dirname(file_path), "Charts")
    if not os.path.exists(chart_path):
        os.makedirs(chart_path)
# Function to ask for chart type (FTP, Duty Cycle, Bubble Chart, Pie Chart)
def ask_chart_type():
    global plot_type
    chart_type_window = Toplevel()
    chart_type_window.title("Select Chart Type")
    chart_type_window.geometry("300x200")

    def select_chart(chart_name):
        global plot_type
        plot_type = chart_name
        chart_type_window.destroy()
        get_sheets(file_path)

    tk.Button(chart_type_window, text="Steady Cycle (FTP, WHSC, WNTE)", command=lambda: select_chart("FTP")).pack(pady=5)
    tk.Button(chart_type_window, text="Transient Cycle (EMS Parameters)", command=lambda: select_chart("Duty Cycle")).pack(pady=5)
    tk.Button(chart_type_window, text="Bubble Chart (Duty Cycle)", command=lambda: select_chart("Bubble Chart")).pack(pady=5)
    tk.Button(chart_type_window, text="Pie Chart (Duty Cycle )", command=lambda: select_chart("Pie Chart")).pack(pady=5)

# Function to get available sheets in the selected Excel file
def get_sheets(file_path):
    global sheet_names
    sheets = pd.ExcelFile(file_path).sheet_names
    sheet_selection_window = Toplevel()
    sheet_selection_window.title("Select Sheets")
    sheet_selection_window.geometry("300x300")

    # Change selectmode based on plot_type (Single for Pie and Bubble Chart, Multiple for others)
    selectmode = tk.SINGLE if plot_type in ["Pie Chart", "Bubble Chart"] else tk.MULTIPLE
    listbox = tk.Listbox(sheet_selection_window, selectmode=selectmode)
    listbox.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

    for sheet in sheets:
        listbox.insert(tk.END, sheet)

    def select_sheets():
        global sheet_names
        selected_indices = listbox.curselection()

        if selected_indices:
            sheet_names = [listbox.get(i) for i in selected_indices]

            # Check if the selection is valid for Pie and Bubble Chart (only one sheet allowed)
            if plot_type in ["Pie Chart", "Bubble Chart"] and len(sheet_names) > 1:
                messagebox.showerror("Error", "Please select only one sheet for Pie Chart or Bubble Chart.")
                return

            sheet_selection_window.destroy()

            if plot_type == "Pie Chart":
                get_y_variables(file_path, sheet_names)
            else:
                get_x_variable(file_path, sheet_names)
        else:
            messagebox.showerror("Error", "Please select at least one sheet.")

    def select_all_sheets():
        # Select all sheets in the listbox
        listbox.select_set(0, tk.END)

    # Add the "Select All" button only for multiple selection
    if plot_type not in ["Pie Chart", "Bubble Chart"]:
        tk.Button(sheet_selection_window, text="Select All", command=select_all_sheets).pack(pady=5)

    tk.Button(sheet_selection_window, text="OK", command=select_sheets).pack(pady=10)

# Function to select X variable
def get_x_variable(file_path, sheet_names):
    global x_variable
    dfs = [pd.read_excel(file_path, sheet_name=sheet, header=[0, 1]) for sheet in sheet_names]
    for df in dfs:
        df.columns = pd.MultiIndex.from_tuples([(col[0].strip(), col[1].strip()) for col in df.columns])

    common_columns = list(set.intersection(*(set(df.columns) for df in dfs)))

    # Ensure common_columns is in the order as in the first sheet
    common_columns_ordered = [col for col in dfs[0].columns if col in common_columns]

    x_variable_window = Toplevel()
    x_variable_window.title("Select X-Axis Variable")
    x_variable_window.geometry("300x300")

    listbox = tk.Listbox(x_variable_window, selectmode=tk.SINGLE)
    listbox.pack(fill=tk.BOTH, expand=True)

    for var in common_columns_ordered:
        listbox.insert(tk.END, var[0])

    def select_x():
        global x_variable
        selected = listbox.curselection()
        if selected:
            x_variable = common_columns_ordered[selected[0]]
            x_variable_window.destroy()
            get_y_variables(file_path, sheet_names)
        else:
            messagebox.showerror("Error", "Please select an X-axis variable.")

    tk.Button(x_variable_window, text="OK", command=select_x).pack(pady=10)

# Function to select Y variables
def get_y_variables(file_path, sheet_names):
    global y_variables, range_count

    y_variable_window = Toplevel()
    y_variable_window.title("Select Y-Axis Variables")
    y_variable_window.geometry("300x400")

    dfs = [pd.read_excel(file_path, sheet_name=sheet, header=[0, 1]) for sheet in sheet_names]
    for df in dfs:
        df.columns = pd.MultiIndex.from_tuples([(col[0].strip(), col[1].strip()) for col in df.columns])

    common_columns = list(set.intersection(*(set(df.columns) for df in dfs)))
    if plot_type != "Pie Chart":
        common_columns = [var for var in common_columns if var != x_variable]

    # Ensure common_columns is in the order as in the first sheet
    common_columns_ordered = [col for col in dfs[0].columns if col in common_columns]

    listbox = tk.Listbox(y_variable_window, selectmode=tk.MULTIPLE)
    listbox.pack(fill=tk.BOTH, expand=True)

    for var in common_columns_ordered:
        listbox.insert(tk.END, var[0])

    def select_y():
        global y_variables, range_count
        selected_indices = listbox.curselection()
        if not selected_indices:
            messagebox.showerror("Error", "Please select at least one Y-axis variable.")
            return
        y_variables = [common_columns_ordered[i] for i in selected_indices]
        y_variable_window.destroy()

        # Ask for range count if Pie Chart and more than one Y variable is selected
        if plot_type == "Pie Chart" and len(y_variables) > 1:
            range_count = simpledialog.askinteger("Range Count", "Enter number of ranges:", initialvalue=5, minvalue=2)

        if plot_type == "Bubble Chart":
            bins = simpledialog.askinteger("Bins", "Enter number of bins for Bubble Chart:", initialvalue=10, minvalue=1)
            generate_bubble_chart(file_path, sheet_names[0], x_variable, y_variables[0], os.path.dirname(file_path), bins)
        else:
            run_plot()

    tk.Button(y_variable_window, text="Select All", command=lambda: listbox.select_set(0, tk.END)).pack(pady=5)

    tk.Button(y_variable_window, text="OK", command=select_y).pack(pady=10)

# Function to generate a professional-looking Pie Chart
def generate_pie_chart(file_path, sheet_names, y_variables, output_directory):
    os.makedirs(output_directory, exist_ok=True)
    df = pd.read_excel(file_path, sheet_name=sheet_names[0], header=[0, 1])
    df.columns = pd.MultiIndex.from_tuples([(col[0].strip(), col[1].strip()) for col in df.columns])

    for col in y_variables:
        values = df[col].dropna().values
        bins = np.linspace(values.min(), values.max(), range_count + 1)
        counts, _ = np.histogram(values, bins)

        labels = [f"{int(bins[i])} - {int(bins[i+1])}" for i in range(range_count)]
        colors = plt.cm.Dark2(np.linspace(0, 1, range_count))  # Darker color palette

        plt.figure(figsize=(10, 7))
        wedges, texts, autotexts = plt.pie(
            counts, labels=labels, autopct=lambda p: f"{p:.1f}%" if p > 0 else "",
            explode=[0.05] * range_count, startangle=140,
            colors=colors, textprops={'fontsize': 14, 'weight': 'bold'}
        )

        # Improve percentage text style
        for text in autotexts:
            text.set_fontsize(14)
            text.set_weight("bold")

        plt.title(f"{col[0]} ({col[1]}) Distribution", fontsize=18, fontweight='bold')

        # Move legend further away to avoid overlap
        plt.legend(wedges, labels, title="Ranges", loc="center left",
                   bbox_to_anchor=(1.2, 0.5), fontsize=12)

        # Save and close
        plt.savefig(os.path.join(output_directory, f"{col[0]}.png"), bbox_inches='tight', dpi=300)
        plt.close()
    sys.exit()

def generate_bubble_chart(file_path, sheet_name, x_variable, y_variable, output_directory, bins):
    try:
        os.makedirs(output_directory, exist_ok=True)  # Ensure output directory exists
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=[0, 1])
        df.columns = pd.MultiIndex.from_tuples([(col[0].strip(), col[1].strip()) for col in df.columns])

        for y_variable in y_variables:
            # Read the data from the selected sheet
            df[x_variable] = pd.to_numeric(df[x_variable], errors='coerce')  # Convert to numeric, set invalid values to NaN
            df[y_variable] = pd.to_numeric(df[y_variable], errors='coerce')  # Convert to numeric, set invalid values to NaN

            # Drop rows with NaN values (if any)
            df = df.dropna(subset=[x_variable, y_variable])

            # Extract speed and temperature data
            x_data = df[x_variable].values
            y_data = df[y_variable].values

            # Define bins for the X and Y variables (e.g., 10 bins)
            x_bins = np.linspace(min(x_data), max(x_data), bins)  # 10 bins for X
            y_bins = np.linspace(min(y_data), max(y_data), bins)  # 10 bins for Y

            # Create a 2D histogram (matrix) of the counts for X and Y
            histogram, x_edges, y_edges = np.histogram2d(x_data, y_data, bins=(x_bins, y_bins))

            # Convert counts to percentages
            total_points = histogram.sum()  # Total number of points
            percentages = (histogram / total_points) * 100  # Convert to percentages

            # Convert the matrix back to linear X, Y, and percentage data
            x_centers = (x_edges[:-1] + x_edges[1:]) / 2  # X bin centers
            y_centers = (y_edges[:-1] + y_edges[1:]) / 2  # Y bin centers

            # Create a DataFrame from the percentage matrix
            percent_df = pd.DataFrame(percentages, index=y_centers, columns=x_centers)

            # Convert the matrix to a linear format (for plotting)
            linear_df = percent_df.stack().reset_index()
            linear_df.columns = ['Y', 'X', 'Percentage']

            # Calculate bubble sizes (scaled by percentage)
            linear_df['Bubble Size'] = linear_df['Percentage'] * 250  # Scale the bubble size

            # Plot the scatter plot with bubble sizes
            plt.figure(figsize=(12, 8))
            scatter = plt.scatter(
                linear_df['X'],  # X-axis data
                linear_df['Y'],  # Y-axis data
                s=linear_df['Bubble Size'],  # Bubble size based on percentage
                c=linear_df['Percentage'],  # Color based on percentage
                cmap='viridis',  # Colormap for coloring bubbles
                alpha=0.7,  # Transparency
                edgecolors='black',  # Bubble border color
                linewidths=0.5  # Border thickness
            )

            # Add color bar for percentage
            cbar = plt.colorbar(scatter)
            cbar.set_label('Percentage')

            # Add text inside the bubbles (without % symbol)
            for i, row in linear_df.iterrows():
                plt.text(
                    row['X'],  # X position
                    row['Y'],  # Y position
                    f'{row["Percentage"]:.1f}',  # Display percentage without the '%' symbol
                    ha='center',  # Horizontal alignment
                    va='center',  # Vertical alignment
                    fontsize=8,  # Font size
                    color='black'  # Text color
                )

            # Ensure x_min and x_max are scalar values
            x_min = df[x_variable].min()  # Ensure it's a scalar
            x_max = df[x_variable].max()  # Ensure it's a scalar
            y_min = df[y_variable].min()  # Ensure it's a scalar
            y_max = df[y_variable].max()  # Ensure it's a scalar

            # Calculate 5% clearance for both axes
            x_range = x_max - x_min
            y_range = y_max - y_min
            x_clearance = 0.05 * x_range  # 5% clearance
            y_clearance = 0.05 * y_range  # 5% clearance

            # Set new limits with clearance
            x_min -= x_clearance
            x_max += x_clearance
            y_min -= y_clearance
            y_max += y_clearance

            # Set plot labels and title
            plt.xlabel(f'{x_variable[0]} ({x_variable[1]})', fontsize=14)
            plt.ylabel(f'{y_variable[0]} ({y_variable[1]})', fontsize=14)
            plt.title(f'{x_variable[0]} vs {y_variable[0]}', fontsize=16)
            plt.xlim(x_min, x_max)
            plt.ylim(y_min, y_max)

            # Add grid for better readability
            plt.grid(True, linestyle='--', alpha=0.5)

            # Save the plot
            output_directory = os.path.join(os.path.dirname(file_path), "Charts")
            output_file = os.path.join(output_directory, f'{y_variable[0]}_vs_{x_variable[0]}.png')
            plt.savefig(output_file, dpi=300)
            plt.close()
    except Exception as e:
        print(f"Error generating bubble chart: {e}")
    sys.exit()

# Function to run the selected plot type
def run_plot():
    global sheet_names, file_path, x_variable, y_variables, plot_type, bins

    if not file_path:
        messagebox.showerror("Error", "File path is not set.")
        return
    if not sheet_names:
        messagebox.showerror("Error", "No sheets selected.")
        return

    output_directory = os.path.join(os.path.dirname(file_path), "Charts")
    os.makedirs(output_directory, exist_ok=True)  # Ensure output directory exists

    if plot_type in ["FTP", "Duty Cycle"]:
        if not x_variable or not y_variables:
            messagebox.showerror("Error", "Please select both X and Y variables.")
            return
        compare_multiple_sheets_plot(file_path, sheet_names, x_variable, y_variables, output_directory, plot_type)

    elif plot_type == "Bubble Chart":
        if not x_variable or not y_variables:
            messagebox.showerror("Error", "Please select both X and Y variables.")
            return
        generate_bubble_chart(file_path, sheet_names, x_variable, y_variables, output_directory, bins)

    elif plot_type == "Pie Chart":
        if not y_variables:
            messagebox.showerror("Error", "Please select at least one Y variable for the pie chart.")
            return
        generate_pie_chart(file_path, sheet_names, y_variables, output_directory)

    else:
        messagebox.showerror("Error", "Invalid plot type selected.")
# Function to generate FTP & Duty Cycle plots
def compare_multiple_sheets_plot(file_path, sheet_names, x_variable, y_variables, output_directory, plot_type):
    os.makedirs(output_directory, exist_ok=True)
    dfs = [pd.read_excel(file_path, sheet_name=sheet, header=[0, 1]) for sheet in sheet_names]

    for df in dfs:
        df.columns = pd.MultiIndex.from_tuples([(col[0].strip(), col[1].strip()) for col in df.columns])

    x_values = [df[x_variable] for df in dfs]

    for col in y_variables:
        all_y_values = []
        plt.figure(figsize=(10, 6))
        for i, df in enumerate(dfs):
            if plot_type == 'FTP':
                plt.plot(x_values[i], df[col], linestyle='-', marker='o', linewidth=2, label=f'{sheet_names[i]}')
            elif plot_type == 'Duty Cycle':
                plt.plot(x_values[i], df[col], linestyle='-', marker='', linewidth=2, label=f'{sheet_names[i]}')  # No marker for Duty Cycle

            all_y_values.extend(df[col].dropna().values)

        all_x_values = [val for sublist in x_values for val in sublist]
        x_margin = (max(all_x_values) - min(all_x_values)) * 0.05
        plt.xlim(min(all_x_values) - x_margin, max(all_x_values) + x_margin)

        y_margin = (max(all_y_values) - min(all_y_values)) * 0.05
        plt.ylim(min(all_y_values) - y_margin, max(all_y_values) + y_margin)

        plt.xlabel(f'{x_variable[0]} ({x_variable[1]})', fontsize=14)
        plt.ylabel(f'{col[0]} ({col[1]})', fontsize=14)
        plt.title(f'{col[0]} ({col[1]})', fontsize=16)
        plt.gca().ticklabel_format(style='plain')  # Ensure normal number format
        plt.grid(True, linestyle="--")

        plt.legend()
        plt.savefig(os.path.join(output_directory, f"{col[0]}.png"))
        plt.close()
    sys.exit()

# Start the program
root = tk.Tk()
root.withdraw()
select_file()
root.mainloop()
