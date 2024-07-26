import json
import subprocess
from time import sleep
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter.filedialog import asksaveasfilename

import openpyxl
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
from matplotlib.figure import Figure
from openpyxl.chart import Reference, LineChart


class BandwidthTest(tk.Tk):
    def __init__(self, server='192.168.2.235', port=5201, duration=10, iterations=1, stream=10):
        super().__init__()
        self.upl = []
        self.dowl = []
        self.server = server
        self.port = port
        self.duration = duration
        self.iterations = iterations
        self.stream = stream
        self.test_results = []
        self.ServerChoosen = None
        self.DurationChoosen = None
        self.StreamChoosen = None
        self.PortChoosen = None
        self.progress = None
        self.main_frame = None
        self.title("Bandwidth Test")
        self.geometry("800x600")
        self.create_widget()

    def create_widget(self):

        menubar = tk.Menu(self)

        option = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='Option', menu=option)
        option.add_command(label="BW Test", command=self.bandwidth_test)
        option.add_command(label="Export as Excel", command=self.export_bandwidth_test_to_excel)
        option.add_command(label="Configure Server", command=self.configure_setting)
        option.add_separator()
        option.add_command(label="Exit", command=self.destroy)

        self.config(menu=menubar)

        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def run_iperf3_test(self, reverse):
        command = [
            'iperf3',
            '-c', self.server,
            '-p', str(self.port),
            '-t', str(self.duration),
            '-P', str(self.stream),
            '-J',  # JSON output,
        ]

        if reverse:
            command.append('-R')

        try:
            result = subprocess.run(command, capture_output=True, text=True, check=True)
            result_data = json.loads(result.stdout)

            if 'error' in result_data:
                self.test_results.append({'error': result_data['error']})
                messagebox.showerror(title='test state', message=f"error: {result_data['error']}")
            else:
                for interval in result_data['intervals']:
                    if reverse:
                        self.test_results.append({
                            'received_Mbps': interval['sum']['bits_per_second'] / 1e6,
                            })
                    else:
                        self.test_results.append({
                            'sent_Mbps': interval['sum']['bits_per_second'] / 1e6,
                        })
        except subprocess.CalledProcessError as e:
            self.test_results.append({'error': str(e)})
        except json.JSONDecodeError:
            self.test_results.append({'error': 'Failed to parse JSON output from iperf3'})

    def run_multiple_tests(self):
        self.test_results.clear()
        self.upl.clear()
        self.dowl.clear()
        error_cnt = 0
        self.run_iperf3_test(False)
        self.progress['value'] = 50
        self.update_idletasks()
        sleep(2)
        self.run_iperf3_test(True)
        self.progress['value'] = 100
        self.update_idletasks()

        for result in self.test_results:
            if 'sent_Mbps' in result:
                self.upl.append(round(result['sent_Mbps'], 1))
            elif 'received_Mbps' in result:
                self.dowl.append(round(result['received_Mbps'], 1))
            else:
                error_cnt += 1

        if error_cnt >= (self.iterations * self.duration / 5):
            messagebox.showerror(title="Test State", message="Test failed")
            return
        messagebox.showinfo(title="Test State", message="Test successfully completed")
        self.display_graph_plot(upl=self.upl, dowl=self.dowl)

    def display_graph_plot(self, upl, dowl):
        for widget in self.main_frame.winfo_children():
            widget.destroy()
        fig = Figure(figsize=(5, 5), dpi=80)
        plot1 = fig.add_subplot(111)

        plot1.plot(upl, label='Upload')
        plot1.plot(dowl, label='Download')
        plot1.legend()

        canvas = FigureCanvasTkAgg(fig, master=self.main_frame)
        canvas.draw()
        canvas.get_tk_widget().pack()
        plot1.set_xlabel('time (t)')
        plot1.set_ylabel('Speed (Mbps)')
        toolbar = NavigationToolbar2Tk(canvas, self.main_frame)
        toolbar.update()
        canvas.get_tk_widget().pack()

        result_text = tk.Text(self.main_frame, height=10, width=50)
        result_text.pack()
        average_upl, average_dowl = self.average_bandwidth(upl=upl, dowl=dowl)
        if average_upl == 'error' and average_dowl == 'error':
            messagebox.showerror(title="Average Bandwidth State", message="Average bandwidth error")
            return
        server = f"{self.server}"
        result_text.insert(tk.END, f"Upload: {average_upl} Mbps\n"
                                   f"Download: {average_dowl} Mbps\n"
                                   f"Server: {server}\n"
                                   f"Port: {self.port}\n"
                                   f"Steam: {self.stream}\n"
                                   f"Duration: {self.duration}\n")

    def bandwidth_test(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        self.progress = ttk.Progressbar(self.main_frame, orient="horizontal", length=100, mode='determinate')
        self.progress.pack(pady=10)

        start_button = tk.Button(self.main_frame, text='Start', command=self.run_multiple_tests)
        start_button.pack(pady=10)

    def export_bandwidth_test_to_excel(self):
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "Bandwidth Test Result"
        sheet['A1'].value = "Upload (Mbps)"
        sheet['B1'].value = "Download (Mbps)"
        row = 2
        for data in self.upl:
            sheet[f'A{row}'].value = data
            row += 1

        row = 2
        for data in self.dowl:
            sheet[f'B{row}'].value = data
            row += 1

        average_upl, average_dowl = self.average_bandwidth(upl=self.upl, dowl=self.dowl)
        sheet['D5'].value = "Average Upload: "
        sheet['D6'].value = "Average Download: "

        sheet['E5'].value = "{:.2f}".format(average_upl)
        sheet['E6'].value = "{:.2f}".format(average_dowl)

        # Create upload chart
        upload_chart = LineChart()
        upload_chart.title = "Upload Chart"
        upload_chart.x_axis.title = "Times"
        upload_chart.y_axis.title = "Mbps"
        upload_values = Reference(sheet, min_col=1, max_col=1, min_row=1, max_row=len(self.upl) + 1)
        upload_chart.add_data(upload_values, titles_from_data=True)
        sheet.add_chart(upload_chart, "G2")

        # Create download chart
        download_chart = LineChart()
        download_chart.title = "Download Chart"
        download_chart.x_axis.title = "Times"
        download_chart.y_axis.title = "Mbps"
        download_values = Reference(sheet, min_col=2, max_col=2, min_row=1, max_row=len(self.dowl) + 1)
        download_chart.add_data(download_values, titles_from_data=True)
        sheet.add_chart(download_chart, "G20")

        files = [('Excel Files', '*.xlsx')]
        save_path = asksaveasfilename(filetypes=files)
        if save_path:
            wb.save(save_path)
            messagebox.showinfo(title="Export state", message="Export Completed")

    @staticmethod
    def average_bandwidth(upl, dowl):
        if len(upl) == 0 or len(dowl) == 0:
            return 'error', 'error'
        average_upl = sum(upl) / len(upl)
        average_dowl = sum(dowl) / len(dowl)
        return average_upl, average_dowl

    def configure_setting(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        tk.Label(self.main_frame, text="Enter Server IP:", font=("Times New Roman", 14)).grid(row=1, column=0)
        self.ServerChoosen = tk.Entry(self.main_frame, font=("Times New Roman", 14))
        self.ServerChoosen.grid(column=1, row=1)

        tk.Label(self.main_frame, text="Enter No. Duration:", font=("Times New Roman", 14)).grid(row=2, column=0)
        self.DurationChoosen = tk.Entry(self.main_frame, font=("Times New Roman", 14))
        self.DurationChoosen.grid(column=1, row=2)

        tk.Label(self.main_frame, text="Enter No. Stream:", font=("Times New Roman", 14)).grid(row=3, column=0)
        self.StreamChoosen = tk.Entry(self.main_frame, font=("Times New Roman", 14))
        self.StreamChoosen.grid(column=1, row=3)

        tk.Label(self.main_frame, text="Enter No. Port:", font=("Times New Roman", 14)).grid(row=4, column=0)
        self.PortChoosen = tk.Entry(self.main_frame, font=("Times New Roman", 14))
        self.PortChoosen.grid(column=1, row=4)

        save_button = tk.Button(self.main_frame, text="Save", command=self.save_selection)
        save_button.grid(column=1, row=5)

    def save_selection(self):
        self.server = self.ServerChoosen.get()
        self.duration = int(self.DurationChoosen.get())
        self.stream = int(self.StreamChoosen.get())
        self.port = int(self.PortChoosen.get())
        messagebox.showinfo(title='save state', message="save completedly")
        for widget in self.main_frame.winfo_children():
            widget.destroy()


if __name__ == '__main__':
    app = BandwidthTest()
    app.mainloop()
