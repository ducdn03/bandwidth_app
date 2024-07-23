import json
import subprocess
import time
from time import sleep
import tkinter as tk
from functools import partial
from tkinter import messagebox
from tkinter import ttk
from tkinter.filedialog import asksaveasfilename

import openpyxl
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
from matplotlib.figure import Figure
from openpyxl.chart import Reference, LineChart


class BandwidthTest(tk.Tk):
    def __init__(self, server='89.187.162.1', port=5201, duration=10, iterations=1, stream=2):
        super().__init__()
        self.upl = []
        self.dowl = []
        self.server = server
        self.port = port
        self.duration = duration
        self.iterations = iterations
        self.stream = stream
        self.test_results = []
        self.window = None
        self.ServerChoosen = None
        self.InterationChoosen = None
        self.StreamChoosen = None
        self.PortChoosen = None
        self.progress = None
        self.title("Bandwidth Test")
        self.geometry("800x600")
        self.create_widget()

    def create_widget(self):

        menubar = tk.Menu(self)

        option = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='option', menu=option)
        option.add_command(label="BW Test", command=self.bandwidth_test)
        option.add_command(label="Export as Excel", command=self.export_bandwidth_test_to_excel)
        option.add_command(label="Configure Server", command=self.configure_setting)
        option.add_separator()
        option.add_command(label="Exit", command=self.destroy)

        self.config(menu=menubar)

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
            else:
                for interval in result_data['intervals']:
                    for i, stream in enumerate(interval['streams']):
                        if len(self.test_results) <= i:
                            self.test_results.append([])  # Ensure there is a list for each stream
                        if reverse:
                            self.test_results[i].append({
                                'received_Mbps': stream['bits_per_second'] / 1e6,
                            })
                        else:
                            self.test_results[i].append({
                                'sent_Mbps': stream['bits_per_second'] / 1e6,
                            })
        except subprocess.CalledProcessError as e:
            self.test_results.append({'error': str(e)})
        except json.JSONDecodeError:
            self.test_results.append({'error': 'Failed to parse JSON output from iperf3'})

    def run_multiple_tests(self, window):
        self.test_results.clear()
        self.upl.clear()
        self.dowl.clear()
        error_cnt = 0
        start_time = time.time()
        for i in range(self.iterations):
            self.progress['value'] = i * (100 / self.iterations)
            window.update_idletasks()
            self.run_iperf3_test(False)
            sleep(1)
            self.run_iperf3_test(True)

        end_time = time.time()
        print(end_time - start_time)
        self.progress['value'] = 100
        window.update_idletasks()

        for i, stream_result in enumerate(self.test_results):
            self.upl.append([])
            self.dowl.append([])
            for res in stream_result:
                if 'sent_Mbps' in res:
                    self.upl[i].append(round(res["sent_Mbps"], 1))
                elif 'received_Mbps' in res:
                    self.dowl[i].append(round(res["received_Mbps"], 1))
                else:
                    error_cnt += 1

        print(time.time() - end_time)
        if error_cnt >= (self.iterations * self.duration / 5):
            window.message = messagebox.showerror(title="test state", message="test failed")
            window.destroy()
            return
        window.message = messagebox.showinfo(title="test state", message="successfully test")

        for i, upl_res in enumerate(self.upl):
            if len(upl_res) != 0:
                self.display_graph_plot(i=i, upl=self.upl[i], dowl=self.dowl[i])

    def display_graph_plot(self, i, upl, dowl):
        window = tk.Tk()
        window.title(f"Stream {i + 1}")
        window.geometry('800x600')
        fig = Figure(figsize=(5, 5), dpi=80)
        plot1 = fig.add_subplot(111)

        plot1.plot(upl, label='Upload')
        plot1.plot(dowl, label='Download')
        plot1.legend()

        canvas = FigureCanvasTkAgg(fig, master=window)
        canvas.draw()
        canvas.get_tk_widget().pack()

        toolbar = NavigationToolbar2Tk(canvas, window)
        toolbar.update()
        canvas.get_tk_widget().pack()

        window.result_text = tk.Text(window, height=10, width=50)
        window.result_text.pack()
        average_upl, average_dowl = self.average_bandwidth(upl=upl, dowl=dowl)
        if average_upl == 'error' and average_dowl == ' error':
            window.message = messagebox.showerror(title="average bandwidth state", message="average bandwidth error")
            window.destroy()
            return
        server = f"{self.server}"
        window.result_text.insert(tk.END, f"Upload: {average_upl} Mbps\n"
                                          f"Dowload: {average_dowl} Mbps\n"
                                          f"Server: {server}\n")

    def bandwidth_test(self):
        self.progress = ttk.Progressbar(self.window, orient="horizontal", length=100, mode='determinate')
        self.progress.pack(pady=10)

        start_button = tk.Button(self.window, text='Start', command=partial(self.run_multiple_tests, self.window))
        start_button.pack(pady=10)

    def export_bandwidth_test_to_excel(self):
        wb = openpyxl.Workbook()
        for i in range(len(self.upl)):
            sheet = wb.create_sheet(title=f"Stream {i + 1}")
            sheet['A1'].value = "Upload (Mbps)"
            sheet['B1'].value = "Download (Mbps)"
            row = 2

            for data in self.upl[i]:
                sheet[f'A{row}'].value = data
                row += 1

            row = 2
            for data in self.dowl[i]:
                sheet[f'B{row}'].value = data
                row += 1

            average_upl, average_dowl = self.average_bandwidth(self.upl[i], self.dowl[i])
            sheet['D5'].value = "Average Upload: "
            sheet['D6'].value = "Average Download: "

            sheet['E5'].value = "{:.2f}".format(average_upl)
            sheet['E6'].value = "{:.2f}".format(average_dowl)

            # Create upload chart
            upload_chart = LineChart()
            upload_chart.title = "Upload Chart"
            upload_chart.x_axis.title = "Times"
            upload_chart.y_axis.title = "Mbps"
            upload_values = Reference(sheet, min_col=1, min_row=2, max_row=row - 1)
            upload_chart.add_data(upload_values, titles_from_data=True)
            sheet.add_chart(upload_chart, "G2")

            # Create download chart
            download_chart = LineChart()
            download_chart.title = "Download Chart"
            download_chart.x_axis.title = "Times"
            download_chart.y_axis.title = "Mbps"
            download_values = Reference(sheet, min_col=2, min_row=2, max_row=row - 1)
            download_chart.add_data(download_values, titles_from_data=True)
            sheet.add_chart(download_chart, "G20")

        # Remove default sheet created with workbook
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])

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

        tk.Label(self.window, text="Enter Server IP :", font=("Times New Roman", 14)).grid(row=1, column=0)
        self.ServerChoosen = tk.Entry(self.window, font=("Times New Roman", 14))
        self.ServerChoosen.grid(column=1, row=1)

        tk.Label(self.window, text="Enter No. Iteration :", font=("Times New Roman", 14)).grid(row=2, column=0)
        self.InterationChoosen = tk.Entry(self.window, font=("Times New Roman", 14))
        self.InterationChoosen.grid(column=1, row=2)

        tk.Label(self.window, text="Enter No. Stream :", font=("Times New Roman", 14)).grid(row=3, column=0)
        self.StreamChoosen = tk.Entry(self.window, font=("Times New Roman", 14))
        self.StreamChoosen.grid(column=1, row=3)

        tk.Label(self.window, text="Enter No. Port :", font=("Times New Roman", 14)).grid(row=4, column=0)
        self.PortChoosen = tk.Entry(self.window, font=("Times New Roman", 14))
        self.PortChoosen.grid(column=1, row=4)

        save_button = tk.Button(self.window, text="Save", command=partial(self.save_seletion, self.window))
        save_button.grid(column=1, row=5)

    def save_seletion(self, window):
        self.server = self.ServerChoosen.get()
        self.iterations = int(self.InterationChoosen.get())
        self.stream = int(self.StreamChoosen.get())
        self.port = int(self.PortChoosen.get())


if __name__ == '__main__':
    app = BandwidthTest()
    app.mainloop()
