import json
import subprocess
import time
from time import sleep
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter.filedialog import asksaveasfilename

import openpyxl
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
from matplotlib.figure import Figure
from openpyxl.chart import Reference, LineChart
import threading
from threading import Thread


class ThreadWithReturnValue(Thread):

    def __init__(self, group=None, target=None, name=None,
                 args=(), kwargs={}, Verbose=None):
        Thread.__init__(self, group, target, name, args, kwargs)
        self._return = None

    def run(self):
        if self._target is not None:
            self._return = self._target(*self._args,
                                        **self._kwargs)

    def join(self, *args):
        Thread.join(self, *args)
        return self._return


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
        self.spinner_label = None
        self.spinner_running = False
        self.main_frame = None
        self.title("Bandwidth Test")
        self.geometry("800x600")
        self.create_widget()

    def create_widget(self):
        menubar = tk.Menu(self)

        option = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='Option', menu=option)
        option.add_command(label="New Window", command=self.new_window)
        option.add_command(label="BW Test", command=self.bandwidth_test)
        option.add_command(label="Export as Excel", command=self.export_bandwidth_test_to_excel)
        option.add_command(label="Configure Server", command=self.configure_setting)
        option.add_separator()
        option.add_command(label="Exit", command=self.destroy)

        self.config(menu=menubar)

        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    @staticmethod
    def new_window():
        new = BandwidthTest()
        new.mainloop()

    def run_iperf3_test(self, reverse, stop_event):
        if stop_event.is_set():
            return

        command = [
            'iperf3',
            '-c', self.server,
            '-p', str(self.port),
            '-t', str(self.duration),
            '-P', str(self.stream),
            '-J',  # JSON output
        ]

        if reverse:
            command.append('-R')

        try:
            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            stdout, stderr = process.communicate()
            while not stop_event.is_set() and process.poll() is None:
                sleep(10)

            if stop_event.is_set():
                process.terminate()
                return

            result_data = json.loads(stdout)

            if 'error' in result_data:
                self.test_results.append({'error': result_data['error']})
                messagebox.showerror(title='test state', message=f"error: {result_data['error']}", parent=self)
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

    def check_server_status(self, stop_event):
        command = [
            'ping',
            '-c', '1',  # Send only 1 packet
            self.server,
        ]

        while not stop_event.is_set():
            try:
                start_time = time.time()
                process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                stdout, stderr = process.communicate()
                while (time.time() - start_time) < 3:
                    if process.poll() is None:
                        sleep(1)

                if ((time.time() - start_time) >= 3) and process.poll() is None:
                    process.terminate()
                    stop_event.set()
                    return

                # Check if 'bytes from' is in the output to determine success
                if 'bytes from' in stdout:
                    sleep(10)
                else:
                    print("Server did not respond to ping")
                    stop_event.set()
                    return
            except subprocess.CalledProcessError as e:
                print(f"Server check error: {str(e)}")
                stop_event.set()
                return

    def run_multiple_tests(self):
        def test_wrapper():
            self.test_results.clear()
            self.upl.clear()
            self.dowl.clear()
            error_cnt = 0

            stop_event = threading.Event()
            thread1 = threading.Thread(target=self.run_iperf3_test, args=(False, stop_event))
            thread2 = threading.Thread(target=self.run_iperf3_test, args=(True, stop_event))
            thread3 = threading.Thread(target=self.check_server_status, args=(stop_event,))

            thread1.start()
            thread3.start()

            thread1.join()
            sleep(1)
            thread2.start()
            thread2.join()

            stop_event.set()

            for result in self.test_results:
                if 'sent_Mbps' in result:
                    self.upl.append(round(result['sent_Mbps'], 1))
                elif 'received_Mbps' in result:
                    self.dowl.append(round(result['received_Mbps'], 1))
                elif 'server_status' in result:
                    if result['server_status'] == 'down':
                        print("Server is down")

            if (error_cnt >= (self.duration / 5)) or len(self.upl) == 0 or len(self.dowl) == 0:
                messagebox.showerror(title="Test State", message="Test failed", parent=self)
                self.display_graph_plot(upl=self.upl, dowl=self.dowl)
                self.stop_spinner()
                return
            messagebox.showinfo(title="Test State", message="Test successfully completed", parent=self)
            self.display_graph_plot(upl=self.upl, dowl=self.dowl)
            self.stop_spinner()

        self.start_spinner()
        threading.Thread(target=test_wrapper).start()

    def start_spinner(self):
        self.spinner_running = True
        self.update_spinner()

    def update_spinner(self):
        if self.spinner_running:
            current_text = self.spinner_label.cget("text")
            if current_text == "":
                self.spinner_label.config(text="|")
            elif current_text == "|":
                self.spinner_label.config(text="/")
            elif current_text == "/":
                self.spinner_label.config(text="-")
            elif current_text == "-":
                self.spinner_label.config(text="\\")
            else:
                self.spinner_label.config(text="|")
            self.after(100, self.update_spinner)

    def stop_spinner(self):
        self.spinner_running = False
        if self.spinner_label:
            self.spinner_label.config(text="")

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
            messagebox.showinfo(title="Export state", message="Export Completed", parent=self)

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
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

        result_text = tk.Text(self.main_frame, height=10, width=50)
        result_text.pack(fill=tk.BOTH, expand=1)
        average_upl, average_dowl = self.average_bandwidth(upl=upl, dowl=dowl)
        if average_upl == 'error' and average_dowl == 'error':
            messagebox.showerror(title="Average Bandwidth State", message="Average bandwidth error", parent=self)
            return
        result_text.insert(tk.END, f"Upload: {average_upl} Mbps\n"
                                   f"Download: {average_dowl} Mbps\n"
                                   f"Server: {self.server}\n"
                                   f"Port: {self.port}\n"
                                   f"Stream: {self.stream}\n"
                                   f"Duration: {self.duration}\n")

    def bandwidth_test(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        self.spinner_label = tk.Label(self.main_frame, text="", font=("Times New Roman", 50))
        self.spinner_label.pack(pady=10)

        start_button = tk.Button(self.main_frame, text='Start', command=self.run_multiple_tests)
        start_button.pack(pady=10)

    @staticmethod
    def average_bandwidth(upl, dowl):
        if len(upl) == 0 or len(dowl) == 0:
            return 'error', 'error'
        average_upl = sum(upl) / len(upl)
        average_dowl = sum(dowl) / len(dowl)
        return average_upl, average_dowl

    def test_bandwidth_10minute(self, frequency):
        def test_wrapper():
            self.test_results.clear()
            self.upl.clear()
            self.dowl.clear()
            self.duration = 300
            error_cnt = 0

            stop_event = threading.Event()
            thread1 = threading.Thread(target=self.run_iperf3_test, args=(False, stop_event))
            thread2 = threading.Thread(target=self.run_iperf3_test, args=(True, stop_event))
            thread3 = threading.Thread(target=self.check_server_status, args=(stop_event,))

            thread1.start()
            thread3.start()

            thread1.join()
            sleep(1)
            thread2.start()
            thread2.join()

            stop_event.set()

            for result in self.test_results:
                if 'sent_Mbps' in result:
                    self.upl.append(round(result['sent_Mbps'], 1))
                elif 'received_Mbps' in result:
                    self.dowl.append(round(result['received_Mbps'], 1))
                elif 'server_status' in result:
                    if result['server_status'] == 'down':
                        print("Server is down")

            if (error_cnt >= (self.duration / 5)) or len(self.upl) == 0 or len(self.dowl) == 0:
                self.stop_spinner()
                return
            messagebox.showinfo(title="Test State", message="Test successfully completed", parent=self)
            self.stop_spinner()

        self.start_spinner()
        threading.Thread(target=test_wrapper).start()

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
        server = self.ServerChoosen.get()
        duration = self.DurationChoosen.get()
        stream = self.StreamChoosen.get()
        port = self.PortChoosen.get()

        if len(server) > 0:
            self.server = server
        if duration:
            self.duration = int(duration)
        if stream:
            self.stream = int(stream)
        if port:
            self.port = int(port)

        messagebox.showinfo(title='save state', message="save completedly", parent=self)
        for widget in self.main_frame.winfo_children():
            widget.destroy()


if __name__ == '__main__':
    app = BandwidthTest()
    app.mainloop()
