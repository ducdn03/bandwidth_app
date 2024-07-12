import tkinter as tk
from tkinter import messagebox
import iperf3
import threading


class BandwidthTester(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bandwidth Tester")
        self.geometry("400x300")

        self.create_widgets()

    def create_widgets(self):
        # Server address
        self.server_label = tk.Label(self, text="Server Address:")
        self.server_label.pack(pady=5)

        self.server_entry = tk.Entry(self)
        self.server_entry.pack(pady=5)

        # Server port
        self.port_label = tk.Label(self, text="Server Port:")
        self.port_label.pack(pady=5)

        self.port_entry = tk.Entry(self)
        self.port_entry.pack(pady=5)

        # Test button
        self.test_button = tk.Button(self, text="Test Bandwidth", command=self.run_test)
        self.test_button.pack(pady=20)

        # Result display
        self.result_text = tk.Text(self, height=10, width=50)
        self.result_text.pack(pady=10)

    def run_test(self):
        server = self.server_entry.get()
        port = self.port_entry.get()

        if not server or not port:
            messagebox.showwarning("Input Error", "Please enter both server address and port.")
            return

        try:
            port = int(port)
        except ValueError:
            messagebox.showwarning("Input Error", "Port must be a number.")
            return

        threading.Thread(target=self.perform_test, args=(server, port)).start()

    def perform_test(self, server, port):
        self.result_text.delete('1.0', tk.END)
        client = iperf3.Client()
        client.server_hostname = server
        client.port = port

        result = client.run()

        if result.error:
            self.result_text.insert(tk.END, f"Error: {result.error}")
        else:
            self.result_text.insert(tk.END, f"Test Completed:\n"
                                            f"Sent: {result.sent_Mbps} Mbps\n"
                                            f"Received: {result.received_Mbps} Mbps\n")


if __name__ == "__main__":
    app = BandwidthTester()
    app.mainloop()
