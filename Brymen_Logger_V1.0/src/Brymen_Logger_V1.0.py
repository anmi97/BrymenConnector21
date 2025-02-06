#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import serial
import serial.tools.list_ports
import threading
import re
import time
import os

def list_com_ports():
    """
    Return a list of available COM ports with detailed information.
    For each port, include the device name, description, and hardware ID.
    """
    ports = serial.tools.list_ports.comports()
    #return [f"{port.device} - {port.description} ({port.hwid})" for port in ports]
    return [f"{port.device} - {port.description}" for port in ports]

def extract_device_from_selection(selection):
    """
    Given a selection string in the format:
       "COM3 - USB-SERIAL CH340 (USB VID:PID=1A86:7523 SER=XYZ)"
    this function returns just the COM port device (e.g., "COM3").
    """
    return selection.split(" - ")[0]

def parse_all_measurements(message):
    """
    Parses a serial message produced by the device.
    
    The message is assumed to be formatted as follows:
      [FLAGS ...] MAIN: <main_value><space><main_unit> [DC|AC] AUX: <aux_value><space><aux_unit> [AC] [LOW BAT]
      
    For messages that don’t contain 'MAIN:' the parser will simply treat the whole string as raw data.
    
    Returns a dictionary with keys:
      - flags: a list of flag strings (if any)
      - main_value: a float (if parsed) or None
      - main_unit: a string (if parsed) or None
      - aux_value: a float (if parsed) or None
      - aux_unit: a string (if parsed) or None
      - low_battery: Boolean flag (True if LOW BAT is found)
      - raw: if no structured message is detected, the raw message string.
    """
    parsed = {}
    # Look for the "MAIN:" token.
    if "MAIN:" in message:
        # Everything before MAIN: is assumed to be flags.
        pre_main, post_main = message.split("MAIN:", 1)
        flags = pre_main.strip().split()
        parsed["flags"] = flags

        # Parse main measurement.
        # Expect a number (with optional sign/decimal) and a unit made of letters or symbols.
        main_regex = re.compile(r"\s*([+-]?\d+(?:\.\d+)?)\s*([\wΩ%]+)")
        m_main = main_regex.search(post_main)
        if m_main:
            try:
                parsed["main_value"] = float(m_main.group(1))
            except ValueError:
                parsed["main_value"] = None
            parsed["main_unit"] = m_main.group(2)
        else:
            parsed["main_value"] = None
            parsed["main_unit"] = None

        # Look for the "AUX:" token and parse the auxiliary measurement.
        if "AUX:" in post_main:
            _, aux_part = post_main.split("AUX:", 1)
            m_aux = main_regex.search(aux_part)
            if m_aux:
                try:
                    parsed["aux_value"] = float(m_aux.group(1))
                except ValueError:
                    parsed["aux_value"] = None
                parsed["aux_unit"] = m_aux.group(2)
            else:
                parsed["aux_value"] = None
                parsed["aux_unit"] = None
        else:
            parsed["aux_value"] = None
            parsed["aux_unit"] = None

        # Check for low battery indicator.
        parsed["low_battery"] = "LOW BAT" in message
    else:
        # If the message does not contain "MAIN:" we assume it is raw data.
        parsed["raw"] = message
    return parsed

class SerialReaderThread(threading.Thread):
    """
    Thread that continuously reads from an open serial connection.
    It updates the GUI with the parsed measurement and logs all incoming data.
    If CSV logging is enabled, it writes the parsed data to the CSV file.
    """
    def __init__(self, ser, update_callback, log_callback):
        super().__init__()
        self.ser = ser
        self.update_callback = update_callback  # function(parsed_data: dict)
        self.log_callback = log_callback        # function(text: str)
        self._stop_event = threading.Event()
        self.logging_enabled = False
        self.csv_file = None
        self.csv_path = None

    def run(self):
        while not self._stop_event.is_set():
            try:
                if self.ser.in_waiting:
                    line = self.ser.readline().decode(errors="replace").strip()
                    if line:
                        # Log the raw line.
                        self.log_callback("RAW: " + line)
                        
                        # Parse the line.
                        parsed_data = parse_all_measurements(line)
                        
                        # Build a human-readable log string.
                        if "raw" in parsed_data:
                            log_str = "Raw Data: " + parsed_data["raw"]
                        else:
                            flags_str = " ".join(parsed_data.get("flags", []))
                            main_val = parsed_data.get("main_value")
                            main_unit = parsed_data.get("main_unit") or ""
                            aux_val = parsed_data.get("aux_value")
                            aux_unit = parsed_data.get("aux_unit") or ""
                            battery_str = "LOW BAT" if parsed_data.get("low_battery") else "OK"
                            log_str = (f"Flags: [{flags_str}] | "
                                       f"MAIN: {main_val} {main_unit} | "
                                       f"AUX: {aux_val} {aux_unit} | "
                                       f"Battery: {battery_str}")
                        
                        # Update the GUI display with the main measurement if available.
                        if "main_value" in parsed_data and parsed_data["main_value"] is not None:
                            self.update_callback(parsed_data["main_value"])
                        
                        # Log the complete parsed message.
                        self.log_callback(log_str)
                        
                        # If CSV logging is enabled, write a CSV line.
                        if self.logging_enabled:
                            if self.csv_file is None and self.csv_path:
                                try:
                                    os.makedirs(os.path.dirname(self.csv_path), exist_ok=True)
                                    self.csv_file = open(self.csv_path, "a", encoding="utf-8")
                                except Exception as e:
                                    self.log_callback(f"CSV Open Error: {e}")
                            if self.csv_file:
                                # CSV columns: timestamp, flags, main_value, main_unit, aux_value, aux_unit, battery
                                ts = time.strftime("%Y-%m-%d %H:%M:%S")
                                csv_line = (f"{ts},\"{flags_str}\","
                                            f"{main_val if main_val is not None else ''},\"{main_unit}\","
                                            f"{aux_val if aux_val is not None else ''},\"{aux_unit}\","
                                            f"{battery_str}\n")
                                self.csv_file.write(csv_line)
                                self.csv_file.flush()
            except Exception as e:
                print("Error in serial reading:", e)
            time.sleep(0.01)
        if self.csv_file:
            self.csv_file.close()

    def stop(self):
        self._stop_event.set()

    def set_logging(self, enabled, csv_path=None):
        """Enable or disable CSV logging. If enabling, provide the CSV file path."""
        self.logging_enabled = enabled
        if enabled:
            self.csv_path = csv_path
        else:
            if self.csv_file:
                self.csv_file.close()
                self.csv_file = None

class SerialGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Serial Communication & Logger")

        # Row 0: COM Port Selection & Refresh
        self.com_label = tk.Label(root, text="Select COM Port:")
        self.com_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.com_var = tk.StringVar()
        self.combobox = ttk.Combobox(root, textvariable=self.com_var, state="readonly")
        self.combobox['values'] = list_com_ports()
        self.combobox.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        # Refresh list automatically on click.
        self.combobox.bind("<Button-1>", lambda e: self.refresh_com_ports())
        if self.combobox['values']:
            self.combobox.current(0)

        self.refresh_button = tk.Button(root, text="Refresh", command=self.refresh_com_ports)
        self.refresh_button.grid(row=0, column=2, padx=5, pady=5)

        # Row 1: CSV File Destination
        self.csv_label = tk.Label(root, text="CSV File:")
        self.csv_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        self.csv_path_var = tk.StringVar()
        self.csv_entry = tk.Entry(root, textvariable=self.csv_path_var, width=40)
        self.csv_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        self.browse_button = tk.Button(root, text="Browse...", command=self.browse_csv)
        self.browse_button.grid(row=1, column=2, padx=5, pady=5)

        # Row 2: Connection Controls
        self.connect_button = tk.Button(root, text="Connect", command=self.toggle_connection)
        self.connect_button.grid(row=2, column=0, columnspan=3, padx=5, pady=5)

        # Row 3: Logging Controls
        self.logging_button = tk.Button(root, text="Start Logging", command=self.toggle_logging, state="disabled")
        self.logging_button.grid(row=3, column=0, columnspan=3, padx=5, pady=5)

        # Row 4: Latest MAIN Measurement Display
        self.measurement_label = tk.Label(root, text="Latest MAIN Measurement:", font=("Helvetica", 12))
        self.measurement_label.grid(row=4, column=0, padx=5, pady=5, sticky="w")

        self.measurement_var = tk.StringVar(value="N/A")
        self.measurement_display = tk.Label(root, textvariable=self.measurement_var,
                                            font=("Helvetica", 16), fg="blue")
        self.measurement_display.grid(row=4, column=1, padx=5, pady=5, sticky="w")

        # Row 5: Serial Log Label and Autoscroll Checkbutton
        self.log_label = tk.Label(root, text="Serial Log:")
        self.log_label.grid(row=5, column=0, padx=5, pady=(10, 0), sticky="w")

        self.autoscroll_var = tk.BooleanVar(value=True)
        self.autoscroll_check = tk.Checkbutton(root, text="Autoscroll", variable=self.autoscroll_var)
        self.autoscroll_check.grid(row=5, column=2, padx=5, pady=(10, 0), sticky="e")

        # Row 6: Serial Log Window (Scrolled Text)
        self.log_text = scrolledtext.ScrolledText(root, width=50, height=10, state="disabled")
        self.log_text.grid(row=6, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")

        # Row 7: Command Sending Controls
        self.command_label = tk.Label(root, text="Send Command:")
        self.command_label.grid(row=7, column=0, padx=5, pady=5, sticky="w")

        self.command_entry = tk.Entry(root, width=30)
        self.command_entry.grid(row=7, column=1, padx=5, pady=5, sticky="ew")
        # Bind the Return key to send the command.
        self.command_entry.bind("<Return>", self.send_command)

        self.send_button = tk.Button(root, text="Send", command=self.send_command, state="disabled")
        self.send_button.grid(row=7, column=2, padx=5, pady=5)

        # Internal state.
        self.serial_conn = None           # serial.Serial instance when connected.
        self.reader_thread = None         # SerialReaderThread instance.
        self.connected = False            # Connection state.
        self.logging_active = False       # Logging state.

        # Configure grid weights for resizing.
        root.columnconfigure(1, weight=1)
        root.rowconfigure(6, weight=1)

        # Cleanup on closing.
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def refresh_com_ports(self):
        """Update the COM port list with detailed information."""
        ports = list_com_ports()
        self.combobox['values'] = ports
        if ports:
            self.combobox.current(0)
        else:
            self.combobox.set('')

    def browse_csv(self):
        file_path = filedialog.asksaveasfilename(
            title="Select CSV file destination",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.csv_path_var.set(file_path)

    def toggle_connection(self):
        if not self.connected:
            self.connect()
        else:
            self.disconnect()

    def connect(self):
        selection = self.com_var.get()
        if not selection:
            messagebox.showerror("Input Error", "Please select a COM port.")
            return

        # Extract just the COM device (e.g., "COM3") from the selection string.
        com_port = extract_device_from_selection(selection)
        try:
            self.serial_conn = serial.Serial(com_port, 115200, timeout=1)
        except Exception as e:
            messagebox.showerror("Serial Error", f"Could not open {com_port}:\n{e}")
            return

        self.connected = True
        self.connect_button.config(text="Disconnect")
        self.send_button.config(state="normal")
        self.logging_button.config(state="normal")
        messagebox.showinfo("Connected", f"Connected to {com_port}")

        # Start the reader thread.
        self.reader_thread = SerialReaderThread(self.serial_conn,
                                                  self.update_measurement,
                                                  self.append_log)
        self.reader_thread.daemon = True
        self.reader_thread.start()

    def disconnect(self):
        if self.logging_active:
            self.stop_logging()
        if self.reader_thread:
            self.reader_thread.stop()
            self.reader_thread.join()
            self.reader_thread = None
        if self.serial_conn and self.serial_conn.is_open:
            self.serial_conn.close()
        self.serial_conn = None
        self.connected = False
        self.connect_button.config(text="Connect")
        self.logging_button.config(state="disabled")
        self.send_button.config(state="disabled")
        messagebox.showinfo("Disconnected", "Serial connection closed.")

    def toggle_logging(self):
        if not self.logging_active:
            self.start_logging()
        else:
            self.stop_logging()

    def start_logging(self):
        if not self.connected or self.serial_conn is None:
            messagebox.showerror("Not Connected", "Please connect to a serial port first.")
            return
        csv_path = self.csv_path_var.get().strip()
        if not csv_path:
            messagebox.showerror("Input Error", "Please select a CSV file destination.")
            return

        self.logging_active = True
        self.logging_button.config(text="Stop Logging")
        if self.reader_thread:
            self.reader_thread.set_logging(True, csv_path)

    def stop_logging(self):
        if self.reader_thread:
            self.reader_thread.set_logging(False)
        self.logging_active = False
        self.logging_button.config(text="Start Logging")

    def update_measurement(self, measurement):
        # Update the main measurement display.
        self.root.after(0, lambda: self.measurement_var.set(f"{measurement:.3f}"))

    def append_log(self, text):
        def inner():
            self.log_text.config(state="normal")
            self.log_text.insert(tk.END, text + "\n")
            if self.autoscroll_var.get():
                self.log_text.yview(tk.END)
            self.log_text.config(state="disabled")
        self.root.after(0, inner)

    def send_command(self, event=None):
        command = self.command_entry.get().strip()
        if command == "":
            messagebox.showerror("Input Error", "Please enter a command to send.")
            return
        if self.serial_conn and self.serial_conn.is_open:
            try:
                self.serial_conn.write((command + "\n").encode())
            except Exception as e:
                print("Error sending command:", e)
        else:
            messagebox.showerror("Not Connected", "Serial connection is not active.")

        self.append_log(f"Sent: {command}")
        self.command_entry.delete(0, tk.END)

    def on_closing(self):
        if self.logging_active:
            self.stop_logging()
        if self.connected:
            self.disconnect()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    gui = SerialGUI(root)
    root.mainloop()
