from py_canoe import CANoe
from robot.api import logger
import threading
import time
import os
import re
import xlsxwriter
import subprocess

class canoe_robot_lib:
    def __init__(self):
        self._kill_zombie_canoe()
        self.canoe_inst = CANoe(clean_gen_py_cache=True)
        self.log_file = os.path.join(os.getcwd(), 'canoe_write_log.txt')
        self.excel_file = os.path.join(os.getcwd(), 'uds_frames_log.xlsx')
        self._stop_logging = threading.Event()
        self._log_thread = None
        self.workbook = None
        self.worksheet = None
        self.row = 1

    def _kill_zombie_canoe(self):
        try:
            subprocess.run(['taskkill', '/F', '/IM', 'CANoe64.exe', '/T'], capture_output=True)
            subprocess.run(['taskkill', '/F', '/IM', 'CANoe32.exe', '/T'], capture_output=True)
            # Wait a moment for OS to release resources/file locks
            time.sleep(1)
        except Exception as e:
            logger.warn(f"Failed to kill existing CANoe processes: {e}")

    def open_canoe_configuration(self, cfg_path):
        # Clear old log files if they exist to ensure fresh start
        for f_path in [self.log_file, self.excel_file]:
            if os.path.exists(f_path):
                try:
                    os.remove(f_path)
                except Exception as e:
                    # If remove fails, try to truncate the text log
                    if f_path == self.log_file:
                        try:
                            with open(f_path, 'w') as f:
                                f.truncate(0)
                        except:
                            pass
                    logger.warn(f"Could not delete old log file {f_path}: {e}. Ensure it is closed.")
        
        self.canoe_inst.open(canoe_cfg=cfg_path)
        # Enable write window logging to a file
        self.canoe_inst.enable_write_window_output_file(self.log_file)

    def _tail_log(self):
        # Ensure log file exists
        if not os.path.exists(self.log_file):
            with open(self.log_file, 'w') as f:
                pass
        
        # Regex to parse the frame data
        pattern = re.compile(r"Program / Model\tSend request: (0x[0-9A-Fa-f]{2}), (0x[0-9A-Fa-f]{2}), (0x[0-9A-Fa-f]{2}), (0x[0-9A-Fa-f]{2}), (0x[0-9A-Fa-f]{2}), (0x[0-9A-Fa-f]{2}), (0x[0-9A-Fa-f]{2}), (0x[0-9A-Fa-f]{2})")
        
        with open(self.log_file, 'r', errors='ignore') as f:
            # Go to end of file
            f.seek(0, os.SEEK_END)
            while not self._stop_logging.is_set():
                line = f.readline()
                if line:
                    logger.console(line.strip())
                    match = pattern.search(line)
                    if match and self.worksheet:
                        # Extract 8 bytes
                        bytes_data = match.groups()
                        # Write to Excel
                        self.worksheet.write(self.row, 0, time.strftime("%H:%M:%S"))
                        self.worksheet.write(self.row, 1, "0x7E4") # Constant for Dia_Req
                        for i, b in enumerate(bytes_data):
                            self.worksheet.write(self.row, i + 2, b)
                        self.row += 1
                else:
                    time.sleep(0.1)

    def start_canoe_measurement(self):
        # Initialize Excel (XlsxWriter will overwrite by default)
        try:
            self.workbook = xlsxwriter.Workbook(self.excel_file)
        except Exception as e:
            logger.error(f"Failed to create Excel file: {e}. Please ensure it is closed.")
            raise
            
        self.worksheet = self.workbook.add_worksheet("UDS Frames")
        headers = ["Timestamp", "CAN ID", "B0 (DLC)", "B1 (SID)", "B2", "B3", "B4", "B5", "B6", "B7"]
        for i, h in enumerate(headers):
            self.worksheet.write(0, i, h)
        self.row = 1

        self.canoe_inst.start_measurement()
        # Start log tailing thread
        self._stop_logging.clear()
        self._log_thread = threading.Thread(target=self._tail_log, daemon=True)
        self._log_thread.start()

    def run_test_module(self, test_module_name):
        self.canoe_inst.execute_test_module(test_module_name)

    def stop_canoe_measurement(self):
        # Stop logging thread
        self._stop_logging.set()
        if self._log_thread:
            self._log_thread.join(timeout=2)
        
        self.canoe_inst.stop_measurement()
        
        # Close Excel
        if self.workbook:
            self.workbook.close()
            self.workbook = None
            logger.console(f"Excel log saved to: {self.excel_file}")
        
        # Attempt to disable logging to release the file lock
        try:
            self.canoe_inst.enable_write_window_output_file("") # Try to release lock
        except:
            pass

        # Clean up log file
        if os.path.exists(self.log_file):
            try:
                os.remove(self.log_file)
            except:
                # If deletion still fails, try to truncate it so it's empty for the next viewing
                try:
                    with open(self.log_file, 'w') as f:
                        f.truncate(0)
                except:
                    pass

