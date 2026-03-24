import sys
import os
import time
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                             QTextEdit, QFileDialog, QMessageBox)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont

# Import the existing library
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from canoe_robot_lib import canoe_robot_lib

class WorkerThread(QThread):
    """
    A generic worker thread to run canoe_robot_lib methods securely in the background 
    without blocking the PyQt5 main GUI loop.
    """
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, action, *args, **kwargs):
        super().__init__()
        self.action = action
        self.args = args
        self.kwargs = kwargs

    def run(self):
        try:
            self.action(*self.args, **self.kwargs)
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))


class CanoeCIApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CANoe CI Automation Interface")
        self.resize(750, 550)
        
        self.canoe_lib = None
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout()
        central_widget.setLayout(layout)
        
        # Title
        title = QLabel("CANoe CI Automation Pipeline")
        title.setFont(QFont("Segoe UI", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        layout.addSpacing(15)

        # Config File Selection
        cfg_layout = QHBoxLayout()
        cfg_label = QLabel("CANoe Config (.cfg):")
        cfg_label.setFixedWidth(160)
        cfg_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        
        self.cfg_input = QLineEdit()
        self.cfg_input.setPlaceholderText("Select the path to the CANoe configuration file...")
        
        self.cfg_btn = QPushButton("Browse...")
        self.cfg_btn.setFixedWidth(100)
        self.cfg_btn.clicked.connect(self.browse_cfg)
        
        cfg_layout.addWidget(cfg_label)
        cfg_layout.addWidget(self.cfg_input)
        cfg_layout.addWidget(self.cfg_btn)
        layout.addLayout(cfg_layout)

        # Test Module Name
        tm_layout = QHBoxLayout()
        tm_label = QLabel("Test Module Name:")
        tm_label.setFixedWidth(160)
        tm_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        
        self.tm_input = QLineEdit()
        self.tm_input.setPlaceholderText("e.g. TC_Main or test_capl")
        
        tm_layout.addWidget(tm_label)
        tm_layout.addWidget(self.tm_input)
        
        layout.addLayout(tm_layout)
        layout.addSpacing(20)

        # Controls Layout
        controls_layout = QHBoxLayout()
        
        self.btn_start = QPushButton("🚀 1. Start CANoe")
        self.btn_start.setMinimumHeight(45)
        self.btn_start.clicked.connect(self.start_canoe)
        
        self.btn_run = QPushButton("⚡ 2. Run Test Module")
        self.btn_run.setMinimumHeight(45)
        self.btn_run.setEnabled(False)
        self.btn_run.clicked.connect(self.run_test_module)
        
        self.btn_stop = QPushButton("💾 3. Stop & Export Logs")
        self.btn_stop.setMinimumHeight(45)
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop_canoe)
        
        controls_layout.addWidget(self.btn_start)
        controls_layout.addWidget(self.btn_run)
        controls_layout.addWidget(self.btn_stop)
        layout.addLayout(controls_layout)

        layout.addSpacing(15)

        # Log Output
        log_label = QLabel("Execution Status Logs:")
        log_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        layout.addWidget(log_label)
        
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        # Modern Dark Console for logs
        self.log_output.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #4CAF50;
                font-family: Consolas, monospace;
                font-size: 13px;
                border-radius: 6px;
                padding: 10px;
            }
        """)
        layout.addWidget(self.log_output)
        
        self.log("[System] Application Ready.")

    def log(self, text):
        time_str = time.strftime("%H:%M:%S")
        self.log_output.append(f"[{time_str}] {text}")
        # Auto-scroll to bottom
        scrollbar = self.log_output.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def browse_cfg(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select CANoe Configuration", "", "CANoe Config (*.cfg);;All Files (*)")
        if file_path:
            self.cfg_input.setText(os.path.abspath(file_path))

    def disable_buttons(self):
        self.btn_start.setEnabled(False)
        self.btn_run.setEnabled(False)
        self.btn_stop.setEnabled(False)

    def start_canoe(self):
        cfg_path = self.cfg_input.text().strip()
        if not cfg_path or not os.path.exists(cfg_path):
            QMessageBox.warning(self, "Configuration Error", "Please browse and select a valid .cfg file first.")
            return

        self.log(f"[Action] Connecting to CANoe and loading config -> {os.path.basename(cfg_path)}...")
        self.disable_buttons()
        self.log("[Info] Clearing stale zombie CANoe instances and building COM objects. Please wait...")
        
        def _task():
            if not self.canoe_lib:
                self.canoe_lib = canoe_robot_lib()
            self.canoe_lib.open_canoe_configuration(cfg_path)
            self.canoe_lib.start_canoe_measurement()

        self.worker = WorkerThread(_task)
        self.worker.finished.connect(self.on_start_success)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_start_success(self):
        self.log("[Success] CANoe Configuration Loaded. Measurement is active.")
        self.btn_start.setEnabled(False)
        self.btn_run.setEnabled(True)
        self.btn_stop.setEnabled(True)

    def run_test_module(self):
        tm_name = self.tm_input.text().strip()
        if not tm_name:
            QMessageBox.warning(self, "Missing Module Name", "Please input the Test Module Name before executing.")
            return
            
        self.log(f"[Action] Triggering execution for Test Module: {tm_name} ...")
        self.disable_buttons()
        
        def _task():
            self.canoe_lib.run_test_module(tm_name)

        self.worker = WorkerThread(_task)
        self.worker.finished.connect(self.on_run_success)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_run_success(self):
        self.log(f"[Success] Test Module execution finished.")
        self.btn_start.setEnabled(False)
        self.btn_run.setEnabled(True)
        self.btn_stop.setEnabled(True)

    def stop_canoe(self):
        self.log("[Action] Stopping CANoe Measurement & Generating Excel Logs...")
        self.disable_buttons()
        
        def _task():
            if self.canoe_lib:
                self.canoe_lib.stop_canoe_measurement()

        self.worker = WorkerThread(_task)
        self.worker.finished.connect(self.on_stop_success)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_stop_success(self):
        self.log("[Success] Measurement Halted. UDS frames Excel log created successfully.")
        self.log("[Info] Ready for next test configuration.")
        self.btn_start.setEnabled(True)
        self.btn_run.setEnabled(False)
        self.btn_stop.setEnabled(False)
        # Release COM lock immediately to start fresh next time
        self.canoe_lib = None

    def on_error(self, err_msg):
        self.log(f"[Exception Encountered] {err_msg}")
        QMessageBox.critical(self, "Runtime Error", f"An error occurred:\n\n{err_msg}")
        # Decide recovery based on variables
        if self.canoe_lib is None:
            self.btn_start.setEnabled(True)
        else:
            self.btn_start.setEnabled(False)
            self.btn_run.setEnabled(True)
            self.btn_stop.setEnabled(True)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    app.setStyleSheet('''
    QWidget {
        font-family: "Segoe UI";
        font-size: 13px;
    }
    QPushButton {
        background-color: #0078D7;
        color: white;
        font-weight: bold;
        border: none;
        border-radius: 5px;
        padding: 8px 15px;
    }
    QPushButton:hover {
        background-color: #005A9E;
    }
    QPushButton:pressed {
        background-color: #004578;
    }
    QPushButton:disabled {
        background-color: #D3D3D3;
        color: #A9A9A9;
    }
    QLineEdit {
        padding: 8px;
        border: 1px solid #B0BEC5;
        border-radius: 4px;
        background-color: #FFFFFF;
    }
    QLineEdit:focus {
        border: 2px solid #0078D7;
    }
    ''')
    
    window = CanoeCIApp()
    window.show()
    sys.exit(app.exec_())
