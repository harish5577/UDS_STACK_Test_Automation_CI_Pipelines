from tenma.tenmaDcLib import instantiate_tenma_class_from_device_response
import time

class tenma_robot_lib:
    def __init__(self, port='COM10'):
        self.port = port
        self.psu = None

    def connect_tenma(self, port=None):
        if port:
            self.port = port
        self.psu = instantiate_tenma_class_from_device_response(self.port)
        
    def set_tenma_voltage_and_current(self, voltage_v, current_a, channel=1):
        if not self.psu:
            self.connect_tenma()
        # Convert V to mV and A to mA
        mv = int(float(voltage_v) * 1000)
        ma = int(float(current_a) * 1000)
        self.psu.setVoltage(channel, mv)
        self.psu.setCurrent(channel, ma)
        
    def tenma_power_on(self):
        if not self.psu:
            self.connect_tenma()
        self.psu.ON()
        
    def tenma_power_off(self):
        if not self.psu:
            self.connect_tenma()
        self.psu.OFF()

    def close_tenma(self):
        if self.psu:
            self.psu.close()
            self.psu = None
