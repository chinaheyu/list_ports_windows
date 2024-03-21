# list_ports_windows

This is better solution for enumerating serial devices.

## Features

- Only use the functions provided by the [`Cfgmgr32`](https://learn.microsoft.com/en-us/windows-hardware/drivers/install/porting-from-setupapi-to-cfgmgr32?source=recommendations) api for maximum compatibility;
- Provides more comprehensive information;
- Get the proper usb information of the serial devices;
- Compatible with pyserial.

## Quick Start

### Run as a script

For the first time, we recommend running `list_ports_windows.py` as a script file to see if it produces the results you need.

```
python list_ports_windows.py
```

Example output:

```
[COM9]
Bus 1: Intel(R) USB 3.20 eXtensible Host Controller - 1.20 (Microsoft)
|__ Port 5: VIA Labs, Inc. - USB2.0 Hub
    |__ Port 2: FTDI - FT232R USB UART
Location: 1-5.1
Vendor: 0403
Product: 6001
Manufacturer: FTDI
Product: FT232R USB UART
SerialNumber: B001ADZM
Interface: FT232R USB UART
```

### Enumerates all serial port devices.

```python
from list_ports_windows import iterate_comports

for port_device in iterate_comports():
    print(f'[{port_device.port_name}] {port_device.description} - {port_device.manufacturer}')
```

`port_device` is an instance of the `list_ports_windows.DeviceInterface` object. You can access driver provided device information through its properties.

### Enumerates all serial port devices with usb information.

```python
from list_ports_windows import iterate_comports

# Set the retrieve_usb_info option to True to get the usb information
# corresponding to the serial device.
for port_device, usb_info in iterate_comports(retrieve_usb_info=True):
    if usb_info is None:
        # It could be a virtual serial port or some other device.
        print(f'[{port_device.port_name}] {port_device.description} - {port_device.manufacturer}')
    else:
        # It is a USB serial port device.
        print(f'[{port_device.port_name}] {usb_info.vid:04X}:{usb_info.pid:04X} - {usb_info.product} - {usb_info.manufacturer} - {usb_info.serial_number}')
```

This will get the usb information (pid, vid, product, manufacturer, serial_number, location, interface) while enumerating the serial devices.

`usb_info` is an instance of the `list_ports_windows.USBInfo` object. You can access the usb information through its properties. For some special (non-usb interface) devices, usb_info will be set to None.

Example output:

```
[COM4] 0403:6001 - FT232R USB UART - FTDI - B001ADZM
[COM5] 303A:4001 - Espressif Device - Espressif Systems - 123456
[COM6] 10C4:EA60 - CP2104 USB to UART Bridge Controller - Silicon Labs - 02G7YZB3
[COM7] 0483:5740 - STM32 Virtual ComPort - STMicroelectronics - 2053346D5856
[COM8] com0com - serial port emulator - Vyacheslav Frolov
[COM9] com0com - serial port emulator - Vyacheslav Frolov
```

### Gets information about a specified COM port.

```python
from list_ports_windows import find_device_from_port_name

# Find device by port name.
port_device = find_device_from_port_name('COMx')

# Get the hardware ids of the device.
print(port_device.hardware_ids)

# Get the usb information of the device.
usb_info = port_device.get_usb_info()

# Display usb information.
print(usb_info)
```

### Gets device information

This is an example showing how to get more information about a serial device, such as class name, driver, location, etc.

```pycon
>>> from list_ports_windows import iterate_comports
>>> port_device = next(iterate_comports()) 
>>> port_device.class_name
'Ports'
>>> port_device.driver_inf_path
'usbser.inf'
>>> port_device.location_paths
['PCIROOT(0)#PCI(1400)#USBROOT(0)#USB(5)#USB(1)#USBMI(0)', 'PCIROOT(0)#PCI(1400)#USBROOT(0)#USB(5)#USB(1)#USB(1)', 'ACPI(_SB_)#ACPI(PC00)#ACPI(XHCI)#ACPI(RHUB)#ACPI(HS05)#USB(1)#USBMI(0)', 'ACPI(_SB_)#ACPI(PC00)#ACPI(XHCI)#ACPI(RHUB)#ACPI(HS05)#USB(1)#USB(1)']
>>> port_device.hardware_ids  
['USB\\VID_303A&PID_4001&REV_0100&MI_00', 'USB\\VID_303A&PID_4001&MI_00']
```

### Used in conjunction with pyserial

There are two ways to use it with pyserial. One is to directly replace `serial/tools/list_ports_windows.py` file in the `pyserial` package with the code from `list_ports_windows.py`, and the other is to directly call the `comports()` function as shown below.

```python
from list_ports_windows import comports
from serial.tools.list_ports_common import ListPortInfo

for port in comports():
    print(isinstance(port, ListPortInfo))
```

### Hot Plug (experimental)

```python
from list_ports_windows import PortHotPlugDetector

def arrival_callback(port_name):
    print(port_name)

def removal_callback(port_name):
    print(port_name)

detector = PortHotPlugDetector(arrival_callback, removal_callback)
```
