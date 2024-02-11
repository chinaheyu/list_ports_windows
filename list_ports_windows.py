import ctypes


class GUID(ctypes.Structure):
    _fields_ = [
        ('Data1', ctypes.c_ulong),
        ('Data2', ctypes.c_ushort),
        ('Data3', ctypes.c_ushort),
        ('Data4', ctypes.c_uint8 * 8),
    ]

    def __str__(self):
        return "{{{:08x}-{:04x}-{:04x}-{}-{}}}".format(
            self.Data1,
            self.Data2,
            self.Data3,
            ''.join(["{:02x}".format(d) for d in self.Data4[:2]]),
            ''.join(["{:02x}".format(d) for d in self.Data4[2:]]),
        )


class DEVPROPKEY(ctypes.Structure):
    _fields_ = [
        ('fmtid', GUID),
        ('pid', ctypes.c_ulong)
    ]


class USB_DEVICE_DESCRIPTOR(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('bLength', ctypes.c_uint8),
        ('bDescriptorType', ctypes.c_uint8),
        ('bcdUSB', ctypes.c_uint16),
        ('bDeviceClass', ctypes.c_uint8),
        ('bDeviceSubClass', ctypes.c_uint8),
        ('bDeviceProtocol', ctypes.c_uint8),
        ('bMaxPacketSize0', ctypes.c_uint8),
        ('idVendor', ctypes.c_uint16),
        ('idProduct', ctypes.c_uint16),
        ('bcdDevice', ctypes.c_uint16),
        ('iManufacturer', ctypes.c_uint8),
        ('iProduct', ctypes.c_uint8),
        ('iSerialNumber', ctypes.c_uint8),
        ('bNumConfigurations', ctypes.c_uint8),
    ]


class USB_STRING_DESCRIPTOR(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('bLength', ctypes.c_uint8),
        ('bDescriptorType', ctypes.c_uint8),
        ('bString', ctypes.c_wchar * 255),
    ]


class SetupPacket(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('bmRequest', ctypes.c_uint8),
        ('bRequest', ctypes.c_uint8),
        ('wValue', ctypes.c_uint16),
        ('wIndex', ctypes.c_uint16),
        ('wLength', ctypes.c_uint16),
    ]


class USB_DESCRIPTOR_REQUEST(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('ConnectionIndex', ctypes.c_uint32),
        ('SetupPacket', SetupPacket),
    ]


CreateFileW = ctypes.windll.kernel32.CreateFileW
CreateFileW.argtypes = [
    ctypes.c_void_p,
    ctypes.c_ulong,
    ctypes.c_ulong,
    ctypes.c_void_p,
    ctypes.c_ulong,
    ctypes.c_ulong,
    ctypes.c_void_p
]
CreateFileW.restype = ctypes.c_void_p

CloseHandle = ctypes.windll.kernel32.CloseHandle
CloseHandle.argtypes = [
    ctypes.c_void_p
]
CloseHandle.restype = ctypes.c_int

DeviceIoControl = ctypes.windll.kernel32.DeviceIoControl
DeviceIoControl.argtypes = [
    ctypes.c_void_p,
    ctypes.c_ulong,
    ctypes.c_void_p,
    ctypes.c_ulong,
    ctypes.c_void_p,
    ctypes.c_ulong,
    ctypes.POINTER(ctypes.c_ulong),
    ctypes.c_void_p
]
DeviceIoControl.restype = ctypes.c_int

RegQueryValueExW = ctypes.windll.advapi32.RegQueryValueExW
RegQueryValueExW.argtypes = [
    ctypes.c_void_p,
    ctypes.c_wchar_p,
    ctypes.POINTER(ctypes.c_ulong),
    ctypes.POINTER(ctypes.c_ulong),
    ctypes.c_void_p,
    ctypes.POINTER(ctypes.c_ulong)
]
RegQueryValueExW.restype = ctypes.c_long

RegCloseKey = ctypes.windll.advapi32.RegCloseKey
RegCloseKey.argtypes = [
    ctypes.c_void_p
]
RegCloseKey.restype = ctypes.c_long

CM_Get_Device_Interface_List_SizeW = ctypes.windll.cfgmgr32.CM_Get_Device_Interface_List_SizeW
CM_Get_Device_Interface_List_SizeW.argtypes = [
    ctypes.POINTER(ctypes.c_ulong),
    ctypes.POINTER(GUID),
    ctypes.c_wchar_p,
    ctypes.c_ulong
]
CM_Get_Device_Interface_List_SizeW.restype = ctypes.c_ulong

CM_Get_Device_Interface_ListW = ctypes.windll.cfgmgr32.CM_Get_Device_Interface_ListW
CM_Get_Device_Interface_ListW.argtypes = [
    ctypes.POINTER(GUID),
    ctypes.c_wchar_p,
    ctypes.c_wchar_p,
    ctypes.c_ulong,
    ctypes.c_ulong
]
CM_Get_Device_Interface_ListW.restype = ctypes.c_ulong

CM_Locate_DevNodeW = ctypes.windll.cfgmgr32.CM_Locate_DevNodeW
CM_Locate_DevNodeW.argtypes = [
    ctypes.POINTER(ctypes.c_ulong),
    ctypes.c_wchar_p,
    ctypes.c_ulong
]
CM_Locate_DevNodeW.restype = ctypes.c_ulong

CM_Get_DevNode_PropertyW = ctypes.windll.cfgmgr32.CM_Get_DevNode_PropertyW
CM_Get_DevNode_PropertyW.argtypes = [
    ctypes.c_ulong,
    ctypes.POINTER(DEVPROPKEY),
    ctypes.POINTER(ctypes.c_ulong),
    ctypes.c_char_p,
    ctypes.POINTER(ctypes.c_ulong),
    ctypes.c_ulong
]
CM_Get_DevNode_PropertyW.restype = ctypes.c_ulong

CM_MapCrToWin32Err = ctypes.windll.cfgmgr32.CM_MapCrToWin32Err
CM_MapCrToWin32Err.argtypes = [
    ctypes.c_ulong,
    ctypes.c_ulong
]
CM_MapCrToWin32Err.restype = ctypes.c_ulong

CM_Get_Device_Interface_PropertyW = ctypes.windll.cfgmgr32.CM_Get_Device_Interface_PropertyW
CM_Get_Device_Interface_PropertyW.argtypes = [
    ctypes.c_wchar_p,
    ctypes.POINTER(DEVPROPKEY),
    ctypes.POINTER(ctypes.c_ulong),
    ctypes.c_char_p,
    ctypes.POINTER(ctypes.c_ulong),
    ctypes.c_ulong
]
CM_Get_Device_Interface_PropertyW.restype = ctypes.c_ulong

CM_Get_DevNode_Status = ctypes.windll.cfgmgr32.CM_Get_DevNode_Status
CM_Get_DevNode_Status.argtypes = [
    ctypes.POINTER(ctypes.c_ulong),
    ctypes.POINTER(ctypes.c_ulong),
    ctypes.c_ulong,
    ctypes.c_ulong
]
CM_Get_DevNode_Status.restype = ctypes.c_ulong

CM_Get_Parent = ctypes.windll.cfgmgr32.CM_Get_Parent
CM_Get_Parent.argtypes = [
    ctypes.POINTER(ctypes.c_ulong),
    ctypes.c_ulong,
    ctypes.c_ulong
]
CM_Get_Parent.restype = ctypes.c_ulong

CM_Open_DevNode_Key = ctypes.windll.cfgmgr32.CM_Open_DevNode_Key
CM_Open_DevNode_Key.argtypes = [
    ctypes.c_ulong,
    ctypes.c_ulong,
    ctypes.c_ulong,
    ctypes.c_ulong,
    ctypes.c_void_p,
    ctypes.c_ulong
]
CM_Open_DevNode_Key.restype = ctypes.c_ulong

GUID_DEVINTERFACE_USB_HUB = GUID(0xf18a0e88, 0xc30c, 0x11d0, (0x88, 0x15, 0x00, 0xa0, 0xc9, 0x06, 0xbe, 0xd8))
GUID_DEVINTERFACE_COMPORT = GUID(0X86E0D1E0, 0X8089, 0X11D0, (0X9C, 0XE4, 0X08, 0X00, 0X3E, 0X30, 0X1F, 0X73))

CM_GET_DEVICE_INTERFACE_LIST_PRESENT = 0
CM_LOCATE_DEVNODE_NORMAL = 0
CR_SUCCESS = 0
CR_BUFFER_SMALL = 26
ERROR_MORE_DATA = 234
ERROR_SUCCESS = 0

DEVPKEY_NAME = DEVPROPKEY((0xb725f130, 0x47ef, 0x101a, (0xa5, 0xf1, 0x02, 0x60, 0x8c, 0x9e, 0xeb, 0xac)), 10)
DEVPKEY_Device_DeviceDesc = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 2)
DEVPKEY_Device_HardwareIds = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 3)
DEVPKEY_Device_Class = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 9)
DEVPKEY_Device_Address = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 30)
DEVPKEY_Device_BusReportedDeviceDesc = DEVPROPKEY((0x540b947e, 0x8b40, 0x45bc, (0xa8, 0xa2, 0x6a, 0x0b, 0x89, 0x4c, 0xbd, 0xa2)), 4)
DEVPKEY_Device_FriendlyName = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 14)
DEVPKEY_Device_LocationPaths = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 37)
DEVPKEY_Device_InstanceId = DEVPROPKEY((0x78c34fc8, 0x104a, 0x4aca, (0x9e, 0xa4, 0x52, 0x4d, 0x52, 0x99, 0x6e, 0x57)), 256)

DEVPROP_TYPEMOD_LIST = 0x00002000
DEVPROP_TYPE_UINT32 = 0x00000007
DEVPROP_TYPE_STRING = 0x00000012
DEVPROP_TYPE_STRING_LIST = (DEVPROP_TYPE_STRING | DEVPROP_TYPEMOD_LIST)

KEY_READ = 0x20019
RegDisposition_OpenExisting = 1
CM_REGISTRY_HARDWARE = 0

USB_REQUEST_GET_DESCRIPTOR = 6
USB_DEVICE_DESCRIPTOR_TYPE = 1
USB_STRING_DESCRIPTOR_TYPE = 3
IOCTL_USB_GET_DESCRIPTOR_FROM_NODE_CONNECTION = 2229264

GENERIC_WRITE = 1073741824
OPEN_EXISTING = 3
INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value


def parse_device_property(buffer, buffer_size, property_type):
    if property_type.value == DEVPROP_TYPE_STRING:
        return ctypes.wstring_at(buffer, buffer_size.value // 2 - 1)
    if property_type.value == DEVPROP_TYPE_UINT32:
        return ctypes.cast(buffer, ctypes.POINTER(ctypes.c_uint32)).contents.value
    if property_type.value == DEVPROP_TYPE_STRING_LIST:
        return ctypes.wstring_at(buffer, buffer_size.value // 2).strip('\0').split('\0')
    raise NotImplementedError(f'DEVPROPTYPE {property_type.value} is not implemented!')


def enumerate_device_interface(guid):
    device_interface_list_size = ctypes.c_ulong()
    cr = CM_Get_Device_Interface_List_SizeW(
        ctypes.byref(device_interface_list_size),
        ctypes.byref(guid),
        None,
        CM_GET_DEVICE_INTERFACE_LIST_PRESENT
    )
    if cr != CR_SUCCESS:
        raise ctypes.WinError(CM_MapCrToWin32Err(cr, 0))
    if device_interface_list_size.value > 1:
        device_interface_list = ctypes.create_unicode_buffer(device_interface_list_size.value)
        cr = CM_Get_Device_Interface_ListW(
            ctypes.byref(guid),
            None,
            device_interface_list,
            ctypes.sizeof(device_interface_list),
            CM_GET_DEVICE_INTERFACE_LIST_PRESENT
        )
        if cr != CR_SUCCESS:
            raise ctypes.WinError(CM_MapCrToWin32Err(cr, 0))
        for interface in ctypes.wstring_at(device_interface_list, device_interface_list_size.value).strip('\0').split('\0'):
            yield DeviceInterface(interface)


def request_usb_string_description(h_hub_device, usb_hub_port, idx):
    if idx == 0:
        # There is no string description at this index.
        return None

    # Filling setup package.
    description_request_buffer = ctypes.create_string_buffer(ctypes.sizeof(USB_DESCRIPTOR_REQUEST) + ctypes.sizeof(USB_STRING_DESCRIPTOR))
    description_request = ctypes.cast(description_request_buffer, ctypes.POINTER(USB_DESCRIPTOR_REQUEST))
    description_request.contents.ConnectionIndex = usb_hub_port
    description_request.contents.SetupPacket.wValue = (USB_STRING_DESCRIPTOR_TYPE << 8) | idx
    description_request.contents.SetupPacket.wIndex = 0
    description_request.contents.SetupPacket.wLength = ctypes.sizeof(description_request_buffer) - ctypes.sizeof(USB_DESCRIPTOR_REQUEST)

    # Send string description request.
    if not DeviceIoControl(
            h_hub_device,
            IOCTL_USB_GET_DESCRIPTOR_FROM_NODE_CONNECTION,
            description_request_buffer,
            ctypes.sizeof(description_request_buffer),
            description_request_buffer,
            ctypes.sizeof(description_request_buffer),
            None,
            None
    ):
        return None

    # Parse string description from wstring buffer.
    description = ctypes.cast(ctypes.byref(description_request_buffer, ctypes.sizeof(USB_DESCRIPTOR_REQUEST)),
                              ctypes.POINTER(USB_STRING_DESCRIPTOR))
    return ctypes.wstring_at(description.contents.bString, description.contents.bLength // 2 - 1)


def request_usb_device_description(h_hub_device, usb_hub_port):
    # Filling setup package for requesting device description.
    device_description_request_buffer = ctypes.create_string_buffer(
        ctypes.sizeof(USB_DESCRIPTOR_REQUEST) + ctypes.sizeof(USB_DEVICE_DESCRIPTOR)
    )
    device_description_request = ctypes.cast(device_description_request_buffer, ctypes.POINTER(USB_DESCRIPTOR_REQUEST))
    device_description_request.contents.ConnectionIndex = usb_hub_port
    device_description_request.contents.SetupPacket.bmRequest = 0x80
    device_description_request.contents.SetupPacket.bRequest = USB_REQUEST_GET_DESCRIPTOR
    device_description_request.contents.SetupPacket.wValue = USB_DEVICE_DESCRIPTOR_TYPE << 8
    device_description_request.contents.SetupPacket.wLength = ctypes.sizeof(USB_DEVICE_DESCRIPTOR)

    # Send usb device description request.
    if not DeviceIoControl(
            h_hub_device,
            IOCTL_USB_GET_DESCRIPTOR_FROM_NODE_CONNECTION,
            device_description_request_buffer,
            ctypes.sizeof(device_description_request_buffer),
            device_description_request_buffer,
            ctypes.sizeof(device_description_request_buffer),
            None,
            None
    ):
        return None

    # Get device description.
    device_description = USB_DEVICE_DESCRIPTOR()
    ptr = ctypes.cast(
        ctypes.byref(device_description_request_buffer, ctypes.sizeof(USB_DESCRIPTOR_REQUEST)),
        ctypes.POINTER(USB_DEVICE_DESCRIPTOR)
    )
    ctypes.memmove(ctypes.addressof(device_description), ptr, ctypes.sizeof(USB_DEVICE_DESCRIPTOR))
    return device_description


class DeviceNode:
    def __init__(self, instance_handle):
        self.__instance_handle = instance_handle

    @property
    def status(self):
        node_status = ctypes.c_ulong()
        problem_number = ctypes.c_ulong()
        cr = CM_Get_DevNode_Status(ctypes.byref(node_status), ctypes.byref(problem_number), self.__instance_handle, 0)
        if cr != CR_SUCCESS:
            return None
        return node_status, problem_number

    @property
    def parent(self):
        parent_instance_number = ctypes.c_ulong()
        cr = CM_Get_Parent(ctypes.byref(parent_instance_number), self.__instance_handle, 0)
        if cr != CR_SUCCESS:
            return None
        return DeviceNode(parent_instance_number.value)

    @property
    def name(self):
        return self.get_property(DEVPKEY_NAME)

    @property
    def description(self):
        return self.get_property(DEVPKEY_Device_DeviceDesc)

    @property
    def hardware_ids(self):
        return self.get_property(DEVPKEY_Device_HardwareIds)

    @property
    def class_name(self):
        return self.get_property(DEVPKEY_Device_Class)

    @property
    def address(self):
        return self.get_property(DEVPKEY_Device_Address)

    @property
    def bus_reported_device_description(self):
        return self.get_property(DEVPKEY_Device_BusReportedDeviceDesc)

    @property
    def friendly_name(self):
        return self.get_property(DEVPKEY_Device_FriendlyName)

    @property
    def location_paths(self):
        return self.get_property(DEVPKEY_Device_LocationPaths)

    @property
    def port_name(self):
        hkey = ctypes.c_void_p()
        cr = CM_Open_DevNode_Key(
            self.__instance_handle,
            KEY_READ,
            0,
            RegDisposition_OpenExisting,
            ctypes.byref(hkey),
            CM_REGISTRY_HARDWARE
        )
        if cr != CR_SUCCESS:
            return None
        port_name_buffer_size = ctypes.c_ulong()
        cr = RegQueryValueExW(
            hkey,
            "PortName",
            None,
            None,
            None,
            ctypes.byref(port_name_buffer_size)
        )
        if (cr != ERROR_SUCCESS) and (cr != ERROR_MORE_DATA):
            return None
        port_name_buffer = ctypes.create_unicode_buffer(port_name_buffer_size.value // 2)
        cr = RegQueryValueExW(
            hkey,
            "PortName",
            None,
            None,
            port_name_buffer,
            ctypes.byref(port_name_buffer_size)
        )
        if cr != ERROR_SUCCESS:
            return None
        RegCloseKey(hkey)
        return port_name_buffer.value

    def get_property(self, property_key):
        buffer_size = ctypes.c_ulong()
        property_type = ctypes.c_ulong()
        cr = CM_Get_DevNode_PropertyW(
            self.__instance_handle,
            ctypes.byref(property_key),
            ctypes.byref(property_type),
            None,
            ctypes.byref(buffer_size),
            0
        )
        if cr != CR_BUFFER_SMALL and cr != CR_SUCCESS:
            return None
        buffer = ctypes.create_string_buffer(buffer_size.value)
        cr = CM_Get_DevNode_PropertyW(
                self.__instance_handle,
                ctypes.byref(property_key),
                ctypes.byref(property_type),
                buffer,
                ctypes.byref(buffer_size),
                0
        )
        if cr != CR_SUCCESS:
            return None
        return parse_device_property(buffer, buffer_size, property_type)

    @property
    def instance_identifier(self):
        return self.get_property(DEVPKEY_Device_InstanceId)

    @property
    def instance_handle(self):
        return self.__instance_handle


class DeviceInterface(DeviceNode):
    def __init__(self, interface):
        self.__interface = interface
        super().__init__(self.instance_handle)

    @property
    def interface(self):
        return self.__interface

    @property
    def instance_identifier(self):
        return self.get_interface_property(DEVPKEY_Device_InstanceId)

    @property
    def instance_handle(self):
        instance_handle = ctypes.c_ulong()
        cr = CM_Locate_DevNodeW(
            ctypes.byref(instance_handle),
            self.instance_identifier,
            CM_LOCATE_DEVNODE_NORMAL
        )
        if cr != CR_SUCCESS:
            return None
        return instance_handle.value

    def get_interface_property(self, property_key):
        buffer_size = ctypes.c_ulong()
        property_type = ctypes.c_ulong()
        cr = CM_Get_Device_Interface_PropertyW(
            self.__interface,
            ctypes.byref(property_key),
            ctypes.byref(property_type),
            None,
            ctypes.byref(buffer_size),
            0
        )
        if cr != CR_BUFFER_SMALL and cr != CR_SUCCESS:
            return None
        buffer = ctypes.create_string_buffer(buffer_size.value)
        cr = CM_Get_Device_Interface_PropertyW(
                self.__interface,
                ctypes.byref(property_key),
                ctypes.byref(property_type),
                buffer,
                ctypes.byref(buffer_size),
                0
        )
        if cr != CR_SUCCESS:
            return None
        return parse_device_property(buffer, buffer_size, property_type)


class USBInfo:
    def __init__(self, pid, vid, product, manufacturer, serial_number):
        self.pid = pid
        self.vid = vid
        self.product = product
        self.manufacturer = manufacturer
        self.serial_number = serial_number

    def __str__(self):
        return f'{self.pid:04X}:{self.vid:04X} - {self.product} - {self.manufacturer} - {self.serial_number}'


def get_usb_info(hub_device, usb_device):
    hub_path = "\\\\?\\" + hub_device.instance_identifier.replace("\\", "#") + f"#{GUID_DEVINTERFACE_USB_HUB}"
    h_hub_device = CreateFileW(
        hub_path,
        GENERIC_WRITE,
        2,
        None,
        OPEN_EXISTING,
        0,
        0
    )
    if h_hub_device == INVALID_HANDLE_VALUE:
        return None

    usb_hub_port = usb_device.address
    if usb_hub_port is None:
        return None

    device_description = request_usb_device_description(h_hub_device, usb_hub_port)
    if device_description is None:
        return None

    # Get product identifier and vendor identifier.
    pid = device_description.idProduct
    vid = device_description.idVendor

    # Request string description at iProduct, iManufacturer and iSerialNumber.
    product = request_usb_string_description(h_hub_device, usb_hub_port, device_description.iProduct)
    manufacturer = request_usb_string_description(h_hub_device, usb_hub_port, device_description.iManufacturer)
    serial_number = request_usb_string_description(h_hub_device, usb_hub_port, device_description.iSerialNumber)

    CloseHandle(h_hub_device)
    return USBInfo(pid, vid, product, manufacturer, serial_number)


def find_parent_hub_and_usb(device_node, hub_instance_handles):
    child_device = device_node
    while True:
        parent_device = child_device.parent
        if parent_device is None:
            return None
        if parent_device.instance_handle in hub_instance_handles:
            return parent_device, child_device
        child_device = parent_device


def iterate_comports():
    hub_instance_handles = [i.instance_handle for i in enumerate_device_interface(GUID_DEVINTERFACE_USB_HUB)]
    for port_device in enumerate_device_interface(GUID_DEVINTERFACE_COMPORT):
        hub_and_usb = find_parent_hub_and_usb(port_device, hub_instance_handles)
        if hub_and_usb is None:
            yield port_device, None
        else:
            usb_info = get_usb_info(*hub_and_usb)
            if usb_info is None:
                yield port_device, None
            else:
                yield port_device, usb_info


def main():
    for port_device, usb_info in iterate_comports():
        print(f'[{port_device.port_name}] {usb_info}')


if __name__ == '__main__':
    main()
