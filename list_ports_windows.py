#! python
# https://github.com/chinaheyu/list_ports_windows
# Enumerate serial ports on Windows including a human-readable description and hardware information.

import ctypes
import re


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


class USB_COMMON_DESCRIPTOR(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('bLength', ctypes.c_uint8),
        ('bDescriptorType', ctypes.c_uint8),
    ]


class USB_CONFIGURATION_DESCRIPTOR(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('bLength', ctypes.c_uint8),
        ('bDescriptorType', ctypes.c_uint8),
        ('wTotalLength', ctypes.c_uint16),
        ('bNumInterfaces', ctypes.c_uint8),
        ('bConfigurationValue', ctypes.c_uint8),
        ('iConfiguration', ctypes.c_uint8),
        ('bmAttributes', ctypes.c_uint8),
        ('MaxPower', ctypes.c_uint8),
    ]


class USB_INTERFACE_DESCRIPTOR(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('bLength', ctypes.c_uint8),
        ('bDescriptorType', ctypes.c_uint8),
        ('bInterfaceNumber', ctypes.c_uint8),
        ('bAlternateSetting', ctypes.c_uint8),
        ('bNumEndpoints', ctypes.c_uint8),
        ('bInterfaceClass', ctypes.c_uint8),
        ('bInterfaceSubClass', ctypes.c_uint8),
        ('bInterfaceProtocol', ctypes.c_uint8),
        ('iInterface', ctypes.c_uint8),
    ]


class USB_INTERFACE_ASSOCIATION_DESCRIPTOR(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('bLength', ctypes.c_uint8),
        ('bDescriptorType', ctypes.c_uint8),
        ('bFirstInterface', ctypes.c_uint8),
        ('bInterfaceCount', ctypes.c_uint8),
        ('bFunctionClass', ctypes.c_uint8),
        ('bFunctionSubClass', ctypes.c_uint8),
        ('bFunctionProtocol', ctypes.c_uint8),
        ('iFunction', ctypes.c_uint8)
    ]


class USB_NODE_CONNECTION_INFORMATION_EX(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('ConnectionIndex', ctypes.c_ulong),
        ('DeviceDescriptor', USB_DEVICE_DESCRIPTOR),
        ('CurrentConfigurationValue', ctypes.c_uint8),
        ('Speed', ctypes.c_uint8),
        ('DeviceIsHub', ctypes.c_uint8),
        ('DeviceAddress', ctypes.c_uint16),
        ('NumberOfOpenPipes', ctypes.c_ulong),
        ('ConnectionStatus', ctypes.c_uint),
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


class USB_ROOT_HUB_NAME(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('ActualLength', ctypes.c_ulong),
        ('RootHubName', ctypes.c_wchar),
    ]


class USB_NODE_CONNECTION_NAME(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ('ConnectionIndex', ctypes.c_ulong),
        ('ActualLength', ctypes.c_ulong),
        ('NodeName', ctypes.c_wchar),
    ]


class CM_POWER_DATA(ctypes.Structure):
    _fields_ = [
        ('PD_Size', ctypes.c_ulong),
        ('PD_MostRecentPowerState', ctypes.c_int),
        ('PD_Capabilities', ctypes.c_ulong),
        ('PD_D1Latency', ctypes.c_ulong),
        ('PD_D2Latency', ctypes.c_ulong),
        ('PD_D3Latency', ctypes.c_ulong),
        ('PD_PowerStateMapping', ctypes.c_int * 7),
        ('PD_DeepestSystemWake', ctypes.c_int),
    ]


class USBUSER_REQUEST_HEADER(ctypes.Structure):
    _fields_ = [
        ('UsbUserRequest', ctypes.c_ulong),
        ('UsbUserStatusCode', ctypes.c_int),
        ('RequestBufferLength', ctypes.c_ulong),
        ('ActualBufferLength', ctypes.c_ulong)
    ]


class USB_CONTROLLER_INFO_0(ctypes.Structure):
    _fields_ = [
        ('PciVendorId', ctypes.c_ulong),
        ('PciDeviceId', ctypes.c_ulong),
        ('PciRevision', ctypes.c_ulong),
        ('NumberOfRootPorts', ctypes.c_ulong),
        ('ControllerFlavor', ctypes.c_int),
        ('HcFeatureFlags', ctypes.c_ulong)
    ]


class USBUSER_CONTROLLER_INFO_0(ctypes.Structure):
    _fields_ = [
        ('Header', USBUSER_REQUEST_HEADER),
        ('Info0', USB_CONTROLLER_INFO_0)
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

GetUserDefaultLangID = ctypes.windll.kernel32.GetUserDefaultLangID
GetUserDefaultLangID.restype = ctypes.c_ushort

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

CM_Get_Child = ctypes.windll.cfgmgr32.CM_Get_Child
CM_Get_Child.argtypes = [
    ctypes.POINTER(ctypes.c_ulong),
    ctypes.c_ulong,
    ctypes.c_ulong
]
CM_Get_Child.restype = ctypes.c_ulong

CM_Get_Sibling = ctypes.windll.cfgmgr32.CM_Get_Sibling
CM_Get_Sibling.argtypes = [
    ctypes.POINTER(ctypes.c_ulong),
    ctypes.c_ulong,
    ctypes.c_ulong
]
CM_Get_Sibling.restype = ctypes.c_ulong

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

CM_Enable_DevNode = ctypes.windll.cfgmgr32.CM_Enable_DevNode
CM_Enable_DevNode.argtypes = [
    ctypes.c_ulong,
    ctypes.c_ulong
]
CM_Enable_DevNode.restype = ctypes.c_ulong

CM_Disable_DevNode = ctypes.windll.cfgmgr32.CM_Disable_DevNode
CM_Disable_DevNode.argtypes = [
    ctypes.c_ulong,
    ctypes.c_ulong
]
CM_Disable_DevNode.restype = ctypes.c_ulong

GUID_DEVINTERFACE_USB_HUB = GUID(0xf18a0e88, 0xc30c, 0x11d0, (0x88, 0x15, 0x00, 0xa0, 0xc9, 0x06, 0xbe, 0xd8))
GUID_DEVINTERFACE_COMPORT = GUID(0X86E0D1E0, 0X8089, 0X11D0, (0X9C, 0XE4, 0X08, 0X00, 0X3E, 0X30, 0X1F, 0X73))
GUID_DEVINTERFACE_MODEM = GUID(0x2c7089aa, 0x2e0e, 0x11d1, (0xb1, 0x14, 0x00, 0xc0, 0x4f, 0xc2, 0xaa, 0xe4))
GUID_DEVINTERFACE_USB_HOST_CONTROLLER = GUID(0x3abf6f2d, 0x71c4, 0x462a, (0x8a, 0x92, 0x1e, 0x68, 0x61, 0xe6, 0xaf, 0x27))

CM_GET_DEVICE_INTERFACE_LIST_PRESENT = 0
CM_LOCATE_DEVNODE_NORMAL = 0
CR_SUCCESS = 0
CR_BUFFER_SMALL = 26
ERROR_MORE_DATA = 234
ERROR_SUCCESS = 0

DEVPKEY_NAME = DEVPROPKEY((0xb725f130, 0x47ef, 0x101a, (0xa5, 0xf1, 0x02, 0x60, 0x8c, 0x9e, 0xeb, 0xac)), 10)
DEVPKEY_Device_DeviceDesc = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 2)
DEVPKEY_Device_HardwareIds = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 3)
DEVPKEY_Device_CompatibleIds = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 4)
DEVPKEY_Device_Service = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 6)
DEVPKEY_Device_Class = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 9)
DEVPKEY_Device_ClassGuid = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 10)
DEVPKEY_Device_Driver = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 11)
DEVPKEY_Device_Manufacturer = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 13)
DEVPKEY_Device_LocationInfo = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 15)
DEVPKEY_Device_FriendlyName = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 14)
DEVPKEY_Device_PDOName = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 16)
DEVPKEY_Device_BusNumber = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 23)
DEVPKEY_Device_EnumeratorName = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 24)
DEVPKEY_Device_PowerData = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 32)
DEVPKEY_Device_LocationPaths = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 37)
DEVPKEY_Device_Address = DEVPROPKEY((0xa45c254e, 0xdf1c, 0x4efd, (0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0)), 30)
DEVPKEY_Device_InstanceId = DEVPROPKEY((0x78c34fc8, 0x104a, 0x4aca, (0x9e, 0xa4, 0x52, 0x4d, 0x52, 0x99, 0x6e, 0x57)), 256)
DEVPKEY_Device_BusReportedDeviceDesc = DEVPROPKEY((0x540b947e, 0x8b40, 0x45bc, (0xa8, 0xa2, 0x6a, 0x0b, 0x89, 0x4c, 0xbd, 0xa2)), 4)
DEVPKEY_Device_DriverInfPath = DEVPROPKEY((0xa8b865dd, 0x2e3d, 0x4094, (0xad, 0x97, 0xe5, 0x93, 0xa7, 0xc, 0x75, 0xd6)), 5)

DEVPROP_TYPEMOD_ARRAY = 0x00001000
DEVPROP_TYPEMOD_LIST = 0x00002000
DEVPROP_TYPE_UINT32 = 0x00000007
DEVPROP_TYPE_STRING = 0x00000012
DEVPROP_TYPE_GUID = 0x0000000D
DEVPROP_TYPE_BYTE = 0x00000003
DEVPROP_TYPE_BINARY = (DEVPROP_TYPE_BYTE | DEVPROP_TYPEMOD_ARRAY)
DEVPROP_TYPE_STRING_LIST = (DEVPROP_TYPE_STRING | DEVPROP_TYPEMOD_LIST)

KEY_READ = 0x20019
RegDisposition_OpenExisting = 1
CM_REGISTRY_HARDWARE = 0

USB_REQUEST_GET_DESCRIPTOR = 6
USB_DEVICE_DESCRIPTOR_TYPE = 1
USB_CONFIGURATION_DESCRIPTOR_TYPE = 2
USB_STRING_DESCRIPTOR_TYPE = 3
USB_INTERFACE_DESCRIPTOR_TYPE = 4
USB_INTERFACE_ASSOCIATION_DESCRIPTOR_TYPE = 0x0B
IOCTL_USB_GET_DESCRIPTOR_FROM_NODE_CONNECTION = 2229264
IOCTL_USB_GET_NODE_CONNECTION_INFORMATION_EX = 2229320
IOCTL_USB_GET_ROOT_HUB_NAME = 2229256
IOCTL_USB_GET_NODE_CONNECTION_NAME = 2229268
IOCTL_USB_USER_REQUEST = 2229304
USBUSER_GET_CONTROLLER_INFO_0 = 0x00000001

GENERIC_READ = 2147483648
GENERIC_WRITE = 1073741824
OPEN_EXISTING = 3
FILE_ATTRIBUTE_NORMAL = 128
INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value


def parse_device_property(buffer, buffer_size, property_type):
    if property_type.value == DEVPROP_TYPE_STRING:
        return ctypes.wstring_at(buffer, buffer_size.value // 2 - 1)
    if property_type.value == DEVPROP_TYPE_UINT32:
        return ctypes.cast(buffer, ctypes.POINTER(ctypes.c_uint32)).contents.value
    if property_type.value == DEVPROP_TYPE_STRING_LIST:
        return ctypes.wstring_at(buffer, buffer_size.value // 2).strip('\0').split('\0')
    if property_type.value == DEVPROP_TYPE_GUID:
        return GUID.from_buffer_copy(buffer)
    if property_type.value == DEVPROP_TYPE_BINARY:
        return buffer
    raise NotImplementedError(f'DEVPROPTYPE {property_type.value} is not implemented!')


def request_supported_languages(h_hub_device, usb_hub_port):
    # Filling setup package.
    description_request_buffer = ctypes.create_string_buffer(
        ctypes.sizeof(USB_DESCRIPTOR_REQUEST) + ctypes.sizeof(USB_STRING_DESCRIPTOR)
    )
    description_request = ctypes.cast(description_request_buffer, ctypes.POINTER(USB_DESCRIPTOR_REQUEST))
    description_request.contents.ConnectionIndex = usb_hub_port
    description_request.contents.SetupPacket.wValue = USB_STRING_DESCRIPTOR_TYPE << 8
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
    description = ctypes.cast(
        ctypes.byref(description_request_buffer, ctypes.sizeof(USB_DESCRIPTOR_REQUEST)),
        ctypes.POINTER(USB_STRING_DESCRIPTOR)
    )
    languages = []
    for i in range(2, description.contents.bLength, 2):
        languages.append(ctypes.c_uint16.from_buffer(
            description_request_buffer,
            ctypes.sizeof(USB_DESCRIPTOR_REQUEST) + i
        ).value)
    return languages


def request_and_select_language_id(h_hub_device, usb_hub_port):
    available_languages = request_supported_languages(h_hub_device, usb_hub_port)
    if available_languages is None:
        return 0x0409
    default_language = GetUserDefaultLangID()
    if default_language in available_languages:
        return default_language
    else:
        return available_languages[0]


def request_usb_string_description(h_hub_device, usb_hub_port, idx, langid=None):
    if idx == 0:
        # There is no string description at this index.
        return None

    # Get a list of supported languages from index 0, and match langid to system default language.
    if langid is None:
        langid = request_and_select_language_id(h_hub_device, usb_hub_port)

    # Filling setup package.
    description_request_buffer = ctypes.create_string_buffer(ctypes.sizeof(USB_DESCRIPTOR_REQUEST) + ctypes.sizeof(USB_STRING_DESCRIPTOR))
    description_request = ctypes.cast(description_request_buffer, ctypes.POINTER(USB_DESCRIPTOR_REQUEST))
    description_request.contents.ConnectionIndex = usb_hub_port
    description_request.contents.SetupPacket.wValue = (USB_STRING_DESCRIPTOR_TYPE << 8) | idx
    description_request.contents.SetupPacket.wIndex = langid
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
    return USB_DEVICE_DESCRIPTOR.from_buffer_copy(device_description_request_buffer, ctypes.sizeof(USB_DESCRIPTOR_REQUEST))


def request_usb_configuration_description(h_hub_device, usb_hub_port, bConfigurationValue):
    # Filling setup package for requesting configuration description.
    configuration_description_request_buffer = ctypes.create_string_buffer(
        ctypes.sizeof(USB_DESCRIPTOR_REQUEST) + ctypes.sizeof(USB_CONFIGURATION_DESCRIPTOR)
    )
    configuration_description_request = ctypes.cast(
        configuration_description_request_buffer,
        ctypes.POINTER(USB_DESCRIPTOR_REQUEST)
    )
    configuration_description_request.contents.ConnectionIndex = usb_hub_port
    configuration_description_request.contents.SetupPacket.bmRequest = 0x80
    configuration_description_request.contents.SetupPacket.bRequest = USB_REQUEST_GET_DESCRIPTOR
    configuration_description_request.contents.SetupPacket.wValue = (USB_CONFIGURATION_DESCRIPTOR_TYPE << 8) | (bConfigurationValue - 1)
    configuration_description_request.contents.SetupPacket.wLength = ctypes.sizeof(USB_CONFIGURATION_DESCRIPTOR)

    # Send usb configuration description request.
    if not DeviceIoControl(
            h_hub_device,
            IOCTL_USB_GET_DESCRIPTOR_FROM_NODE_CONNECTION,
            configuration_description_request_buffer,
            ctypes.sizeof(configuration_description_request_buffer),
            configuration_description_request_buffer,
            ctypes.sizeof(configuration_description_request_buffer),
            None,
            None
    ):
        return None

    # Get configuration description.
    return USB_CONFIGURATION_DESCRIPTOR.from_buffer_copy(configuration_description_request_buffer, ctypes.sizeof(USB_DESCRIPTOR_REQUEST))


def request_usb_interface_descriptions(h_hub_device, usb_hub_port, configuration_description, bInterfaceNumber, bAlternateSetting=0):
    # Single interface usb device.
    if bInterfaceNumber is None:
        bInterfaceNumber = 0

    # Now request the entire Configuration Descriptor using a dynamically
    # allocated buffer which is sized big enough to hold the entire descriptor
    configuration_description_request_buffer = ctypes.create_string_buffer(
        ctypes.sizeof(USB_DESCRIPTOR_REQUEST) + configuration_description.wTotalLength
    )
    configuration_description_request = ctypes.cast(
        configuration_description_request_buffer,
        ctypes.POINTER(USB_DESCRIPTOR_REQUEST)
    )
    configuration_description_request.contents.ConnectionIndex = usb_hub_port
    configuration_description_request.contents.SetupPacket.bmRequest = 0x80
    configuration_description_request.contents.SetupPacket.bRequest = USB_REQUEST_GET_DESCRIPTOR
    configuration_description_request.contents.SetupPacket.wValue = (USB_CONFIGURATION_DESCRIPTOR_TYPE << 8) | (configuration_description.bConfigurationValue - 1)
    configuration_description_request.contents.SetupPacket.wLength = configuration_description.wTotalLength

    # Send usb configuration description request.
    if not DeviceIoControl(
            h_hub_device,
            IOCTL_USB_GET_DESCRIPTOR_FROM_NODE_CONNECTION,
            configuration_description_request_buffer,
            ctypes.sizeof(configuration_description_request_buffer),
            configuration_description_request_buffer,
            ctypes.sizeof(configuration_description_request_buffer),
            None,
            None
    ):
        return None

    # https://github.com/microsoft/Windows-driver-samples/blob/main/usb/usbview/enum.c#L2413
    common_description_offset = 0
    target_interface_association_description = None
    target_interface_description = None
    while True:
        if common_description_offset + ctypes.sizeof(USB_COMMON_DESCRIPTOR) < configuration_description.wTotalLength:
            common_description = ctypes.cast(
                ctypes.byref(configuration_description_request_buffer, ctypes.sizeof(USB_DESCRIPTOR_REQUEST) + common_description_offset),
                ctypes.POINTER(USB_COMMON_DESCRIPTOR)
            )
            if common_description_offset + common_description.contents.bLength <= configuration_description.wTotalLength:
                if common_description.contents.bDescriptorType == USB_INTERFACE_ASSOCIATION_DESCRIPTOR_TYPE:
                    if common_description.contents.bLength == ctypes.sizeof(USB_INTERFACE_ASSOCIATION_DESCRIPTOR):
                        interface_association_description = ctypes.cast(
                            ctypes.byref(configuration_description_request_buffer, ctypes.sizeof(USB_DESCRIPTOR_REQUEST) + common_description_offset),
                            ctypes.POINTER(USB_INTERFACE_ASSOCIATION_DESCRIPTOR)
                        )
                        if interface_association_description.contents.bFirstInterface <= bInterfaceNumber <= interface_association_description.contents.bFirstInterface + interface_association_description.contents.bInterfaceCount - 1:
                            target_interface_association_description = USB_INTERFACE_ASSOCIATION_DESCRIPTOR.from_buffer_copy(
                                configuration_description_request_buffer,
                                ctypes.sizeof(USB_DESCRIPTOR_REQUEST) + common_description_offset
                            )
                    else:
                        break

                if common_description.contents.bDescriptorType == USB_INTERFACE_DESCRIPTOR_TYPE:
                    if common_description.contents.bLength == ctypes.sizeof(USB_INTERFACE_DESCRIPTOR):
                        interface_description = ctypes.cast(
                            ctypes.byref(configuration_description_request_buffer, ctypes.sizeof(USB_DESCRIPTOR_REQUEST) + common_description_offset),
                            ctypes.POINTER(USB_INTERFACE_DESCRIPTOR)
                        )
                        if interface_description.contents.bInterfaceNumber == bInterfaceNumber and interface_description.contents.bAlternateSetting == bAlternateSetting:
                            target_interface_description = USB_INTERFACE_DESCRIPTOR.from_buffer_copy(
                                configuration_description_request_buffer,
                                ctypes.sizeof(USB_DESCRIPTOR_REQUEST) + common_description_offset
                            )
                    else:
                        break

                # interface association description precedes interface description
                # https://www.usb.org/sites/default/files/iadclasscode_r10.pdf figure 2-1
                if target_interface_description is not None:
                    break

                common_description_offset += common_description.contents.bLength
            else:
                break
        else:
            break
    return target_interface_association_description, target_interface_description


def request_usb_connection_info(h_hub_device, usb_hub_port):
    connection_info = USB_NODE_CONNECTION_INFORMATION_EX()
    connection_info.ConnectionIndex = usb_hub_port
    if not DeviceIoControl(
            h_hub_device,
            IOCTL_USB_GET_NODE_CONNECTION_INFORMATION_EX,
            ctypes.byref(connection_info),
            ctypes.sizeof(USB_NODE_CONNECTION_INFORMATION_EX),
            ctypes.byref(connection_info),
            ctypes.sizeof(USB_NODE_CONNECTION_INFORMATION_EX),
            None,
            None
    ):
        return None
    return connection_info


class cached_property:
    # Decorator (non-data) for building an attribute on-demand on first use.
    def __init__(self, function):
        self.__attr_name = function.__name__
        self.__function = function

    def __get__(self, instance, owner):
        attr = self.__function(instance)
        setattr(instance, self.__attr_name, attr)
        return attr


class DeviceNode:
    def __init__(self, instance_handle):
        self.__instance_handle = instance_handle

    def __str__(self):
        return self.name

    def __eq__(self, other):
        return self.__instance_handle == other.__instance_handle

    def __lt__(self, other):
        return self.__instance_handle < other.__instance_handle

    def __hash__(self):
        return hash(self.__instance_handle)

    def enable(self):
        return CM_Enable_DevNode(self.__instance_handle, 0) == CR_SUCCESS

    def disable(self):
        return CM_Disable_DevNode(self.__instance_handle, 0) == CR_SUCCESS

    @property
    def instance_handle(self):
        return self.__instance_handle

    @cached_property
    def instance_identifier(self):
        return self.get_property(DEVPKEY_Device_InstanceId)

    @property
    def status(self):
        node_status = ctypes.c_ulong()
        problem_number = ctypes.c_ulong()
        cr = CM_Get_DevNode_Status(ctypes.byref(node_status), ctypes.byref(problem_number), self.__instance_handle, 0)
        if cr != CR_SUCCESS:
            return None
        return node_status, problem_number

    @cached_property
    def parent(self):
        parent_instance_handle = ctypes.c_ulong()
        cr = CM_Get_Parent(ctypes.byref(parent_instance_handle), self.__instance_handle, 0)
        if cr != CR_SUCCESS:
            return None
        return DeviceNode(parent_instance_handle.value)

    @cached_property
    def first_child(self):
        child_instance_handle = ctypes.c_ulong()
        cr = CM_Get_Child(ctypes.byref(child_instance_handle), self.__instance_handle, 0)
        if cr != CR_SUCCESS:
            return None
        return DeviceNode(child_instance_handle.value)

    @cached_property
    def sibling(self):
        sibling_instance_handle = ctypes.c_ulong()
        cr = CM_Get_Sibling(ctypes.byref(sibling_instance_handle), self.__instance_handle, 0)
        if cr != CR_SUCCESS:
            return None
        return DeviceNode(sibling_instance_handle.value)

    @cached_property
    def children(self):
        children = []
        child = self.first_child
        while child is not None:
            children.append(DeviceNode(child.instance_handle))
            child = child.sibling
        return children

    @cached_property
    def name(self):
        return self.get_property(DEVPKEY_NAME)

    @cached_property
    def description(self):
        return self.get_property(DEVPKEY_Device_DeviceDesc)

    @cached_property
    def hardware_ids(self):
        return self.get_property(DEVPKEY_Device_HardwareIds)

    @cached_property
    def class_name(self):
        return self.get_property(DEVPKEY_Device_Class)

    @cached_property
    def class_guid(self):
        return self.get_property(DEVPKEY_Device_ClassGuid)

    @cached_property
    def driver(self):
        return self.get_property(DEVPKEY_Device_Driver)

    @cached_property
    def driver_inf_path(self):
        return self.get_property(DEVPKEY_Device_DriverInfPath)

    @cached_property
    def manufacturer(self):
        return self.get_property(DEVPKEY_Device_Manufacturer)

    @cached_property
    def location_info(self):
        return self.get_property(DEVPKEY_Device_LocationInfo)

    @cached_property
    def address(self):
        return self.get_property(DEVPKEY_Device_Address)

    @cached_property
    def bus_reported_device_description(self):
        return self.get_property(DEVPKEY_Device_BusReportedDeviceDesc)

    @cached_property
    def friendly_name(self):
        return self.get_property(DEVPKEY_Device_FriendlyName)

    @cached_property
    def location_paths(self):
        return self.get_property(DEVPKEY_Device_LocationPaths)

    @cached_property
    def physical_device_object_name(self):
        return self.get_property(DEVPKEY_Device_PDOName)

    @cached_property
    def compatible_ids(self):
        return self.get_property(DEVPKEY_Device_CompatibleIds)

    @cached_property
    def service(self):
        return self.get_property(DEVPKEY_Device_Service)

    @cached_property
    def enumerator_name(self):
        return self.get_property(DEVPKEY_Device_EnumeratorName)

    @cached_property
    def bus_number(self):
        return self.get_property(DEVPKEY_Device_BusNumber)

    @property
    def power_data(self):
        buffer = self.get_property(DEVPKEY_Device_PowerData)
        if buffer is None:
            return None
        return CM_POWER_DATA.from_buffer_copy(buffer)

    @cached_property
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
        buffer_size = ctypes.c_ulong()
        status = RegQueryValueExW(
            hkey,
            "PortName",
            None,
            None,
            None,
            ctypes.byref(buffer_size)
        )
        if (status != ERROR_SUCCESS) and (status != ERROR_MORE_DATA):
            return None
        buffer = ctypes.create_unicode_buffer(buffer_size.value // 2)
        status = RegQueryValueExW(
            hkey,
            "PortName",
            None,
            None,
            buffer,
            ctypes.byref(buffer_size)
        )
        if status != ERROR_SUCCESS:
            return None
        RegCloseKey(hkey)
        return buffer.value

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


class DeviceInterface(DeviceNode):
    def __init__(self, interface):
        self.__interface = interface
        super().__init__(self.instance_handle)

    @property
    def interface(self):
        return self.__interface

    @cached_property
    def instance_identifier(self):
        return self.get_interface_property(DEVPKEY_Device_InstanceId)

    @cached_property
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
    def __init__(self, pid, vid, product, manufacturer, serial_number, location, function, interface):
        self.pid = pid
        self.vid = vid
        self.product = product
        self.manufacturer = manufacturer
        self.serial_number = serial_number
        self.location = location
        self.function = function
        self.interface = interface

    def __str__(self):
        return f'{self.location} - {self.vid:04X}:{self.pid:04X} - {self.product} - {self.manufacturer} - {self.serial_number} - {self.function} - {self.interface}'


def wake_up_device(port_device):
    # Ugly solution: trigger the port to wake up the usb device.
    port_handle = CreateFileW(
        '\\\\.\\' + port_device.port_name,
        GENERIC_READ | GENERIC_WRITE,
        0,
        None,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        0
    )
    if port_handle == INVALID_HANDLE_VALUE:
        raise ctypes.WinError(ctypes.GetLastError())
    CloseHandle(port_handle)


def get_usb_info(port_device: DeviceInterface, all_hub_devices=None, sorted_usb_host_controllers=None):
    # Enumerate and store the instance numbers of all connected usb hubs.
    if all_hub_devices is None:
        all_hub_devices = set(enumerate_device_interface(GUID_DEVINTERFACE_USB_HUB))
    if sorted_usb_host_controllers is None:
        sorted_usb_host_controllers = sorted(set(enumerate_device_interface(GUID_DEVINTERFACE_USB_HOST_CONTROLLER)), key=host_controller_sorting_key)

    # Look up parent hub and usb recursively.
    hub_and_usb = find_parent_hub_and_usb(port_device, all_hub_devices)
    if hub_and_usb is None:
        return None
    hub_device, usb_device = hub_and_usb

    # Check the power status of the usb device.
    power_data = usb_device.power_data
    if power_data is not None:
        if power_data.PD_MostRecentPowerState != 1:
            # The device is in sleep state. It needs to be woken up to D0 state to get the usb description.
            wake_up_device(port_device)

    # Generate the hub path from PDO name and open it.
    # Solution 1: https://stackoverflow.com/questions/28007468/how-do-i-obtain-usb-device-descriptor-given-a-device-path
    # hub_path = "\\\\?\\" + hub_device.instance_identifier.replace("\\", "#") + f"#{GUID_DEVINTERFACE_USB_HUB}"
    # Solution 2: https://stackoverflow.com/questions/21703592/open-device-name-using-createfile
    hub_path = "\\\\?\\Global\\GLOBALROOT" + hub_device.physical_device_object_name
    h_hub_device = CreateFileW(
        hub_path,
        GENERIC_WRITE,
        2,
        None,
        OPEN_EXISTING,
        0,
        None
    )
    if h_hub_device == INVALID_HANDLE_VALUE:
        return None

    # Get the port number that the usb device is connected to.
    # https://learn.microsoft.com/en-us/windows-hardware/drivers/install/devpkey-device-address
    usb_hub_port = usb_device.address
    if usb_hub_port is None:
        return None

    # Request usb device connection info.
    connection_info = request_usb_connection_info(h_hub_device, usb_hub_port)
    if connection_info is None:
        # Fallback to request usb device description.
        device_description = request_usb_device_description(h_hub_device, usb_hub_port)
        if device_description is None:
            # We can't get any information.
            CloseHandle(h_hub_device)
            return None
    else:
        device_description = connection_info.DeviceDescriptor

    # Get product identifier and vendor identifier.
    pid = device_description.idProduct
    vid = device_description.idVendor

    # Request string description at iProduct, iManufacturer and iSerialNumber.
    language_id = request_and_select_language_id(h_hub_device, usb_hub_port)
    product = request_usb_string_description(h_hub_device, usb_hub_port, device_description.iProduct, language_id)
    manufacturer = request_usb_string_description(h_hub_device, usb_hub_port, device_description.iManufacturer, language_id)
    serial_number = request_usb_string_description(h_hub_device, usb_hub_port, device_description.iSerialNumber, language_id)

    # Get the bConfigurationValue.
    bConfigurationValue = None
    if connection_info is not None:
        # Get bConfigurationValue from device connection info.
        bConfigurationValue = connection_info.CurrentConfigurationValue
    elif device_description.bNumConfigurations == 1:
        # Mostly, there is only one Configuration, and bConfigurationValue is equal to 1.
        bConfigurationValue = 1

    # Get the bInterfaceNumber.
    bInterfaceNumber = None
    location_paths = port_device.location_paths
    if location_paths is not None:
        # Parse interface number from location paths.
        for p in location_paths:
            # Find the #USBMI(?) patten in the last part of location string.
            m = re.fullmatch(r'.*?#USBMI\((\d+)\)', p, re.IGNORECASE)
            if m is not None:
                bInterfaceNumber = int(m.group(1))
                break
    else:
        # Try to parse interface number from hardware identifier string.
        bInterfaceNumber = parse_interface_number(port_device.instance_identifier)

    # Read interface description.
    function = None
    interface = None
    if bConfigurationValue is not None:
        configuration_description = request_usb_configuration_description(h_hub_device, usb_hub_port, bConfigurationValue)
        if configuration_description is not None:
            if device_description.bNumConfigurations == 1 and configuration_description.bNumInterfaces == 1:
                # Usb device with single interface.
                bInterfaceNumber = None  # Set to None to distinguish composite devices.
            # Get interface string of bInterfaceNumber.
            # There is only one alternative interface in most cases (bAlternateSetting == 0).
            # If we don't know bAlternateSetting, then use the default interface.
            interface_association_description, interface_description = request_usb_interface_descriptions(h_hub_device, usb_hub_port, configuration_description, bInterfaceNumber, 0)
            if interface_association_description is not None:
                function = request_usb_string_description(h_hub_device, usb_hub_port, interface_association_description.iFunction, language_id)
            if interface_description is not None:
                interface = request_usb_string_description(h_hub_device, usb_hub_port, interface_description.iInterface, language_id)

    # Generate location string compatible with linux.
    location = get_location_string(
        usb_device,
        sorted_usb_host_controllers,
        bConfigurationValue=bConfigurationValue,
        bInterfaceNumber=bInterfaceNumber
    )

    # Everything is fine! Finished and close the hub handle.
    CloseHandle(h_hub_device)
    return USBInfo(pid, vid, product, manufacturer, serial_number, location, function, interface)


def get_host_controller_information(c):
    h_device = CreateFileW(
        c.interface,
        GENERIC_WRITE,
        2,
        None,
        OPEN_EXISTING,
        0,
        None
    )
    if h_device == INVALID_HANDLE_VALUE:
        return None

    usb_controller_info = USBUSER_CONTROLLER_INFO_0()
    usb_controller_info.Header.UsbUserRequest = USBUSER_GET_CONTROLLER_INFO_0
    usb_controller_info.Header.RequestBufferLength = ctypes.sizeof(usb_controller_info)

    success = DeviceIoControl(
        h_device,
        IOCTL_USB_USER_REQUEST,
        ctypes.byref(usb_controller_info),
        ctypes.sizeof(usb_controller_info),
        ctypes.byref(usb_controller_info),
        ctypes.sizeof(usb_controller_info),
        None,
        None
    )
    if not success:
        return None

    CloseHandle(h_device)
    return usb_controller_info


def host_controller_sorting_key(c, compat_linux=False):
    compare_list = []

    if compat_linux:
        # Sort the usb host controllers in the unix fashion.
        # https://unix.stackexchange.com/questions/297178/how-usb-bus-number-and-device-number-been-assigned
        usb_controller_info = get_host_controller_information(c)

        if usb_controller_info is not None:
            compare_list.append(usb_controller_info.Info0.ControllerFlavor)
        else:
            compare_list.append(float('inf'))

        bus_number = c.bus_number
        if bus_number is None:
            compare_list.append(float('inf'))
        else:
            compare_list.append(bus_number)

        address = c.address
        if address is None:
            compare_list.extend([float('inf'), c.instance_handle])
        else:
            device = c.address >> 16
            function = c.address & 0xffff
            compare_list.extend([device, function])
    else:
        compare_list.append(c.instance_handle)

    return compare_list


def get_bus_number(usb_device, sorted_usb_host_controllers):
    # Get USB bus number (usb host controller).
    while usb_device is not None:
        if usb_device in sorted_usb_host_controllers:
            return sorted_usb_host_controllers.index(usb_device) + 1
        usb_device = usb_device.parent
    return None


def get_location_string(usb_device, sorted_usb_host_controllers, bConfigurationValue=None, bInterfaceNumber=None):
    location_paths = usb_device.location_paths
    if not location_paths:
        return None
    if (bConfigurationValue is None) and (bInterfaceNumber is not None):
        bConfigurationValue = 'x'
    bus_number = get_bus_number(usb_device, sorted_usb_host_controllers)
    if bus_number is None:
        return None
    location = [str(bus_number)]
    for g in re.finditer(r'#USB\((\w+)\)', location_paths[0]):
        if len(location) > 1:
            location.append('.')
        else:
            location.append('-')
        location.append(g.group(1))
    if bInterfaceNumber is not None:
        location.append(':{}.{}'.format(
            bConfigurationValue,
            bInterfaceNumber
        ))
    return ''.join(location)


def get_root_hub_name(pci_device):
    hdev = CreateFileW(
        pci_device.interface,
        GENERIC_WRITE,
        2,
        None,
        OPEN_EXISTING,
        0,
        None
    )
    if hdev == INVALID_HANDLE_VALUE:
        return None
    root_hub_name_request = USB_ROOT_HUB_NAME()
    success = DeviceIoControl(
        hdev,
        IOCTL_USB_GET_ROOT_HUB_NAME,
        0,
        0,
        ctypes.byref(root_hub_name_request),
        ctypes.sizeof(root_hub_name_request),
        None,
        None
    )
    if not success:
        CloseHandle(hdev)
        return None
    root_hub_name_buffer = ctypes.create_string_buffer(root_hub_name_request.ActualLength)
    success = DeviceIoControl(
        hdev,
        IOCTL_USB_GET_ROOT_HUB_NAME,
        0,
        0,
        root_hub_name_buffer,
        ctypes.sizeof(root_hub_name_buffer),
        None,
        None
    )
    if not success:
        CloseHandle(hdev)
        return None
    CloseHandle(hdev)
    return ctypes.wstring_at(root_hub_name_buffer[ctypes.sizeof(USB_ROOT_HUB_NAME) - 2:ctypes.sizeof(root_hub_name_buffer)])


def get_external_hub_name(h_hub_device, port):
    node_name_request = USB_NODE_CONNECTION_NAME()
    node_name_request.ConnectionIndex = port
    success = DeviceIoControl(
        h_hub_device,
        IOCTL_USB_GET_NODE_CONNECTION_NAME,
        ctypes.byref(node_name_request),
        ctypes.sizeof(node_name_request),
        ctypes.byref(node_name_request),
        ctypes.sizeof(node_name_request),
        None,
        None
    )
    if not success:
        return None
    node_name_buffer = ctypes.create_string_buffer(node_name_request.ActualLength)
    ctypes.cast(node_name_buffer, ctypes.POINTER(USB_NODE_CONNECTION_NAME)).contents.ConnectionIndex = port
    success = DeviceIoControl(
        h_hub_device,
        IOCTL_USB_GET_NODE_CONNECTION_NAME,
        node_name_buffer,
        ctypes.sizeof(node_name_buffer),
        node_name_buffer,
        ctypes.sizeof(node_name_buffer),
        None,
        None
    )
    if not success:
        return None
    return ctypes.wstring_at(node_name_buffer[ctypes.sizeof(USB_NODE_CONNECTION_NAME) - 2:ctypes.sizeof(node_name_buffer)])


def parse_location_string(location, sorted_usb_host_controllers=None):
    # Parses the location string and returns human-readable node names.
    # http://www.linux-usb.org/FAQ.html#i6
    nodes = []
    if sorted_usb_host_controllers is None:
        sorted_usb_host_controllers = sorted(set(enumerate_device_interface(GUID_DEVINTERFACE_USB_HOST_CONTROLLER)), key=host_controller_sorting_key)
    bus_num, port_chain = location.split('-')
    pci_device = sorted_usb_host_controllers[int(bus_num) - 1]
    m = re.match(r'PCI\\VEN_([0-9a-f]{4})&DEV_([0-9a-f]{4})&SUBSYS_[0-9a-f]{8}&REV_[0-9a-f]{2}', pci_device.instance_identifier, re.IGNORECASE)
    if m is None:
        return None
    nodes.append(f'Bus {bus_num}: {pci_device.friendly_name}')

    hub_path = "\\\\.\\" + get_root_hub_name(pci_device)
    for port in port_chain:
        if port == '.':
            # Skip port spliter
            continue
        if port == ':':
            # Finally outputs the configuration value and interface number.
            configuration, interface = port_chain.split(':')[-1].split('.')
            nodes.append(f'Configuration {configuration}')
            nodes.append(f'Interface {interface}')
            break

        # Convert port to int.
        port = int(port)

        # Open current hub device
        h_hub_device = CreateFileW(
            hub_path,
            GENERIC_WRITE,
            2,
            None,
            OPEN_EXISTING,
            0,
            None
        )
        if h_hub_device == INVALID_HANDLE_VALUE:
            return None

        device_description = request_usb_device_description(h_hub_device, port)

        language_id = request_and_select_language_id(h_hub_device, port)
        product = request_usb_string_description(h_hub_device, port, device_description.iProduct, language_id)
        manufacturer = request_usb_string_description(h_hub_device, port, device_description.iManufacturer, language_id)
        nodes.append(f'Port {port}: {manufacturer} - {product}')

        hub_path = "\\\\.\\" + get_external_hub_name(h_hub_device, port)
        CloseHandle(h_hub_device)
    return nodes


def parse_interface_number(instance_identifier):
    # https://learn.microsoft.com/en-us/windows-hardware/drivers/install/standard-usb-identifiers
    # USB example: USB\VID_303A&PID_4001&MI_00\6&182C0AC9&0&0000
    m = re.fullmatch(r'USB\\VID_[0-9a-f]{4}&PID_[0-9a-f]{4}&MI_(\d{2})\\.*?', instance_identifier, re.IGNORECASE)
    if m is not None:
        return int(m.group(1))
    # FTDIBUS example: FTDIBUS\VID_0403+PID_6001+B001ADZMA\0000
    m = re.fullmatch(r'FTDIBUS\\.*\\([0-9a-f]*)', instance_identifier, re.IGNORECASE)
    if m is not None:
        return int(m.group(1))
    return None


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


def find_parent_hub_and_usb(device_node, all_hub_devices):
    # Get parent hub device and usb device recursively.
    child_device = device_node
    while True:
        parent_device = child_device.parent
        if parent_device is None:
            return None
        if parent_device in all_hub_devices:
            return parent_device, child_device
        child_device = parent_device


def split_text_number(s):
    result = []
    for group in re.split(r'(\d+)', s):
        if group:
            try:
                group = int(group)
            except ValueError:
                pass
            result.append(group)
    return result


def port_name_sorting_key(c):
    # Used to sort COM ports in a language-conforming manner.
    if isinstance(c, tuple):
        c = c[0]
    return split_text_number(c.port_name)


def iterate_comports(retrieve_usb_info=False, unique=True):
    # Enumerate and store the instance numbers of all connected usb hubs.
    if retrieve_usb_info:
        all_hub_devices = set(enumerate_device_interface(GUID_DEVINTERFACE_USB_HUB))
        sorted_usb_host_controllers = sorted(set(enumerate_device_interface(GUID_DEVINTERFACE_USB_HOST_CONTROLLER)), key=host_controller_sorting_key)

    # Generate non-repeating serial devices.
    if unique:
        yielded_devices = []

    # Repeat for all possible GUIDs.
    # Use GUID_DEVINTERFACE_COMPORT and GUID_DEVINTERFACE_MODEM instead of GUID_CLASS_COMPORT and GUID_CLASS_MODEM.
    # https://code.qt.io/cgit/qt/qtserialport.git/commit/?id=63bfe5ea4203af3c294691216ddfb7dc29f310f7
    for guid in [GUID_DEVINTERFACE_COMPORT, GUID_DEVINTERFACE_MODEM]:
        # Iterate through each serial device.
        for port_device in enumerate_device_interface(guid):
            if unique:
                if port_device in yielded_devices:
                    continue
                else:
                    yielded_devices.append(port_device)
            # Request usb information.
            if retrieve_usb_info:
                usb_info = get_usb_info(port_device, all_hub_devices, sorted_usb_host_controllers)
                # usb_info may be None because some serial ports are not usb interfaces (e.g. virtual serial ports).
                yield port_device, usb_info
            else:
                yield port_device


def find_device_from_port_name(port_name):
    # Finds the corresponding device node from the com port name (COMx).
    for port_device in iterate_comports():
        if port_device.port_name == port_name:
            return port_device
    return None


def pyserial_comports():
    # Generate pyserial-compatible ListPortInfo.
    from serial.tools.list_ports_common import ListPortInfo

    for port_device, usb_info in iterate_comports(retrieve_usb_info=True, unique=True):
        info = ListPortInfo(port_device.port_name, skip_link_detection=True)
        if usb_info is None:
            info.name = port_device.name
            info.hwid = port_device.instance_identifier
            info.manufacturer = port_device.manufacturer
            info.description = port_device.description
            info.interface = port_device.bus_reported_device_description
            info.location = port_device.location_info
        else:
            info.name = port_device.name
            info.pid = usb_info.pid
            info.vid = usb_info.vid
            info.product = usb_info.product
            info.manufacturer = usb_info.manufacturer
            info.serial_number = usb_info.serial_number
            info.location = usb_info.location
            info.interface = usb_info.interface
            info.apply_usb_info()
        yield info


def comports(include_links=False):
    # Compatible with serial/tools/list_ports_windows.py in pyserial.
    return list(pyserial_comports())


def test():
    for port_device, usb_info in sorted(iterate_comports(retrieve_usb_info=True), key=port_name_sorting_key):
        print(f'[{port_device.port_name}]')
        if usb_info is None:
            print(f'Description: {port_device.description}')
            print(f'Manufacturer: {port_device.manufacturer}')
        else:
            if usb_info.location is not None:
                for depth, node in enumerate(parse_location_string(usb_info.location)):
                    if depth == 0:
                        print(node)
                    else:
                        print('    ' * (depth - 1) + '|__ ' + node)
            print(f'Location: {usb_info.location}')
            print(f'Vendor: {usb_info.vid:04X}')
            print(f'Product: {usb_info.pid:04X}')
            print(f'Manufacturer: {usb_info.manufacturer}')
            print(f'Product: {usb_info.product}')
            print(f'SerialNumber: {usb_info.serial_number}')
            print(f'Function: {usb_info.function}')
            print(f'Interface: {usb_info.interface}')
        print()
    print("If we've lost some serial ports or have incorrect information listed, please open an issue on github to let us know.")
    print('Link: https://github.com/chinaheyu/list_ports_windows/issues/new')


if __name__ == '__main__':
    test()


__all__ = [
    'get_usb_info',
    'find_device_from_port_name',
    'iterate_comports',
    'pyserial_comports',
    'comports',
    'DeviceInterface',
    'USBInfo',
    'parse_location_string'
]
