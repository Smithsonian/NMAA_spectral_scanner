Option Strict Off
Option Explicit On
Imports System
Imports System.Text
Imports Microsoft.Win32.SafeHandles
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports VB = Microsoft.VisualBasic
Imports System.Formats.Asn1

Public Module modWinUSB



    Public CurrentUSBDevice As Integer
    Public NumUSBDevices As Integer
    Public Const DPP_WINUSB_GUID_STRING As String = "{5A8ED6A1-7FC3-4b6a-A536-95DF35D03448}"

    Public Enum ERRORLEVEL
        elevCodeReturn      'Return error code.
        elevDescReturn      'Return error description string.
        elevSrcDescReturn   'Return error description string and error source string.
        elevMsgBox          'Return/Display error description string and error source string.
    End Enum

    Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Integer, ByRef lpSource As Short, ByVal dwMessageId As Integer, ByVal dwLanguageZId As Integer, ByVal lpBuffer As String, ByVal nSize As Integer, ByVal Arguments As Integer) As Integer
    Public Const FORMAT_MESSAGE_FROM_SYSTEM As Short = &H1000S
    Public Const DBT_DEVICEARRIVAL As Int32 = 32768
    Public Const DBT_DEVICEREMOVECOMPLETE As Int32 = 32772
    Public Const DBT_DEVTYP_DEVICEINTERFACE As Int32 = 5
    Public Const DEVICE_NOTIFY_WINDOW_HANDLE As Int32 = 0
    Public Const WM_DEVICECHANGE As Int32 = &H219S
    Public Const DIGCF_PRESENT As Int16 = &H2S
    Public Const DIGCF_DEVICEINTERFACE As Int16 = &H10S

    Public Structure DEV_BROADCAST_DEVICEINTERFACE
        Dim dbcc_size As Int32
        Dim dbcc_devicetype As Int32
        Dim dbcc_reserved As Int32
        Dim dbcc_classguid As Guid
        Dim dbcc_name As Int16
    End Structure

    Public Structure DEV_BROADCAST_DEVICEINTERFACE2
        Dim dbcc_size As Int32
        Dim dbcc_devicetype As Int32
        Dim dbcc_reserved As Int32
        Dim dbcc_classguid As Guid
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=1024)> Friend dbcc_name() As Char
    End Structure

    Public Structure DEV_BROADCAST_HDR
        Dim dbch_size As Int32
        Dim dbch_devicetype As Int32
        Dim dbch_reserved As Int32
    End Structure

    Public Structure SP_DEVICE_INTERFACE_DATA
        Dim cbSize As Int32
        Dim InterfaceClassGuid As System.Guid
        Dim Flags As Int32
        Dim Reserved As Int32
    End Structure

    Public Structure SP_DEVICE_INTERFACE_DETAIL_DATA
        Dim cbSize As Int32
        Dim DevicePath As String
    End Structure



    'older code
    'Public Declare Function SetupDiDestroyDeviceInfoList Lib "setupapi.dll" (ByVal DeviceInfoSet As IntPtr) As Int32
    'Public Declare Function SetupDiEnumDeviceInterfaces Lib "setupapi.dll" (ByVal DeviceInfoSet As IntPtr, ByVal DeviceInfoData As Int32, ByRef InterfaceClassGuid As System.Guid, ByVal MemberIndex As Int32, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA) As Boolean
    'Public Declare Function SetupDiGetClassDevs Lib "setupapi.dll" Alias "SetupDiGetClassDevsW" (ByRef ClassGuid As System.Guid, ByVal Enumerator As String, ByVal hwndParent As Int32, ByVal Flags As Int32) As IntPtr
    'Public Declare Function SetupDiGetDeviceInterfaceDetail Lib "setupapi.dll" Alias "SetupDiGetDeviceInterfaceDetailW" (ByVal DeviceInfoSet As IntPtr, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA, ByVal DeviceInterfaceDetailData As IntPtr, ByVal DeviceInterfaceDetailDataSize As Int32, ByRef RequiredSize As Int32, ByVal DeviceInfoData As IntPtr) As Boolean

    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function SetupDiDestroyDeviceInfoList(ByVal DeviceInfoSet As IntPtr) As Int32
    End Function

    <DllImport("setupapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function SetupDiEnumDeviceInterfaces(ByVal DeviceInfoSet As IntPtr, ByVal DeviceInfoData As Int32, ByRef InterfaceClassGuid As System.Guid, ByVal MemberIndex As Int32, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA) As Boolean
    End Function

    <DllImport("setupapi.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Function SetupDiGetClassDevsW(ByRef ClassGuid As Guid, ByVal Enumerator As String, ByVal hwndParent As IntPtr, ByVal Flags As Integer) As IntPtr
    End Function


    <DllImport("setupapi.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Function SetupDiGetDeviceInterfaceDetailW(ByVal DeviceInfoSet As IntPtr, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA, ByVal DeviceInterfaceDetailData As IntPtr, ByVal DeviceInterfaceDetailDataSize As Int32, ByRef RequiredSize As Int32, ByVal DeviceInfoData As IntPtr) As Boolean
    End Function




    'onecore code - win10 and later - DID NOT WORK, NEED TO FIND NEW FUNCTIONS
    'Public Declare Function SetupDiDestroyDeviceInfoList Lib "OneCore.lib" (ByVal DeviceInfoSet As IntPtr) As Int32
    'Public Declare Function SetupDiEnumDeviceInterfaces Lib "OneCore.lib" (ByVal DeviceInfoSet As IntPtr, ByVal DeviceInfoData As Int32, ByRef InterfaceClassGuid As System.Guid, ByVal MemberIndex As Int32, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA) As Boolean
    'Public Declare Function SetupDiGetClassDevs Lib "OneCore.lib" Alias "SetupDiGetClassDevsW" (ByRef ClassGuid As System.Guid, ByVal Enumerator As String, ByVal hwndParent As Int32, ByVal Flags As Int32) As IntPtr
    'Public Declare Function SetupDiGetDeviceInterfaceDetail Lib "OneCore.lib" Alias "SetupDiGetDeviceInterfaceDetailW" (ByVal DeviceInfoSet As IntPtr, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA, ByVal DeviceInterfaceDetailData As IntPtr, ByVal DeviceInterfaceDetailDataSize As Int32, ByRef RequiredSize As Int32, ByVal DeviceInfoData As IntPtr) As Boolean


    Public Const FILE_ATTRIBUTE_NORMAL As Int16 = &H80S
    Public Const FILE_FLAG_OVERLAPPED As Int32 = &H40000000
    Public Const FILE_SHARE_READ As Int16 = &H1S
    Public Const FILE_SHARE_WRITE As Int16 = &H2S
    Public Const GENERIC_READ As Long = &H80000000
    Public Const GENERIC_WRITE As UInt32 = &H40000000
    Public Const OPEN_EXISTING As Int16 = 3

    Public Structure SECURITY_ATTRIBUTES
        Dim nLength As Integer
        Dim lpSecurityDescriptor As Integer
        Dim bInheritHandle As Integer
    End Structure

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function CloseHandle(ByVal hObject As SafeFileHandle) As Boolean
    End Function

    <DllImport("kernel32.dll", EntryPoint:="CreateFileW", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Function CreateFile(ByVal lpFileName As String, ByVal dwDesiredAccess As Int32, ByVal dwShareMode As Int32, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Int32, ByVal dwFlagsAndAttributes As Int32, ByVal hTemplateFile As Int32) As SafeFileHandle
    End Function

    Public Enum PIPE_POLICY_TYPE
        SHORT_PACKET_TERMINATE = &H1S
        AUTO_CLEAR_STALL
        PIPE_TRANSFER_TIMEOUT
        IGNORE_SHORT_PACKETS
        ALLOW_PARTIAL_READS
        AUTO_FLUSH
        RAW_IO
        MAXIMUM_TRANSFER_SIZE
    End Enum

    Public Structure WINUSB_PIPE_INFORMATION
        Dim PipeType As USBD_PIPE_TYPE
        Dim PipeId As Byte
        Dim MaximumPacketSize As Short
        Dim Interval As Byte
    End Structure

    Public Enum USBD_PIPE_TYPE
        UsbdPipeTypeControl
        UsbdPipeTypeIsochronous
        UsbdPipeTypeBulk
        UsbdPipeTypeInterrupt
    End Enum

    Public Structure USB_INTERFACE_DESCRIPTOR
        Dim bLength As Byte
        Dim bDescriptorType As Byte
        Dim bInterfaceNumber As Byte
        Dim bAlternateSetting As Byte
        Dim bNumEndpoints As Byte
        Dim bInterfaceClass As Byte
        Dim bInterfaceSubClass As Byte
        Dim bInterfaceProtocol As Byte
        Dim iInterface As Byte
    End Structure

    <DllImport("winusb.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function WinUsb_Initialize(ByVal DeviceHandle As SafeFileHandle, ByRef InterfaceHandle As IntPtr) As Boolean
    End Function

    <DllImport("winusb.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function WinUsb_Free(ByVal InterfaceHandle As IntPtr) As Boolean
    End Function

    <DllImport("winusb.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function WinUsb_QueryInterfaceSettings(ByVal InterfaceHandle As IntPtr, ByVal AlternateInterfaceNumber As Byte, ByRef UsbAltInterfaceDescriptor As USB_INTERFACE_DESCRIPTOR) As Boolean
    End Function

    <DllImport("winusb.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function WinUsb_QueryPipe(ByVal InterfaceHandle As IntPtr, ByVal AlternateInterfaceNumber As Byte, ByVal PipeIndex As Byte, ByRef PipeInformation As WINUSB_PIPE_INFORMATION) As Boolean
    End Function

    <DllImport("winusb.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function WinUsb_SetPipePolicy(ByVal InterfaceHandle As IntPtr, ByVal PipeID As Byte, ByVal PolicyType As UInt32, ByVal ValueLength As UInt32, ByRef Value As Byte) As Boolean
    End Function

    <DllImport("winusb.dll", EntryPoint:="WinUsb_SetPipePolicy", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function WinUsb_SetPipePolicy1(ByVal InterfaceHandle As IntPtr, ByVal PipeID As Byte, ByVal PolicyType As UInt32, ByVal ValueLength As UInt32, ByRef Value As UInt32) As Boolean
    End Function

    <DllImport("winusb.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function WinUsb_ReadPipe(ByVal InterfaceHandle As IntPtr, ByVal PipeID As Byte, ByRef Buffer As Byte, ByVal BufferLength As UInt32, ByRef LengthTransferred As UInt32, ByVal Overlapped As IntPtr) As Boolean
    End Function

    <DllImport("winusb.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function WinUsb_WritePipe(ByVal InterfaceHandle As IntPtr, ByVal PipeID As Byte, ByRef Buffer As Byte, ByVal BufferLength As UInt32, ByRef LengthTransferred As UInt32, ByVal Overlapped As IntPtr) As Boolean
    End Function

    <DllImport("winusb.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function WinUsb_ResetPipe(ByVal InterfaceHandle As IntPtr, ByVal PipeId As Byte) As Boolean
    End Function

    <DllImport("winusb.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function WinUsb_AbortPipe(ByVal InterfaceHandle As IntPtr, ByVal PipeId As Byte) As Boolean
    End Function

    <DllImport("winusb.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function WinUsb_FlushPipe(ByVal InterfaceHandle As IntPtr, ByVal PipeId As Byte) As Boolean
    End Function

    Public Structure WinUsbDevice
        Dim DeviceHandle As SafeFileHandle  'Communications device resource handle.
        Dim winUsbHandle As IntPtr          'WinUSB interface handle.
        Dim bulkInPipe As Int32             'Bulk IN pipe id.
        Dim bulkOutPipe As Int32            'Bulk OUT pipe id.
        Dim interruptInPipe As Int32        'Interrupt IN pipe id.
        Dim interruptOutPipe As Int32       'Interrupt OUT pipe id.
        Dim devicespeed As UInt32           'Device speed indicator.
    End Structure

    Public DppWinUSB As WinUsbDevice 'DppWinUSB device information

    Public USBDeviceNotificationHandle As Integer
    Public USBDeviceConnected As Boolean
    Public USBDevicePathName As String

    'Registry functions for counting WinUSB devices
    <DllImport("advapi32.dll", EntryPoint:="RegOpenKeyExA", CharSet:=CharSet.Ansi, SetLastError:=True)>
    Public Function RegOpenKeyEx(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
    End Function

    <DllImport("advapi32.dll", EntryPoint:="RegEnumValueA", CharSet:=CharSet.Ansi, SetLastError:=True)>
    Public Function RegEnumValue(ByVal hKey As Integer, ByVal dwIndex As Integer, ByVal lpValueName As String, ByRef lpcbValueName As Integer, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Byte, ByRef lpcbData As Integer) As Integer
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Function RegCloseKey(ByVal hKey As Integer) As Integer
    End Function

    Public Const HKEY_LOCAL_MACHINE As Integer = &H80000002
    Public Const KEY_QUERY_VALUE As Short = &H1S
    Public Const ERROR_SUCCESS As Short = 0
    Public Const MAX_PATH As Short = 260
    Public Const MAXREGBUFFER As Short = 128
    Public Const MAXERRBUFFER As Short = 256
    Public Const MAX_DEVPATH_LENGTH As Short = 256
    Public Const NUM_ASYNCH_IO As Short = 100
    Public Const BUFFER_SIZE As Short = 1024

    Public Const MAXDP5S As Short = 128 ' max number of devices
    Public Const WinUSBService As String = "SYSTEM\CurrentControlSet\Services\WinUSB\Enum"
    Public Const WinUsbDP5 As String = "USB\Vid_10c4&Pid_842a"
    Public Const WinUsbDP5Size As Short = 21
    Public Const ERROR_NO_MORE_ITEMS As Short = 259

    Public Const REG_NONE As Short = 0 ' No value type
    Public Const REG_DWORD As Short = 4 ' 32-bit number

    Public Function OpenDevice(ByRef DppWinUSB As WinUsbDevice, ByRef DeviceConnected As Boolean, ByRef DevicePathName As String, ByRef MemberIndex As Integer) As Boolean
        Dim deviceFound As Boolean
        Dim NewDevicePathName As String
        Dim success As Boolean
        Dim winUsbDemoGuid As New System.Guid(DPP_WINUSB_GUID_STRING)
        OpenDevice = False
        NewDevicePathName = ""
        On Error GoTo OpenDeviceErr
        'If Not DeviceConnected Then
        'If DeviceConnected Then 'The device was detected.
        CloseDeviceHandle(DppWinUSB)
        'End If


        'Fill an array with the device path names of all attached devices with matching GUIDs.
        deviceFound = FindDeviceFromGuid(winUsbDemoGuid, NewDevicePathName, MemberIndex)
        If deviceFound = True Then
            success = GetDeviceHandle(DppWinUSB, NewDevicePathName)
            If (success) Then
                DeviceConnected = True
                DevicePathName = NewDevicePathName 'Save NewDevicePathName so OnDeviceChange() knows which name is Test device.
            Else 'There was a problem in retrieving the information.
                DeviceConnected = False
                CloseDeviceHandle(DppWinUSB)
            End If
        End If
        If DeviceConnected Then 'The device was detected.
            InitializeDevice(DppWinUSB)
        End If
        'End If
        OpenDevice = DeviceConnected
        Exit Function
OpenDeviceErr:
        MessageBox.Show("OpenDeviceErr")
        ProcessError(Err)
    End Function

    Public Sub CloseDeviceHandle(ByRef DppWinUSB As WinUsbDevice)
        'updated to .net code
        Try
            If DppWinUSB.winUsbHandle <> IntPtr.Zero Then
                If WinUsb_Free(DppWinUSB.winUsbHandle) Then
                    DppWinUSB.winUsbHandle = IntPtr.Zero
                End If
            End If

            If DppWinUSB.DeviceHandle IsNot Nothing AndAlso Not DppWinUSB.DeviceHandle.IsInvalid Then
                DppWinUSB.DeviceHandle.Close() ' Let SafeFileHandle manage it
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error DppWinUSB", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Function GetDeviceHandle(ByRef DppWinUSB As WinUsbDevice, ByVal DevicePathName As String) As Boolean
        Dim security As SECURITY_ATTRIBUTES
        security.lpSecurityDescriptor = 0
        security.bInheritHandle = 1
        security.nLength = Len(security)
        DppWinUSB.DeviceHandle = CreateFile(DevicePathName, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, security, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL Or FILE_FLAG_OVERLAPPED, 0)
        If (Not DppWinUSB.DeviceHandle.IsInvalid) Then
            GetDeviceHandle = True
        Else
            MsgBox(GetErrorString(Err.LastDllError))
            GetDeviceHandle = False
        End If
    End Function

    Public Function InitializeDevice(ByRef DppWinUSB As WinUsbDevice) As Boolean
        Dim ifaceDescriptor As USB_INTERFACE_DESCRIPTOR
        Dim pipeInfo As WINUSB_PIPE_INFORMATION
        Dim pipeTimeout As Integer
        Dim success As Boolean
        Dim success1 As Boolean
        Dim i As Integer

        pipeTimeout = 500 ' 2/4/10 - seems like 500mS should be plenty?
        'Erase FPGA' takes 4s, worst case! Change timeout just for that command
        On Error GoTo InitializeDeviceErr
        success = WinUsb_Initialize(DppWinUSB.DeviceHandle, DppWinUSB.winUsbHandle)
        If success Then
            success = WinUsb_QueryInterfaceSettings(DppWinUSB.winUsbHandle, 0, ifaceDescriptor)
            If success Then
                ' Get the transfer type, endpoint number, and direction for the interface's
                ' bulk and interrupt endpoints. Set pipe policies.
                For i = 0 To ifaceDescriptor.bNumEndpoints - 1
                    success1 = WinUsb_QueryPipe(DppWinUSB.winUsbHandle, 0, CByte(i), pipeInfo)
                    If ((pipeInfo.PipeType = USBD_PIPE_TYPE.UsbdPipeTypeBulk) And UsbEndpointDirectionIn(pipeInfo.PipeId)) Then
                        DppWinUSB.bulkInPipe = pipeInfo.PipeId
                        SetPipePolicy(DppWinUSB, CByte(DppWinUSB.bulkInPipe), CShort(PIPE_POLICY_TYPE.IGNORE_SHORT_PACKETS), False) ' this is the default, but set it anyway
                        SetPipePolicy(DppWinUSB, CByte(DppWinUSB.bulkInPipe), CShort(PIPE_POLICY_TYPE.AUTO_CLEAR_STALL), True) ' new - 2/5/2010
                        SetPipePolicy(DppWinUSB, CByte(DppWinUSB.bulkInPipe), CShort(PIPE_POLICY_TYPE.ALLOW_PARTIAL_READS), False) ' new - 5/25/2010
                        SetPipePolicy1(DppWinUSB, CByte(DppWinUSB.bulkInPipe), CShort(PIPE_POLICY_TYPE.PIPE_TRANSFER_TIMEOUT), pipeTimeout)
                    ElseIf ((pipeInfo.PipeType = USBD_PIPE_TYPE.UsbdPipeTypeBulk) And UsbEndpointDirectionOut(pipeInfo.PipeId)) Then
                        DppWinUSB.bulkOutPipe = pipeInfo.PipeId
                        SetPipePolicy(DppWinUSB, CByte(DppWinUSB.bulkOutPipe), CShort(PIPE_POLICY_TYPE.SHORT_PACKET_TERMINATE), True)
                        SetPipePolicy(DppWinUSB, CByte(DppWinUSB.bulkOutPipe), CShort(PIPE_POLICY_TYPE.AUTO_CLEAR_STALL), True)
                        SetPipePolicy1(DppWinUSB, CByte(DppWinUSB.bulkOutPipe), CShort(PIPE_POLICY_TYPE.PIPE_TRANSFER_TIMEOUT), pipeTimeout)
                    End If
                Next i
            Else
                success = False
            End If
        End If
        InitializeDevice = success
        Exit Function
InitializeDeviceErr:
        ProcessError(Err)
    End Function

    Private Function SetPipePolicy(ByRef DppWinUSB As WinUsbDevice, ByVal PipeId As Byte, ByVal PolicyType As Integer, ByVal Value As Boolean) As Boolean
        Dim success As Boolean
        On Error GoTo SetPipePolicyErr
        success = WinUsb_SetPipePolicy(DppWinUSB.winUsbHandle, PipeId, PolicyType, 1, CByte(Value))
        SetPipePolicy = success
        Exit Function
SetPipePolicyErr:
        ProcessError(Err)
    End Function

    Public Function SetPipePolicy1(ByRef DppWinUSB As WinUsbDevice, ByVal PipeId As Byte, ByVal PolicyType As Integer, ByVal Value As Integer) As Boolean
        Dim success As Boolean
        SetPipePolicy1 = False
        On Error GoTo SetPipePolicy1Err
        success = WinUsb_SetPipePolicy1(DppWinUSB.winUsbHandle, PipeId, PolicyType, 4, Value)
        SetPipePolicy1 = success
        Exit Function
SetPipePolicy1Err:
        ProcessError(Err)
    End Function

    Private Function UsbEndpointDirectionIn(ByVal addr As Integer) As Boolean
        On Error GoTo UsbEndpointDirectionInErr
        If ((addr And &H80S) = &H80S) Then
            UsbEndpointDirectionIn = True
        Else
            UsbEndpointDirectionIn = False
        End If
        Exit Function
UsbEndpointDirectionInErr:
        ProcessError(Err)
    End Function

    Private Function UsbEndpointDirectionOut(ByVal addr As Integer) As Boolean
        On Error GoTo UsbEndpointDirectionOutErr
        If ((addr And &H80S) = 0) Then
            UsbEndpointDirectionOut = True
        Else
            UsbEndpointDirectionOut = False
        End If
        Exit Function
UsbEndpointDirectionOutErr:
        ProcessError(Err)
    End Function

    Public Function DeviceNameMatch(ByRef m As Message, ByVal DevicePathName As String) As Boolean
        Dim deviceNameString As String
        Dim stringSize As Integer
        deviceNameString = ""
        On Error GoTo DeviceNameMatchErr
        Dim devBroadcastDeviceInterface As New DEV_BROADCAST_DEVICEINTERFACE2
        Dim devBroadcastHeader As DEV_BROADCAST_HDR

        Marshal.PtrToStructure(m.LParam, devBroadcastHeader)
        If (devBroadcastHeader.dbch_devicetype = DBT_DEVTYP_DEVICEINTERFACE) Then
            Marshal.PtrToStructure(m.LParam, devBroadcastDeviceInterface)
            stringSize = CShort((devBroadcastHeader.dbch_size - 32) / 2) + 1
            deviceNameString = CStr(Left(devBroadcastDeviceInterface.dbcc_name, stringSize)) 'Store the device name in a String.
            If (StrComp(deviceNameString, DevicePathName, CompareMethod.Text) = 0) Then 'Set ignorecase True.
                DeviceNameMatch = True
            Else
                DeviceNameMatch = False
            End If
        End If
        Exit Function
DeviceNameMatchErr:
        ProcessError(Err)
    End Function

    Public Function FindDeviceFromGuid(ByRef TestGuid As Guid, ByRef DevicePathName As String, ByRef MemberIndex As Integer) As Boolean
        Dim bufferSize As Int32
        Dim detailDataBuffer As IntPtr
        Dim bufferSizeReq As Int32
        Dim deviceFound As Boolean
        Dim DeviceInfoSet As IntPtr
        Dim lastDevice As Boolean
        Dim TestDeviceInterfaceData As SP_DEVICE_INTERFACE_DATA
        Dim success As Boolean
        Dim pdevicePathName As IntPtr

        On Error GoTo FindDeviceFromGuidErr
        DeviceInfoSet = SetupDiGetClassDevsW(TestGuid, vbNullString, 0, DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE)
        deviceFound = False
        TestDeviceInterfaceData.cbSize = Marshal.SizeOf(TestDeviceInterfaceData) 'The size is 28 bytes.
        'confirmed to 28
        success = SetupDiEnumDeviceInterfaces(DeviceInfoSet, 0, TestGuid, MemberIndex, TestDeviceInterfaceData)
        If (Not success) Then 'Find out if a device information set was retrieved.
            lastDevice = True
        Else    'A device is present.
            success = SetupDiGetDeviceInterfaceDetailW(DeviceInfoSet, TestDeviceInterfaceData, 0, 0, bufferSize, 0)
            detailDataBuffer = Marshal.AllocHGlobal(bufferSize)
            Marshal.WriteInt32(detailDataBuffer, 4 + Marshal.SystemDefaultCharSize)
            success = SetupDiGetDeviceInterfaceDetailW(DeviceInfoSet, TestDeviceInterfaceData, detailDataBuffer, bufferSize, bufferSize, IntPtr.Zero)
            pdevicePathName = New IntPtr(detailDataBuffer.ToInt32 + 4)
            DevicePathName = Marshal.PtrToStringAuto(pdevicePathName)
            Marshal.FreeHGlobal(detailDataBuffer)
            deviceFound = True
        End If
        SetupDiDestroyDeviceInfoList(DeviceInfoSet)
        FindDeviceFromGuid = deviceFound
        Exit Function
FindDeviceFromGuidErr:
        MessageBox.Show("FindDeviceFromGuid Error end")
        ProcessError(Err)
    End Function

    Public Function ProcessError(ByRef errError As ErrObject, Optional ByRef errSource As String = "", Optional ByRef errControl As ERRORLEVEL = ERRORLEVEL.elevSrcDescReturn) As Object
        Select Case errControl
            Case ERRORLEVEL.elevCodeReturn
                ProcessError = Err.Number
            Case ERRORLEVEL.elevDescReturn
                ProcessError = Err.Description
            Case ERRORLEVEL.elevSrcDescReturn
                ProcessError = errSource & " : " & Err.Description
            Case ERRORLEVEL.elevMsgBox
                ProcessError = errSource & " : " & Err.Description
                MsgBox(Err.Description, MsgBoxStyle.Critical, errSource)
            Case Else
                ProcessError = errSource & " : " & Err.Description
        End Select
        'uncomment this line to display error in a message box
        MessageBox.Show("ProcessError: " & Err.Description & " " & errError.LastDllError)
    End Function

    'Purpose: Returns the error message for the last error.
    'Adapted from Dan Appleman's "Win32 API Puzzle Book"
    Public Function GetErrorString(ByVal LastError As Integer) As String
        Dim Bytes As Integer
        Dim ErrorString As String = ""

        GetErrorString = ""
        ErrorString = New String(Chr(0), 129)
        Bytes = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, LastError, 0, ErrorString, 128, 0)
        'Subtract two characters from the message to strip the CR and LF.
        If Bytes > 2 Then
            GetErrorString = Left(ErrorString, Bytes - 2)
        End If
    End Function

    Public Function CountDP5WinusbDevices() As Integer
        Dim hKey As Integer
        Dim retCode As Integer
        Dim lRet As Integer
        Dim idxDP5 As Integer
        Dim DevicePath(MAXREGBUFFER - 1) As Byte
        Dim cbDevicePath As Integer
        Dim KeyName As String = ""
        Dim cbKeyName As Integer
        Dim ErrMsg As String
        Dim cbErrMsg As Integer
        Dim strDevicePath As String = ""

        NumUSBDevices = 0
        CountDP5WinusbDevices = 0
        retCode = RegOpenKeyEx(HKEY_LOCAL_MACHINE, WinUSBService, 0, KEY_QUERY_VALUE, hKey)
        If (retCode <> ERROR_SUCCESS) Then
            CountDP5WinusbDevices = 0
            Exit Function
        End If
        KeyName = Space(MAX_PATH)
        ' Test ALL Keys (0,1,... are device paths, Count,NextInstance,(Default) have other info)
        For idxDP5 = 0 To (MAXDP5S + 3) - 1 'devs + 3 other keys
            cbKeyName = MAX_PATH
            cbDevicePath = MAXREGBUFFER
            retCode = RegEnumValue(hKey, idxDP5, KeyName, cbKeyName, 0, REG_NONE, DevicePath(0), cbDevicePath)
            If (retCode = ERROR_SUCCESS) Then
                strDevicePath = ByteArrayToString(DevicePath)
                If (StrComp(Left(KeyName, 5), "Count", CompareMethod.Text) = 0) Then
                    'do nothing
                ElseIf (StrComp(Left(strDevicePath, WinUsbDP5Size), WinUsbDP5, CompareMethod.Text) = 0) Then  ' DP5 device path found
                    'TRACE("DP5 device [%d]: %s=%s\r\n", (NumUSBDevices + 1), KeyName, DevicePath);
                    NumUSBDevices = NumUSBDevices + 1
                End If
            ElseIf (retCode = ERROR_NO_MORE_ITEMS) Then  ' no more values to read
                Exit For
            Else ' error reading values
                cbErrMsg = MAXERRBUFFER
                ErrMsg = Space(MAXERRBUFFER)
                lRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, retCode, 0, ErrMsg, cbErrMsg, 0)
                If (lRet > 0) Then ErrMsg = Left(ErrMsg, lRet) Else ErrMsg = ""
                'MsgBox ErrMsg
                Exit For
            End If
        Next
        RegCloseKey(hKey)
        CountDP5WinusbDevices = NumUSBDevices '/* return number of devices */
    End Function

    Public Function ByteArrayToString(ByRef byteArray() As Byte) As String
        Dim str_Renamed As String = ""
        Dim enc1252 As Encoding = CodePagesEncodingProvider.Instance.GetEncoding(1252)
        str_Renamed = enc1252.GetString(byteArray)
        ByteArrayToString = str_Renamed
    End Function

    Public Function WinUsbDeviceToString(ByRef DppWinUSB As WinUsbDevice) As String
        Dim strStr As String = ""
        On Error GoTo WinUsbDeviceToStringErr
        WinUsbDeviceToString = ""

        strStr += DppWinUSB.DeviceHandle.GetHashCode().ToString() & " DeviceHandle - hash code" & vbCrLf
        strStr += DppWinUSB.DeviceHandle.ToString & " DeviceHandle - Communications device resource handle." & vbCrLf
        strStr += DppWinUSB.winUsbHandle.ToString & " winUsbHandle - WinUSB interface handle." & vbCrLf
        strStr += DppWinUSB.bulkInPipe.ToString & " bulkInPipe - Bulk IN pipe id." & vbCrLf
        strStr += DppWinUSB.bulkOutPipe.ToString & " bulkOutPipe - Bulk OUT pipe id." & vbCrLf
        strStr += DppWinUSB.interruptInPipe.ToString & " interruptInPipe - Interrupt IN pipe id." & vbCrLf
        strStr += DppWinUSB.interruptOutPipe.ToString & " interruptOutPipe - Interrupt OUT pipe id." & vbCrLf
        strStr += DppWinUSB.devicespeed.ToString & " devicespeed - Device speed indicator." & vbCrLf
        WinUsbDeviceToString = strStr
        Exit Function
WinUsbDeviceToStringErr:
        WinUsbDeviceToString = "WinUsbDeviceToString Error"
    End Function

End Module
