Option Strict Off
Option Explicit On
Friend Class CSCommon

	Public HwCfgDP5 As String
	Public cstrRawCfgIn As String
	Public HwCfgReady As Boolean
	Public HwCfgExReady As Boolean
	Public Dp5CmdList As New Collection
	Public HwCfgDP5Out As String
	Public HwRdBkDP5Out As String

	Public SCAEnabled As Boolean
	Public ScaReadReady As Boolean
	Public ScaReadBack As Boolean
	Public DisplaySCA As Boolean 'parse/load sca settings into SCA controls

	Public HwScaCfgDP5 As String
	Public cstrRawScaCfgIn As String
	Public HwScaCfgReady As Boolean
	
	' Shared Communications Interface Settings
	Public CurrentInterface As Byte ' selected communications interface (0=RS232/1=USB/2=ETHERNET)
	Public isDppConnected As Boolean ' CurrentInterface selected and device found
	Public ComPort As Integer ' Serial Port Settings
	
	' Ethernet Settings
	Public NetfinderActive As Boolean
	'Public SockAddr As UInt32
	Public SockAddr As Long
	Public cstrSockAddr As String
	Public InetPort As Integer
	
	Public CfgReadBack As Boolean 'load dpp config into config editor
	Public DisplayCfg As Boolean 'show dpp config in msgbox dialog
	Public SaveCfg As Boolean 'save dpp config to cfg file
	Public SpectrumCfg As Boolean 'create dpp config for mca file
	
	Public SendCoarseFineGain As Boolean
End Class
