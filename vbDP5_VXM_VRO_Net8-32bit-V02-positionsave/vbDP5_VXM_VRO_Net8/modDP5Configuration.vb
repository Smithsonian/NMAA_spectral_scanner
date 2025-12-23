Option Strict Off
Option Explicit On
Module modDP5Configuration

    Public s As New CSCommon
	
    Public Function GetCmdDesc(ByVal strCmd As String) As String
        Dim strCmdName As String = ""
        Select Case strCmd
            Case "RESC"
                strCmdName = "Reset Configuration"
            Case "CLCK"
                strCmdName = "20MHz/80MHz"
            Case "TPEA"
                strCmdName = "Peaking Time"
            Case "GAIF"
                strCmdName = "Fine Gain"
            Case "GAIN"
                strCmdName = "Total Gain (Analog * Fine)"
            Case "RESL"
                strCmdName = "Detector Reset Lockout"
            Case "TFLA"
                strCmdName = "Flat Top"
            Case "TPFA"
                strCmdName = "Fast Channel Peaking Time"
            Case "RTDE"
                strCmdName = "RTD On/Off"
            Case "MCAS"
                strCmdName = "MCA Source"
            Case "RTDD"
                strCmdName = "Custom RTD Oneshot Delay"
            Case "RTDW"
                strCmdName = "Custom RTD Oneshot Width"
            Case "PURE"
                strCmdName = "PUR Interval On/Off"
            Case "SOFF"
                strCmdName = "Set Spectrum Offset"
            Case "INOF"
                strCmdName = "Input Offset"
            Case "ACKE"
                strCmdName = "ACK / Don't ACK Packets With Errors"
            Case "AINP"
                strCmdName = "Analog Input Pos/Neg"
            Case "AUO1"
                strCmdName = "AUX_OUT Selection"
            Case "AUO2"
                strCmdName = "AUX_OUT2 Selection"
            Case "BLRD"
                strCmdName = "BLR Down Correction"
            Case "BLRM"
                strCmdName = "BLR Mode"
            Case "BLRU"
                strCmdName = "BLR Up Correction"
            Case "BOOT"
                strCmdName = "Turn Supplies On/Off At Power Up"
            Case "CUSP"
                strCmdName = "Non-Trapezoidal Shaping"
            Case "DACF"
                strCmdName = "DAC Offset"
            Case "DACO"
                strCmdName = "DAC Output"
            Case "GAIA"
                strCmdName = "Analog Gain Index"
            Case "GATE"
                strCmdName = "Gate Control"
            Case "GPED"
                strCmdName = "G.P. Counter Edge"
            Case "GPGA"
                strCmdName = "G.P. Counter Uses GATE?"
            Case "GPIN"
                strCmdName = "G.P. Counter Input"
            Case "GPMC"
                strCmdName = "G.P. Counter Cleared With MCA Counters?"
            Case "GPME"
                strCmdName = "G.P. Counter Uses MCA_EN?"
            Case "HVSE"
                strCmdName = "HV Set"
            Case "MCAC"
                strCmdName = "MCA/MCS Channels"
            Case "MCAE"
                strCmdName = "MCA/MCS Enable"
            Case "MCSL"
                strCmdName = "MCS Low Threshold"
            Case "MCSH"
                strCmdName = "MCS High Threshold"
            Case "MCST"
                strCmdName = "MCS Timebase"
            Case "PAPS"
                strCmdName = "Preamp 8.5/5 (N/A)"
            Case "PAPZ"
                strCmdName = "Pole-Zero"
            Case "PDMD"
                strCmdName = "Peak Detect Mode (Min/Max)"
            Case "PRCL"
                strCmdName = "Preset Counts Low Threshold"
            Case "PRCH"
                strCmdName = "Preset Counts High Threshold"
            Case "PREC"
                strCmdName = "Preset Counts"
            Case "PRER"
                strCmdName = "Preset Real Time"
            Case "PREL"
                strCmdName = "Preset Live Time"
            Case "PRET"
                strCmdName = "Preset Time"
            Case "RTDS"
                strCmdName = "RTD Sensitivity"
            Case "RTDT"
                strCmdName = "RTD Threshold"
            Case "SCAH"
                strCmdName = "SCAx High Threshold"
            Case "SCAI"
                strCmdName = "SCA Index"
            Case "SCAL"
                strCmdName = "SCAx Low Theshold"
            Case "SCAO"
                strCmdName = "SCAx Output (SCA1-8 Only)"
            Case "SCAW"
                strCmdName = "SCA Pulse Width (Not Indexed - SCA1-8)"
            Case "SCOE"
                strCmdName = "Scope Trigger Edge"
            Case "SCOG"
                strCmdName = "Digital Scope Gain"
            Case "SCOT"
                strCmdName = "Scope Trigger Position"
            Case "TECS"
                strCmdName = "TEC Set"
            Case "THFA"
                strCmdName = "Fast Threshold"
            Case "THSL"
                strCmdName = "Slow Threshold"
            Case "TLLD"
                strCmdName = "LLD Threshold"
            Case "TPMO"
                strCmdName = "Test Pulser On/Off"
            Case "VOLU"
                strCmdName = "Speaker On/Off"
            Case "CON1"
                strCmdName = "Connector 1"
            Case "CON2"
                strCmdName = "Connector 2"
        End Select
        GetCmdDesc = strCmdName
    End Function

    Public Function CreateFullReadBackCmd() As String
        Dim strCfg As String
        Dim isHVSE As Boolean
        Dim isPAPS As Boolean
        Dim isTECS As Boolean
        Dim isVOLU As Boolean
        Dim isCON1 As Boolean
        Dim isCON2 As Boolean
        Dim isINOF As Boolean
        Dim isBOOT As Boolean
        Dim isGATE As Boolean
        Dim isPAPZ As Boolean
        Dim isSCTC As Boolean
        Dim isDP5_DxK As Boolean
        Dim isDP5_DxL As Boolean
        
        If (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppMCA8000D) Then
            strCfg = CreateFullReadBackCmdMCA8000D()
            CreateFullReadBackCmd = strCfg
            Exit Function
        End If
        
        ' DP5 Rev Dx K,L needs PAPZ
        If ((STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5) And STATUS.isDP5_RevDxGains) Then
            If ((STATUS.DPP_ECO And &HF) = &HA) Then
                isDP5_DxK = True
            End If
            If ((STATUS.DPP_ECO And &HF) = &HB) Then
                isDP5_DxL = True
            End If
        End If

        isHVSE = CBool(((STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppPX5) And STATUS.PC5_PRESENT) Or (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppPX5))
        isPAPS = CBool((STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppDP5G) And (STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppTB5))
        isTECS = CBool(((STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5) And STATUS.PC5_PRESENT) Or (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppPX5) Or (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5X))
        isVOLU = CBool(STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppPX5)
        
        isCON1 = CBool((STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppDP5) And (STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppDP5X))
        isCON2 = CBool((STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppDP5) And (STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppDP5X))
        
        isINOF = CBool((STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppDP5G) And (STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppTB5))
        isSCTC = CBool((STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5G) Or (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppTB5))
        isBOOT = CBool((STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5) Or (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5X))
        
        isGATE = CBool((STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5) Or (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5X))
        isPAPZ = CBool((STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppPX5) Or isDP5_DxK Or isDP5_DxL)
        
        strCfg = ""
        strCfg = strCfg & "RESC=?;"
        strCfg = strCfg & "CLCK=?;"
        strCfg = strCfg & "TPEA=?;"
        strCfg = strCfg & "GAIF=?;"
        strCfg = strCfg & "GAIN=?;"
        strCfg = strCfg & "RESL=?;"
        strCfg = strCfg & "TFLA=?;"
        strCfg = strCfg & "TPFA=?;"
        strCfg = strCfg & "PURE=?;"
        If (isSCTC) Then strCfg = strCfg & "SCTC=?;"
        strCfg = strCfg & "RTDE=?;"
        strCfg = strCfg & "MCAS=?;"
        strCfg = strCfg & "MCAC=?;"
        strCfg = strCfg & "SOFF=?;"
        strCfg = strCfg & "AINP=?;"
        If (isINOF) Then strCfg = strCfg & "INOF=?;"
        strCfg = strCfg & "GAIA=?;"
        strCfg = strCfg & "CUSP=?;"
        strCfg = strCfg & "PDMD=?;"
        strCfg = strCfg & "THSL=?;"
        strCfg = strCfg & "TLLD=?;"
        strCfg = strCfg & "THFA=?;"
        strCfg = strCfg & "DACO=?;"
        strCfg = strCfg & "DACF=?;"
        strCfg = strCfg & "RTDS=?;"
        strCfg = strCfg & "RTDT=?;"
        strCfg = strCfg & "BLRM=?;"
        strCfg = strCfg & "BLRD=?;"
        strCfg = strCfg & "BLRU=?;"
        If (isGATE) Then strCfg = strCfg & "GATE=?;"
        strCfg = strCfg & "AUO1=?;"
        strCfg = strCfg & "PRET=?;"
        strCfg = strCfg & "PRER=?;"
        strCfg = strCfg & "PREC=?;"
        strCfg = strCfg & "PRCL=?;"
        strCfg = strCfg & "PRCH=?;"
        If (isHVSE) Then strCfg = strCfg & "HVSE=?;"
        If (isTECS) Then strCfg = strCfg & "TECS=?;"
        If (isPAPZ) Then strCfg = strCfg & "PAPZ=?;"
        If (isPAPS) Then strCfg = strCfg & "PAPS=?;"
        strCfg = strCfg & "SCOE=?;"
        strCfg = strCfg & "SCOT=?;"
        strCfg = strCfg & "SCOG=?;"
        strCfg = strCfg & "MCSL=?;"
        strCfg = strCfg & "MCSH=?;"
        strCfg = strCfg & "MCST=?;"
        strCfg = strCfg & "AUO2=?;"
        strCfg = strCfg & "TPMO=?;"
        strCfg = strCfg & "GPED=?;"
        strCfg = strCfg & "GPIN=?;"
        strCfg = strCfg & "GPME=?;"
        strCfg = strCfg & "GPGA=?;"
        strCfg = strCfg & "GPMC=?;"
        strCfg = strCfg & "MCAE=?;"
        If (isVOLU) Then strCfg = strCfg & "VOLU=?;"
        If (isCON1) Then strCfg = strCfg & "CON1=?;"
        If (isCON2) Then strCfg = strCfg & "CON2=?;"
        If (isBOOT) Then strCfg = strCfg & "BOOT=?;"
        CreateFullReadBackCmd = strCfg
    End Function

    Public Function CreateFullReadBackCmdMCA8000D() As String
        Dim strCfg As String
        strCfg = ""
        If (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppMCA8000D) Then
            strCfg = strCfg & "RESC=?;"
            strCfg = strCfg & "PURE=?;"
            strCfg = strCfg & "MCAS=?;"
            strCfg = strCfg & "MCAC=?;"
            strCfg = strCfg & "SOFF=?;"
            strCfg = strCfg & "GAIA=?;"
            strCfg = strCfg & "PDMD=?;"
            strCfg = strCfg & "THSL=?;"
            strCfg = strCfg & "TLLD=?;"
            strCfg = strCfg & "GATE=?;"
            strCfg = strCfg & "AUO1=?;"
            strCfg = strCfg & "PRER=?;"
            strCfg = strCfg & "PREL=?;"
            strCfg = strCfg & "PREC=?;"
            strCfg = strCfg & "PRCL=?;"
            strCfg = strCfg & "PRCH=?;"
            strCfg = strCfg & "SCOE=?;"
            strCfg = strCfg & "SCOT=?;"
            strCfg = strCfg & "SCOG=?;"
            strCfg = strCfg & "MCSL=?;"
            strCfg = strCfg & "MCSH=?;"
            strCfg = strCfg & "MCST=?;"
            strCfg = strCfg & "AUO2=?;"
            strCfg = strCfg & "GPED=?;"
            strCfg = strCfg & "GPIN=?;"
            strCfg = strCfg & "GPME=?;"
            strCfg = strCfg & "GPGA=?;"
            strCfg = strCfg & "GPMC=?;"
            strCfg = strCfg & "MCAE=?;"
            strCfg = strCfg & "PDMD=?;"
        End If
        CreateFullReadBackCmdMCA8000D = strCfg
    End Function

    Public Sub MakeDp5CmdList(ByRef strCfgArr As Collection)
        Dim isHVSE As Boolean
        Dim isPAPS As Boolean
        Dim isTECS As Boolean
        Dim isVOLU As Boolean
        Dim isCON1 As Boolean
        Dim isCON2 As Boolean
        Dim isINOF As Boolean
        Dim isBOOT As Boolean
        Dim isGATE As Boolean
        Dim isPAPZ As Boolean
        Dim isSCTC As Boolean
        Dim isDP5_DxK As Boolean
        Dim isDP5_DxL As Boolean

        If (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppMCA8000D) Then
            Call MakeDp5CmdListMCA8000D(strCfgArr)
            Exit Sub
        End If

        ' DP5 Rev Dx K,L needs PAPZ
        If ((STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5) And STATUS.isDP5_RevDxGains) Then
            If ((STATUS.DPP_ECO And &HF) = &HA) Then
                isDP5_DxK = True
            End If
            If ((STATUS.DPP_ECO And &HF) = &HB) Then
                isDP5_DxL = True
            End If
        End If

        isHVSE = CBool(((STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppPX5) And STATUS.PC5_PRESENT) Or (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppPX5))
        isPAPS = CBool((STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppDP5G) And (STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppTB5))
        isTECS = CBool(((STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5) And STATUS.PC5_PRESENT) Or (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppPX5) Or (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5X))
        isVOLU = CBool(STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppPX5)

        isCON1 = CBool((STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppDP5) And (STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppDP5X))
        isCON2 = CBool((STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppDP5) And (STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppDP5X))

        isINOF = CBool((STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppDP5G) And (STATUS.DEVICE_ID <> modParsePacket.DP5_DPP_TYPES.dppTB5))
        isSCTC = CBool((STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5G) Or (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppTB5))
        isBOOT = CBool((STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5) Or (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5X))

        isGATE = CBool((STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5) Or (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppDP5X))
        isPAPZ = CBool((STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppPX5) Or isDP5_DxK Or isDP5_DxL)

        strCfgArr.Add("RESC")
        strCfgArr.Add("CLCK")
        strCfgArr.Add("TPEA")
        strCfgArr.Add("GAIF")
        strCfgArr.Add("GAIN")
        strCfgArr.Add("RESL")
        strCfgArr.Add("TFLA")
        strCfgArr.Add("TPFA")
        strCfgArr.Add("PURE")
        If (isSCTC) Then strCfgArr.Add("SCTC")
        strCfgArr.Add("RTDE")
        strCfgArr.Add("MCAS")
        strCfgArr.Add("MCAC")
        strCfgArr.Add("SOFF")
        strCfgArr.Add("AINP")
        If (isINOF) Then strCfgArr.Add("INOF")
        strCfgArr.Add("GAIA")
        strCfgArr.Add("CUSP")
        strCfgArr.Add("PDMD")
        strCfgArr.Add("THSL")
        strCfgArr.Add("TLLD")
        strCfgArr.Add("THFA")
        strCfgArr.Add("DACO")
        strCfgArr.Add("DACF")
        strCfgArr.Add("RTDS")
        strCfgArr.Add("RTDT")
        strCfgArr.Add("BLRM")
        strCfgArr.Add("BLRD")
        strCfgArr.Add("BLRU")
        If (isGATE) Then strCfgArr.Add("GATE")
        strCfgArr.Add("AUO1")
        strCfgArr.Add("PRET")
        strCfgArr.Add("PRER")
        strCfgArr.Add("PREC")
        strCfgArr.Add("PRCL")
        strCfgArr.Add("PRCH")
        If (isHVSE) Then strCfgArr.Add("HVSE")
        If (isTECS) Then strCfgArr.Add("TECS")
        If (isPAPZ) Then strCfgArr.Add("PAPZ")
        If (isPAPS) Then strCfgArr.Add("PAPS")
        strCfgArr.Add("SCOE")
        strCfgArr.Add("SCOT")
        strCfgArr.Add("SCOG")
        strCfgArr.Add("MCSL")
        strCfgArr.Add("MCSH")
        strCfgArr.Add("MCST")
        strCfgArr.Add("AUO2")
        strCfgArr.Add("TPMO")
        strCfgArr.Add("GPED")
        strCfgArr.Add("GPIN")
        strCfgArr.Add("GPME")
        strCfgArr.Add("GPGA")
        strCfgArr.Add("GPMC")
        strCfgArr.Add("MCAE")
        If (isVOLU) Then strCfgArr.Add("VOLU")
        If (isCON1) Then strCfgArr.Add("CON1")
        If (isCON2) Then strCfgArr.Add("CON2")
        If (isBOOT) Then strCfgArr.Add("BOOT")
    End Sub

    Public Sub MakeDp5CmdListMCA8000D(ByRef strCfgArr As Collection)
        If (STATUS.DEVICE_ID = modParsePacket.DP5_DPP_TYPES.dppMCA8000D) Then
            strCfgArr.Add("RESC")
            strCfgArr.Add("PURE")
            strCfgArr.Add("MCAS")
            strCfgArr.Add("MCAC")
            strCfgArr.Add("SOFF")
            strCfgArr.Add("GAIA")
            strCfgArr.Add("PDMD")
            strCfgArr.Add("THSL")
            strCfgArr.Add("TLLD")
            strCfgArr.Add("GATE")
            strCfgArr.Add("AUO1")
            strCfgArr.Add("PRER")
            strCfgArr.Add("PREL")
            strCfgArr.Add("PREC")
            strCfgArr.Add("PRCL")
            strCfgArr.Add("PRCH")
            strCfgArr.Add("SCOE")
            strCfgArr.Add("SCOT")
            strCfgArr.Add("SCOG")
            strCfgArr.Add("MCSL")
            strCfgArr.Add("MCSH")
            strCfgArr.Add("MCST")
            strCfgArr.Add("AUO2")
            strCfgArr.Add("GPED")
            strCfgArr.Add("GPIN")
            strCfgArr.Add("GPME")
            strCfgArr.Add("GPGA")
            strCfgArr.Add("GPMC")
            strCfgArr.Add("MCAE")
            strCfgArr.Add("PDMD")
        End If
    End Sub

    Public Function GetCmdData(ByVal cstrCmd As String, ByVal cstrCfgData As String) As String
        Dim iStart As Integer
        Dim iEnd As Integer
        Dim iCmd As Integer
        Dim cstrCmdData As String = ""

        GetCmdData = ""
        cstrCmdData = ""
        If (Len(cstrCfgData) < 7) Then Exit Function ' no data
        If (Len(cstrCmd) <> 4) Then Exit Function ' bad command
        iCmd = InStr(1, cstrCfgData, cstrCmd + "=", CompareMethod.Text)
        If (iCmd = 0) Then Exit Function ' cmd not found
        iStart = InStr(iCmd, cstrCfgData, "=")
        If (iStart = 0) Then Exit Function ' data start not found
        iEnd = InStr(iCmd, cstrCfgData, ";")
        If (iEnd = 0) Then GetCmdData = cstrCmdData ' data end found
        If (iStart >= iEnd) Then Exit Function ' data error
        cstrCmdData = Mid(cstrCfgData, iStart + 1, iEnd - (iStart + 1))
        GetCmdData = cstrCmdData
    End Function

    Public Function ReplaceCmdDesc(ByVal cstrCmd As String, ByVal cstrCfgData As String) As String
        Dim iStart As Integer
        Dim iCmd As Integer
        Dim cstrNew As String
        Dim cstrDesc As String

        ReplaceCmdDesc = cstrCfgData
        cstrNew = ""
        If (Len(cstrCfgData) < 7) Then Exit Function ' no data
        If (Len(cstrCmd) <> 4) Then Exit Function ' bad command
        iCmd = InStr(1, cstrCfgData, cstrCmd + "=")
        If (iCmd = 0) Then Exit Function ' cmd not found
        cstrDesc = GetCmdDesc(cstrCmd)
        If (Len(cstrDesc) = 0) Then Exit Function ' cmd desc  not found
        iStart = InStr(iCmd, cstrCfgData, "=")
        If (iStart <> (iCmd + 4)) Then Exit Function ' data start not found
        cstrNew = Left(cstrCfgData, iCmd - 1) & cstrDesc & Mid(cstrCfgData, iStart)
        ReplaceCmdDesc = cstrNew
    End Function

    Public Function AppendCmdDesc(ByVal cstrCmd As String, ByVal cstrCfgData As String) As String
        Dim iStart As Integer
        Dim iEnd As Integer
        Dim iCmd As Integer
        Dim cstrNew As String
        Dim cstrDesc As String

        AppendCmdDesc = cstrCfgData
        cstrNew = ""
        If (Len(cstrCfgData) < 7) Then Exit Function ' no data
        If (Len(cstrCmd) <> 4) Then Exit Function ' bad command
        iCmd = InStr(1, cstrCfgData, cstrCmd + "=")
        If (iCmd = 0) Then Exit Function ' cmd not found
        cstrDesc = GetCmdDesc(cstrCmd)
        If (Len(cstrDesc) = 0) Then Exit Function ' cmd desc  not found
        iStart = InStr(iCmd, cstrCfgData, "=")
        If (iStart <> (iCmd + 4)) Then Exit Function ' data start not found
        iEnd = InStr(iStart + 1, cstrCfgData, ";")
        If (iEnd > (iStart + 11)) Then Exit Function ' data end not found
        cstrNew = Left(cstrCfgData, iEnd) & "    " & cstrDesc & Mid(cstrCfgData, iEnd + 1)
        AppendCmdDesc = cstrNew
    End Function

    Public Function RemoveCmd(ByVal cstrCmd As String, ByVal cstrCfgData As String) As String
        Dim iStart As Integer
        Dim iEnd As Integer
        Dim iCmd As Integer
        Dim cstrNew As String
        Dim strLeft As String
        Dim strRight As String

        cstrNew = ""
        RemoveCmd = cstrCfgData
        If (Len(cstrCfgData) < 7) Then Exit Function ' no data
        If (Len(cstrCmd) <> 4) Then Exit Function ' bad command
        iCmd = InStr(1, cstrCfgData, cstrCmd + "=")
        If (iCmd = 0) Then Exit Function ' cmd not found
        iStart = iCmd
        iEnd = InStr(iCmd, cstrCfgData, ";")
        If (iEnd = 0) Then Exit Function
        If (iEnd <= iStart) Then Exit Function ' unknown error
        strLeft = Left(cstrCfgData, iStart - 1)
        strRight = Mid(cstrCfgData, iEnd + 1)
        cstrNew = strLeft & strRight
        RemoveCmd = cstrNew
    End Function

    'removes selected command by DPP device type
    Public Function RemoveCmdByDeviceType(ByVal strCfgDataIn As String, ByVal PC5_PRESENT As Boolean, ByVal DppType As Byte, ByVal isDP5_RevDxGains As Boolean, ByVal DPP_ECO As Byte) As String
        Dim strCfgData As String
        Dim isHVSE As Boolean
        Dim isPAPS As Boolean
        Dim isTECS As Boolean
        Dim isVOLU As Boolean
        Dim isCON1 As Boolean
        Dim isCON2 As Boolean
        Dim isINOF As Boolean
        Dim isBOOT As Boolean
        Dim isGATE As Boolean
        Dim isPAPZ As Boolean
        Dim isSCTC As Boolean
        Dim isPREL As Boolean
        Dim isDP5_DxK As Boolean
        Dim isDP5_DxL As Boolean

        strCfgData = strCfgDataIn
        If (DppType = modParsePacket.DP5_DPP_TYPES.dppMCA8000D) Then
            strCfgData = Remove_MCA8000D_Cmds(strCfgData, DppType)
            RemoveCmdByDeviceType = strCfgData
        End If

        ' DP5 Rev Dx K,L needs PAPZ
        If ((DppType = modParsePacket.DP5_DPP_TYPES.dppDP5) And isDP5_RevDxGains) Then
            If ((DPP_ECO And &HF) = &HA) Then
                isDP5_DxK = True
            End If
            If ((DPP_ECO And &HF) = &HB) Then
                isDP5_DxL = True
            End If
        End If

        isHVSE = CBool(((DppType <> modParsePacket.DP5_DPP_TYPES.dppPX5) And PC5_PRESENT) Or (DppType = modParsePacket.DP5_DPP_TYPES.dppPX5))
        isPAPS = CBool((DppType <> modParsePacket.DP5_DPP_TYPES.dppDP5G) And (DppType <> modParsePacket.DP5_DPP_TYPES.dppTB5))
        isTECS = CBool(((DppType = modParsePacket.DP5_DPP_TYPES.dppDP5) And PC5_PRESENT) Or (DppType = modParsePacket.DP5_DPP_TYPES.dppPX5) Or (DppType = modParsePacket.DP5_DPP_TYPES.dppDP5X))
        isVOLU = CBool(DppType = modParsePacket.DP5_DPP_TYPES.dppPX5)
        isCON1 = CBool((DppType <> modParsePacket.DP5_DPP_TYPES.dppDP5) And (DppType <> modParsePacket.DP5_DPP_TYPES.dppDP5X))
        isCON2 = CBool((DppType <> modParsePacket.DP5_DPP_TYPES.dppDP5) And (DppType <> modParsePacket.DP5_DPP_TYPES.dppDP5X))
        isINOF = CBool((DppType <> modParsePacket.DP5_DPP_TYPES.dppDP5G) And (DppType <> modParsePacket.DP5_DPP_TYPES.dppTB5))
        isSCTC = CBool((DppType = modParsePacket.DP5_DPP_TYPES.dppDP5G) Or (DppType = modParsePacket.DP5_DPP_TYPES.dppTB5))
        isBOOT = CBool((DppType = modParsePacket.DP5_DPP_TYPES.dppDP5) Or (DppType = modParsePacket.DP5_DPP_TYPES.dppDP5X))
        isGATE = CBool((DppType = modParsePacket.DP5_DPP_TYPES.dppDP5) Or (DppType = modParsePacket.DP5_DPP_TYPES.dppDP5X))
        isPAPZ = CBool((DppType = modParsePacket.DP5_DPP_TYPES.dppPX5) Or isDP5_DxK Or isDP5_DxL)
        isPREL = CBool(DppType = modParsePacket.DP5_DPP_TYPES.dppMCA8000D)
        If (Not isHVSE) Then strCfgData = RemoveCmd("HVSE", strCfgData) 'High Voltage Bias
        If (Not isPAPS) Then strCfgData = RemoveCmd("PAPS", strCfgData) 'Preamp Voltage
        If (Not isTECS) Then strCfgData = RemoveCmd("TECS", strCfgData) 'Cooler Temperature
        If (Not isVOLU) Then strCfgData = RemoveCmd("VOLU", strCfgData) 'px5 speaker
        If (Not isCON1) Then strCfgData = RemoveCmd("CON1", strCfgData) 'connector 1
        If (Not isCON2) Then strCfgData = RemoveCmd("CON2", strCfgData) 'connector 2
        If (Not isINOF) Then strCfgData = RemoveCmd("INOF", strCfgData) 'input offset
        If (Not isBOOT) Then strCfgData = RemoveCmd("BOOT", strCfgData) 'PC5 On At StartUp
        If (Not isGATE) Then strCfgData = RemoveCmd("GATE", strCfgData) 'Gate input
        If (Not isPAPZ) Then strCfgData = RemoveCmd("PAPZ", strCfgData) 'Pole-Zero
        If (Not isSCTC) Then strCfgData = RemoveCmd("SCTC", strCfgData) 'Scintillator Time Constant
        If (Not isPREL) Then strCfgData = RemoveCmd("PREL", strCfgData) 'Preset Live Time
        RemoveCmdByDeviceType = strCfgData
    End Function

    ''removes MCA8000D commands
    Public Function Remove_MCA8000D_Cmds(ByVal strCfgDataIn As String, ByVal DppType As Byte) As String
        Dim strCfgData As String

        strCfgData = strCfgDataIn
        If (DppType = modParsePacket.DP5_DPP_TYPES.dppMCA8000D) Then
            strCfgData = RemoveCmd("CLCK", strCfgData)
            strCfgData = RemoveCmd("TPEA", strCfgData)
            strCfgData = RemoveCmd("GAIF", strCfgData)
            strCfgData = RemoveCmd("GAIN", strCfgData)
            strCfgData = RemoveCmd("RESL", strCfgData)
            strCfgData = RemoveCmd("TFLA", strCfgData)
            strCfgData = RemoveCmd("TPFA", strCfgData)
            'strCfgData = RemoveCmd("PURE", strCfgData)
            strCfgData = RemoveCmd("RTDE", strCfgData)
            strCfgData = RemoveCmd("AINP", strCfgData)
            strCfgData = RemoveCmd("INOF", strCfgData)
            strCfgData = RemoveCmd("CUSP", strCfgData)
            strCfgData = RemoveCmd("THFA", strCfgData)
            strCfgData = RemoveCmd("DACO", strCfgData)
            strCfgData = RemoveCmd("DACF", strCfgData)
            strCfgData = RemoveCmd("RTDS", strCfgData)
            strCfgData = RemoveCmd("RTDT", strCfgData)
            strCfgData = RemoveCmd("BLRM", strCfgData)
            strCfgData = RemoveCmd("BLRD", strCfgData)
            strCfgData = RemoveCmd("BLRU", strCfgData)
            strCfgData = RemoveCmd("PRET", strCfgData)
            strCfgData = RemoveCmd("HVSE", strCfgData)
            strCfgData = RemoveCmd("TECS", strCfgData)
            strCfgData = RemoveCmd("PAPZ", strCfgData)
            strCfgData = RemoveCmd("PAPS", strCfgData)
            strCfgData = RemoveCmd("TPMO", strCfgData)
            strCfgData = RemoveCmd("SCAH", strCfgData)
            strCfgData = RemoveCmd("SCAI", strCfgData)
            strCfgData = RemoveCmd("SCAL", strCfgData)
            strCfgData = RemoveCmd("SCAO", strCfgData)
            strCfgData = RemoveCmd("SCAW", strCfgData)
            strCfgData = RemoveCmd("BOOT", strCfgData)

            ' added to list late, recheck at later date 20120817
            strCfgData = RemoveCmd("CON1", strCfgData)
            strCfgData = RemoveCmd("CON2", strCfgData)

            ' not implemented as of 20120817, will be implemented at some time
            strCfgData = RemoveCmd("VOLU", strCfgData)
        End If
        Remove_MCA8000D_Cmds = strCfgData
    End Function

    Public Function GetCmdChunk(ByVal strCmd As String) As Long
        Dim idxCfg As Long
        Dim lChunk As Long
        Dim lEnd As Long
        GetCmdChunk = 0
        lChunk = 0
        lEnd = 0
        For idxCfg = 1 To Len(strCmd)
            lChunk = InStr(lEnd + 1, strCmd, ";", CompareMethod.Binary)
            If ((lChunk = 0) Or (lChunk > 512)) Then Exit For
            lEnd = lChunk
        Next
        GetCmdChunk = lEnd
    End Function

    Public Function ReplaceText(ByVal strInTextIn As String, ByVal strFrom As String, ByVal strTo As String) As String
        Dim strInText As String
        Dim strOutText As String
        Dim lFromLen As Long
        Dim lMatchPos As Long

        strInText = strInTextIn
        strOutText = ""
        lFromLen = Len(strFrom)
        Do While Len(strInText) > 0
            lMatchPos = InStr(strInText, strFrom)
            If lMatchPos = 0 Then
                strOutText = strOutText & strInText
                strInText = ""
            Else
                strOutText = strOutText & Left(strInText, lMatchPos - 1) & strTo
                strInText = Mid(strInText, lMatchPos + lFromLen)
            End If
        Loop
        ReplaceText = strOutText
    End Function

    '' ''   'See \DotNet\VB.Net\vbDP5_NET implementation
    '' ''    Private Sub cmdSendConfiguration_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSendConfiguration.Click
    '' ''        Dim idxCfg As Short
    '' ''        Dim varCfgHw As Object
    '' ''        Dim varCommentsHw As Object = Nothing
    '' ''        Dim strCfg As String = ""
    '' ''        Dim strDisplay As String = ""
    '' ''        Dim strCmd As String = ""
    '' ''        Dim bStatusDone As Boolean

    '' ''        'reduce size of configuration if necessary
    '' ''        Dim lCfgLen As Long                 'ASCII Configuration Command String Length
    '' ''        Dim bSplitCfg As Boolean            'Configuration split flag
    '' ''        Dim idxSplitCfg As Long             'Configuration split position, only if necessary
    '' ''        Dim strSplitCfg As String           'Configuration split string second buffer

    '' ''        lCfgLen = 0
    '' ''        bSplitCfg = False
    '' ''        idxSplitCfg = 0
    '' ''        strSplitCfg = ""

    '' ''        strDisplay = ""
    '' ''        On Error GoTo cmdSendConfigurationErr
    '' ''        If (dlgOpen(OpenFile, dlgTXT_Filter)) Then
    '' ''            strIniFilename = OpenFile.FileName
    '' ''            varCfgHw = GetDP5ConfigSection(strIniFilename, IniSectionCfg, varCommentsHw) 'read ini config from file
    '' ''            If (IsNothing(varCfgHw)) Then Exit Sub 'not initialized
    '' ''            strCfg = TypeName(varCfgHw) 'test variant data type
    '' ''            If (strCfg = "String(,)") Then 'have data
    '' ''                strCfg = "" 'clear cfg storage
    '' ''            Else
    '' ''                Exit Sub 'no data
    '' ''            End If
    '' ''            For idxCfg = 0 To UBound(varCfgHw, 1)
    '' ''                strCmd = varCfgHw(idxCfg, 0) & "=" & varCfgHw(idxCfg, 1) & ";"
    '' ''                strCfg = strCfg & strCmd
    '' ''                strDisplay = strDisplay & strCmd & vbCrLf
    '' ''            Next
    '' ''            s.HwCfgReady = False
    '' ''            s.HwCfgExReady = False
    '' ''            strCfg = RemoveCmdByDeviceType(strCfg, STATUS.PC5_PRESENT, STATUS.DEVICE_ID)

    '' ''            'Test configuration size
    '' ''            lCfgLen = Len(strCfg)
    '' ''            If (lCfgLen > 512) Then
    '' ''                strCfg = ReplaceText(strCfg, "US;", ";")
    '' ''                strCfg = ReplaceText(strCfg, "OFF;", "OF;")
    '' ''                strCfg = ReplaceText(strCfg, "RISING;", "RI;")
    '' ''                strCfg = ReplaceText(strCfg, "FALLING;", "FA;")
    '' ''                lCfgLen = Len(strCfg)
    '' ''                If (lCfgLen > 512) Then 'configuration is still too large, split cfg
    '' ''                    bSplitCfg = True
    '' ''                    idxSplitCfg = GetCmdChunk(strCfg)
    '' ''                    strSplitCfg = Mid(strCfg, idxSplitCfg + 1)
    '' ''                    strCfg = VB.Strings.Left(strCfg, idxSplitCfg)
    '' ''                End If
    '' ''            End If
    '' ''            If (Len(strCfg) > 0) Then
    '' ''                lblCfgLenValue.Text = CStr(Len(strCfg))
    '' ''                s.HwCfgDP5Out = strCfg
    '' ''                s.HwCfgReady = True
    '' ''                SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_CONFIG_PACKET_TO_HW)
    '' ''                s.HwCfgReady = False
    '' ''            End If
    '' ''            If (bSplitCfg) Then
    '' ''                bStatusDone = msDelay(200)  'give time for timer to finish
    '' ''                lblCfgLenValue.Text = CStr(Len(strSplitCfg))
    '' ''                s.HwCfgDP5Out = strSplitCfg
    '' ''                s.HwCfgReady = True
    '' ''                SendCommand(modDP5_Protocol.TRANSMIT_PACKET_TYPE.XMTPT_SEND_CONFIG_PACKET_TO_HW)
    '' ''                s.HwCfgReady = False
    '' ''            End If
    '' ''            'txtSendCfgToHwNoEdit.Text = strDisplay
    '' ''        Else
    '' ''            'a file was not selected
    '' ''        End If
    '' ''        Exit Sub
    '' ''cmdSendConfigurationErr:
    '' ''    End Sub
End Module
