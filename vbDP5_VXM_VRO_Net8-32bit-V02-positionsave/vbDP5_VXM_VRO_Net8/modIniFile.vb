Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Runtime.InteropServices

Module modIniFile

	'Purpose: <BR>
	'   INI File Utilities for DP5 Configurations.
	'
	'Example:
	'
	'Code:'==========================================================================
	'Code:'Quick DP5 INI Utility Guide
	'Code:'==========================================================================
	'Code:'--------------------------------------------------------------------------
	'Code:' All INI data arrays are string arrays and have the form:
	'Code:'	 varData(Records,Fields)
	'Code:' Where:
	'Code:'	 Records is the Number of INI Section Entries
	'Code:'	 Fields values are 0=INI Key, 1=INI Data
	'Code:'--------------------------------------------------------------------------
	'Code:' Note: varSectionNames is a string array of section names. (NOT INI DATA)
	'Code:'--------------------------------------------------------------------------
	'Code:'--------------------------------------------------------------------------
	'Code:'
	'Code:'----- HOW DO I ? -----
	'Code:'
	'Code:'--------------------------------------------------------------------------
	'Code:'Read a configuration.
	'Code:	varConfig = GetDP5ConfigSection(strFilename, strSection, varComments)
	'Code:'--------------------------------------------------------------------------
	'Code:'Save a configuration.
	'Code:	Call SaveDP5ConfigSection(strFilename, strSection, varConfig, varComments)
	'Code:'--------------------------------------------------------------------------
	'Code:'Save ANY INI data array.
	'Code:	Call SaveIniDataArray(strFilename, strSection, varData)
	'Code:'--------------------------------------------------------------------------
	'Code:'Appends semicolons to configuration data.
	'Code:	varConfig = AppendSemicolonsToConfig(varConfig)
	'Code:'--------------------------------------------------------------------------
	'Code:'Set all keys and config data to upper case.
	'Code:	varConfig = ConfigToUCASE(varConfig)
	'Code:'--------------------------------------------------------------------------
	'Code:'Append comments array to configuration data.
	'Code:	varData = AppendDP5Comments(varConfig, varComments)
	'Code:'--------------------------------------------------------------------------
	'Code:'Search for data in ANY INI array by key.
	'Code:	strData = FindIniData(strKey, varData)
	'Code:'--------------------------------------------------------------------------
	'Code:'Convert ANY INI data array to a formatted text display string.
	'Code:	strDisplaySection = DataArrayString(strSection, varData)
	'Code:'--------------------------------------------------------------------------
	'Code:'Convert ANY INI data file (all sections) to a formatted text display string.
	'Code:	strDisplayIniFile = IniFileString(strFilename)
	'Code:'--------------------------------------------------------------------------
	'Code:'Read all INI file section names.
	'Code:	varSectionNames = GetSectionNames(strFilename)
	'Code:'--------------------------------------------------------------------------
	'Code:'Extract DP5 commands from a INI data array.
	'Code:	varData = GetDP5Commands(varData)
	'Code:'--------------------------------------------------------------------------
	'Code:'Extract DP5 comments from a INI data array.
	'Code:	varComments = GetDP5Comments(varData)
	'Code:'--------------------------------------------------------------------------
	'Code:'Create a copy of an INI array. Copies keys and data.
	'Code:	varDataNew = CreateIniArr(varData) (NOTE: bCopyData = True by default)
	'Code:'--------------------------------------------------------------------------
	'Code:'Create an empty INI array of varData size.
	'Code:	varDataNew = CreateIniArr(varData, False)
	'Code:'--------------------------------------------------------------------------
	'Code:'Display an INI file section to a list control.
	'Code:	Call GetIniListEx2(strFilename, strSection, objList, bRemComments)
	'Code:'--------------------------------------------------------------------------

	Public MAX_SECTION_SIZE As Integer = &H7FFF
	Public MAX_KEY_DATA_SIZE As Integer = &HF

	Public varConfig As Object
	Public varComments As Object
	Public varValues As Object
	Public varValComments As Object
	
	Public varCmdCtrls As Object
	Public varUnitsArray As Object
	
	Public Const IniSectionVal As String = "DP5 Configuration Values"
	Public Const IniSectionCfg As String = "DP5 Configuration File"
	Public Const IniSectionApp As String = "vbDP5 Application Settings"
	
	Public strIniFilename As String
	
	Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
	Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnBuffer As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Private Declare Function WritePrivateProfileSection Lib "kernel32"  Alias "WritePrivateProfileSectionA"(ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
	
	'Purpose: Read INI configuration data and comments array from file.
	'Returns INI configuration data array as return value and comments array as parameter.
	'Parameter: strFilename INI Filename.
	'Parameter: strSection INI file Section.
	'Parameter: varComments INI comment data array.
	Public Function GetDP5ConfigSection(ByRef strFilename As String, ByRef strSection As String, ByRef varComments As Object) As Object
		Dim varData As Object
		Dim varConfig As Object
		
		varData = GetAllIniSettings(strFilename, strSection) 'get the configuration
		If IsNothing(varData) Then
			GetDP5ConfigSection = varData
			Exit Function
		End If
		varConfig = GetDP5Commands(varData) 'remove the comments
		varComments = GetDP5Comments(varData) 'save the comments
		GetDP5ConfigSection = varConfig
	End Function
	
	'Purpose: Saves INI data array to file with optional append comments to configuration
	'Parameter: strFilename INI Filename.
	'Parameter: strSection INI file Section.
	'Parameter: varConfig INI configuration data array.
	'Parameter: varComments INI comment data array.
	Public Sub SaveDP5ConfigSection(ByRef strFilename As String, ByRef strSection As String, ByRef varConfig As Object, Optional ByRef varComments As Object = Nothing)
		Dim varData As Object
		Dim bAppendComments As Boolean
		
		If (IsNothing(varComments)) Then
			bAppendComments = False
		Else
			bAppendComments = True
		End If
		
		If (bAppendComments) Then
			If (IsArray(varComments)) Then 'create new array with comments
				varData = AppendDP5Comments(varConfig, varComments)
				SaveIniDataArray(strFilename, strSection, varData)
			Else
				bAppendComments = False 'shut off
			End If
		End If
		If (Not bAppendComments) Then 'the configuration has not been saved
			varData = AppendSemicolonsToConfig(varConfig)
			SaveIniDataArray(strFilename, strSection, varData)
		End If
	End Sub
	
	'Purpose: Saves INI data array to file.
	'Parameter: strFilename INI Filename.
	'Parameter: strSection INI file Section.
	'Parameter: varData INI data array.
	Public Sub SaveIniDataArray(ByRef strFilename As String, ByRef strSection As String, ByRef varData As Object)
		Dim strKey As String
		Dim strData As String
		Dim idxLRec As Integer
		Dim idxURec As Integer
		Dim idxRecord As Integer
		
		If (Not IsArray(varData)) Then Exit Sub 'if data array is empty cannot save
		idxLRec = LBound(varData, 1)
		idxURec = UBound(varData, 1)
		For idxRecord = idxLRec To idxURec
			strKey = varData(idxRecord, 0)
			strData = varData(idxRecord, 1)
			SaveToIni(strFilename, strSection, strKey, strData)
		Next
	End Sub
	
	'Purpose: Appends semicolons to configuration data array.
	'Returns updated configuration array.
	'Parameter: varConfig INI configuration data array.
	Public Function AppendSemicolonsToConfig(ByRef varConfig As Object) As Object
		Dim idxLRec As Integer
		Dim idxURec As Integer
		Dim idxRecord As Integer
		Dim varData As Object

		If (Not IsArray(varConfig)) Then
			AppendSemicolonsToConfig = Nothing
			Exit Function 'if data array is empty cannot save
		End If
		varData = CreateIniArr(varConfig) 'create a working copy of config array
		idxLRec = LBound(varConfig, 1)
		idxURec = UBound(varConfig, 1)
		For idxRecord = idxLRec To idxURec
			If (InStr(1, varData(idxRecord, 1), ";") = 0) Then
				varData(idxRecord, 1) = varData(idxRecord, 1) & ";"
			End If
		Next
		AppendSemicolonsToConfig = varData
	End Function
	
	'Purpose: Sets configuration data array keys and data to upper case.
	'Returns updated configuration array.
	'Parameter: varConfig INI configuration data array.
	Public Function ConfigToUCASE(ByRef varConfig As Object) As Object
		Dim idxLRec As Integer
		Dim idxURec As Integer
		Dim idxRecord As Integer
		Dim varData As Object
		
		If (Not IsArray(varConfig)) Then
			ConfigToUCASE = Nothing
			Exit Function 'if data array is empty cannot save
		End If
		varData = CreateIniArr(varConfig) 'create a working copy of config array
		idxLRec = LBound(varConfig, 1)
		idxURec = UBound(varConfig, 1)
		For idxRecord = idxLRec To idxURec
			varData(idxRecord, 0) = UCase(varData(idxRecord, 0))
			varData(idxRecord, 1) = UCase(varData(idxRecord, 1))
		Next
		ConfigToUCASE = varData
	End Function
	
	'Purpose: Clears configuration data.
	'Returns updated configuration array.
	'Parameter: varConfig INI configuration data array.
	Public Function ClearConfigData(ByRef varConfig As Object) As Object
		Dim idxLRec As Integer
		Dim idxURec As Integer
		Dim idxRecord As Integer
		Dim varData As Object
		
		If (Not IsArray(varConfig)) Then
			ClearConfigData = Nothing
			Exit Function 'if data array is empty cannot save
		End If
		varData = CreateIniArr(varConfig) 'create a working copy of config array
		idxLRec = LBound(varConfig, 1)
		idxURec = UBound(varConfig, 1)
		For idxRecord = idxLRec To idxURec
			varData(idxRecord, 1) = ""
		Next
		ClearConfigData = varData
	End Function
	
	'Purpose: Appends comments to INI data array.
	'Returns the Command Key (Field 0) and the Data (Field 1) with comments and semicolons.
	'Parameter: varConfig INI configuration data array.
	'Parameter: varComments INI comment data array.
	Public Function AppendDP5Comments(ByRef varConfig As Object, ByRef varComments As Object) As Object
		Dim idxLRec As Integer
		Dim idxURec As Integer
		Dim idxRecord As Integer
		Dim strKey As String = ""
		Dim strComment As String = ""
		Dim varData As Object
		
		If (Not IsArray(varConfig)) Then 'if config array is empty then cannot append
			AppendDP5Comments = Nothing
			Exit Function
		ElseIf (Not IsArray(varComments)) Then  'if comment array is empty no need to append
			AppendDP5Comments = varConfig
			Exit Function
		End If
		varData = CreateIniArr(varConfig) 'create a working copy of config array
		idxLRec = LBound(varConfig, 1) 'the config array size is needed
		idxURec = UBound(varConfig, 1)
		For idxRecord = idxLRec To idxURec
			strKey = varData(idxRecord, 0) 'get the key to search for comment
			strComment = FindIniData(strKey, varComments) 'search the comments array by key
			varData(idxRecord, 1) = varData(idxRecord, 1) & ";" & strComment 'append comment to config
		Next
		AppendDP5Comments = varData
	End Function
	
	'Purpose: Find Data for given Key in INI data array.
	'Returns the Data (Field 1).
	'Parameter: strKey INI item Key.
	'Parameter: varData INI data array.
	Public Function FindIniData(ByRef strKey As Object, ByRef varData As Object) As String
		Dim idxLRec As Integer
		Dim idxURec As Integer
		Dim idxRecord As Integer
		Dim strData As String
		
		If (Not IsArray(varData)) Then
			FindIniData = ""
			Exit Function
		End If
		
		idxLRec = LBound(varData, 1)
		idxURec = UBound(varData, 1)
		strData = ""
		For idxRecord = idxLRec To idxURec
			If (StrComp(strKey, varData(idxRecord, 0), CompareMethod.Text) = 0) Then
				strData = varData(idxRecord, 1)
				Exit For
			End If
		Next
		FindIniData = strData
	End Function
	
	'Purpose: Creates a formatted display string of INI file Section.
	'Parameter: strFilename INI Filename.
	'Parameter: strSection INI file Section.
	Public Function SectionString(ByRef strFilename As String, ByRef strSection As String) As String
		Dim varData As Object
		Dim idxKey As Integer
		Dim strData As String
		
		varData = GetAllIniSettings(strFilename, strSection)
		strData = vbCrLf & "[" & strSection & "]" & vbCrLf
		If IsArray(varData) Then
			For idxKey = LBound(varData, 1) To UBound(varData, 1)
				strData = strData & varData(idxKey, 0) & "=" & varData(idxKey, 1) & vbCrLf
			Next
		End If
		SectionString = strData
	End Function
	
	'Purpose: Creates a formatted display string of a INI data array.
	'Parameter: strSection INI file Section.
	'Parameter: varData INI data array.
	Public Function DataArrayString(ByRef strSection As String, ByRef varData As Object) As String
		Dim idxKey As Integer
		Dim strData As String

		strData = vbCrLf & "[" & strSection & "]" & vbCrLf
		If IsArray(varData) Then
			For idxKey = LBound(varData, 1) To UBound(varData, 1)
				strData = strData & varData(idxKey, 0) & "=" & varData(idxKey, 1) & vbCrLf
			Next
		End If
		DataArrayString = strData
	End Function
	
	'Purpose: Creates a formatted display string of INI file.
	'Parameter: strFilename INI Filename.
	Public Function IniFileString(ByRef strFilename As String) As String
		Dim varSections As Object
		Dim idxSection As Integer
		Dim strSection As String
		Dim strData As String
		
		IniFileString = ""
		strData = ""
		varSections = GetSectionNames(strFilename)
		If (IsArray(varSections)) Then
			For idxSection = LBound(varSections) To UBound(varSections)
				strSection = varSections(idxSection)
				strData = strData & SectionString(strFilename, strSection)
			Next
		Else
			strSection = varSections
			If (Len(strSection) > 0) Then
				strData = SectionString(strFilename, strSection)
			End If
		End If
		IniFileString = strData
	End Function
	
	'Purpose: Returns the INI file section names.
	'Parameter: strFilename INI Filename.
	Public Function GetSectionNames(ByRef strFilename As String) As Object
		Dim lngRetVal As Integer
		Dim strBuffer As String = ""
		Dim strSections() As String
		Dim idxSection As Integer
		Dim lPos As Integer
		Dim lNull As Integer
		Dim lBufSize As Integer

		ReDim strSections(0)
		strSections(0) = ""
		idxSection = 0
		lPos = 1
		strBuffer = Space(2048)
		lngRetVal = GetPrivateProfileSectionNames(strBuffer, Len(strBuffer), strFilename)
		lBufSize = Len(strBuffer)
		Do While (lPos < lBufSize)
			lNull = InStr(lPos, strBuffer, vbNullChar)
			If (lNull <> lPos) Then
				ReDim Preserve strSections(idxSection)
				strSections(idxSection) = Mid(strBuffer, lPos, lNull - lPos)
				idxSection = idxSection + 1
				lPos = lNull + 1
			Else
				Exit Do
			End If
		Loop
		GetSectionNames = strSections
	End Function
	
	'Purpose: Removes comments from INI section data.
	'Returns the Command Key (Field 0) and the Command Setting (Field 1) without semicolons.
	'Parameter: varData INI Section array.
	'Remarks: The comment whitespace is not trimmed.
	Public Function GetDP5Commands(ByRef varData As Object) As Object
		Dim idxLRec As Integer
		Dim idxURec As Integer
		Dim idxRecord As Integer
		Dim lPos As Integer
		Dim varConfig As Object
		
		If (Not IsArray(varData)) Then
			GetDP5Commands = Nothing
			Exit Function
		End If
		varConfig = CreateIniArr(varData)
		idxLRec = LBound(varConfig, 1)
		idxURec = UBound(varConfig, 1)
		For idxRecord = idxLRec To idxURec
			lPos = InStr(varConfig(idxRecord, 1), ";")
			If lPos = 1 Then 'not data only comments
				varConfig(idxRecord, 1) = ""
			ElseIf (lPos > 1) Then
				varConfig(idxRecord, 1) = Left(varConfig(idxRecord, 1), lPos - 1)
			Else
				'not found
			End If
		Next
		GetDP5Commands = varConfig
	End Function
	
	'Purpose: Extracts comments from INI section data.
	'Returns the Command Key (Field 0) and the Comments (Field 1) without semicolons.
	'Parameter: varData INI Section array.
	'Remarks: The comment whitespace is not trimmed.
	Public Function GetDP5Comments(ByRef varData As Object) As Object
		Dim idxLRec As Integer
		Dim idxURec As Integer
		Dim idxRecord As Integer
		Dim lPos As Integer
		Dim varComments As Object
		
		If (Not IsArray(varData)) Then
			GetDP5Comments = Nothing
			Exit Function
		End If
		varComments = CreateIniArr(varData)
		idxLRec = LBound(varComments, 1)
		idxURec = UBound(varComments, 1)
		For idxRecord = idxLRec To idxURec
			lPos = InStr(varComments(idxRecord, 1), ";")
			If (lPos > 0) Then
				varComments(idxRecord, 1) = Mid(varComments(idxRecord, 1), lPos + 1)
			Else 'not found
				varComments(idxRecord, 1) = ""
			End If
		Next
		GetDP5Comments = varComments
	End Function
	
	
	'Purpose: Creates INI data array of existing array size.  Optional copy of existing data.
	'Returns a INI data array.
	'Parameter: varData INI configuration data array.
	'Parameter: bCopyData Copy existing data array.
	Public Function CreateIniArr(ByRef varData As Object, Optional ByRef bCopyData As Boolean = True) As Object
		Dim idxLRec As Integer
		Dim idxURec As Integer
		Dim idxRecord As Integer
		Dim idxLFld As Integer
		Dim idxUFld As Integer
		Dim idxField As Integer
		Dim strData(,) As String

		If (Not IsArray(varData)) Then 'if data array is empty then cannot create new array
			CreateIniArr = Nothing
			Exit Function
		End If
		idxLRec = LBound(varData, 1) 'the config array size is needed
		idxURec = UBound(varData, 1)
		idxLFld = LBound(varData, 2)
		idxUFld = UBound(varData, 2)
		ReDim strData(idxURec, idxUFld) 'create new working array to hold combined data and comments
		If (bCopyData) Then
			For idxRecord = idxLRec To idxURec 'copy the config to the working array
				For idxField = idxLFld To idxUFld
					strData(idxRecord, idxField) = varData(idxRecord, idxField)
				Next
			Next
		End If
		CreateIniArr = strData
	End Function
	
	'Purpose: Creates INI data array of existing array size.  Optional copy of existing data.
	'Returns a INI data array.
	'Parameter: varData INI configuration data array.
	'Parameter: bCopyData Copy existing data array.
	Public Function CopyIniArr(ByRef varData As Object, Optional ByRef bCopyEmpty As Boolean = True) As Object
		Dim idxLRec As Integer
		Dim idxURec As Integer
		Dim idxRecord As Integer
		Dim idxLFld As Integer
		Dim idxUFld As Integer
		Dim idxField As Integer
		Dim strData(,) As String
		Dim idxNew As Integer
		Dim idxNotEmpty As Integer

		ReDim strData(0, 0)
		strData(0, 0) = ""
		If (Not IsArray(varData)) Then 'if data array is empty then cannot create new array
			CopyIniArr = Nothing
			Exit Function
		End If
		idxLRec = LBound(varData, 1) 'the config array size is needed
		idxURec = UBound(varData, 1)
		idxLFld = LBound(varData, 2)
		idxUFld = UBound(varData, 2)
		If (bCopyEmpty) Then
			ReDim strData(idxURec, idxUFld) 'create new working array to hold combined data and comments
			For idxRecord = idxLRec To idxURec 'copy the config to the working array
				For idxField = idxLFld To idxUFld
					strData(idxRecord, idxField) = varData(idxRecord, idxField)
				Next
			Next
		Else
			idxNotEmpty = -1
			For idxRecord = idxLRec To idxURec 'count the working array valid values
				If (Len(Trim(varData(idxRecord, 1))) > 0) Then idxNotEmpty = idxNotEmpty + 1
			Next
			If (idxNotEmpty >= 0) Then
				idxNew = 0
				ReDim strData(idxNotEmpty, idxUFld) 'create new working array to hold combined data and comments
				For idxRecord = idxLRec To idxURec 'copy the config to the working array
					If (Len(Trim(varData(idxRecord, 1))) > 0) Then
						If (idxNew <= idxNotEmpty) Then
							For idxField = idxLFld To idxUFld
								strData(idxNew, idxField) = varData(idxRecord, idxField)
							Next
							idxNew = idxNew + 1
						End If
					End If
				Next
			End If
		End If
		CopyIniArr = strData
	End Function
	
	'Purpose: Reads Ini Section, saves to List control.
	'Parameter: strFilename INI Filename.
	'Parameter: strSection INI Section requested.
	'Parameter: objList INI Section display list.
	'Parameter: bRemComments Remove comments from section data.
	Public Sub GetIniListEx2(ByRef strFilename As String, ByRef strSection As String, ByRef objList As Object, ByRef bRemComments As Boolean)
		Dim varData As Object
		Dim idxLRec As Integer
		Dim idxURec As Integer
		Dim idxRecord As Integer
		Dim strItem As String
		Dim lPos As Integer
		
		objList.Clear()
		varData = GetAllIniSettings(strFilename, strSection)
		If IsNothing(varData) Then Exit Sub
		idxLRec = LBound(varData, 1)
		idxURec = UBound(varData, 1)
		For idxRecord = idxLRec To idxURec
			strItem = varData(idxRecord, 0) & "=" & varData(idxRecord, 1)
			If (bRemComments) Then
				lPos = InStr(strItem, ";")
				If lPos = 1 Then
					strItem = ""
				ElseIf lPos > 1 Then
					strItem = Left(strItem, lPos - 1)
				End If
			End If
			objList.AddItem(strItem)
		Next
		objList.ListIndex = objList.TopIndex
	End Sub
	
	Public Sub SaveToIni(ByRef strFilename As String, ByRef strSection As String, ByRef strKey As String, ByRef strData As String)
		Dim lngRetVal As Integer
		lngRetVal = WritePrivateProfileString(strSection, strKey, strData, strFilename)
	End Sub
	
	Public Function GetFromIni(ByRef strFilename As String, ByRef strSection As String, ByRef strKey As String, Optional ByRef strDefault As String = "") As String
		Dim lngRetVal As Integer
		Dim lngBuffSz As Integer
		Dim strBuff As String
		
		strBuff = Space(255)
		lngBuffSz = 255
		lngRetVal = GetPrivateProfileString(strSection, strKey, strDefault, strBuff, lngBuffSz, strFilename)
		lngBuffSz = InStr(1, strBuff, Chr(0), CompareMethod.Binary)
		If lngBuffSz > 0 Then
			GetFromIni = Trim(Left(strBuff, lngBuffSz - 1))
		Else
			GetFromIni = strBuff
		End If
	End Function
	
	Public Function GetAllIniSettings(ByRef strFilename As String, ByRef Section As String) As Object
		Dim lRet As Integer
		Dim strTemp As String
		Dim Table(,) As String
		Dim Table2(,) As String
		Dim iPnt As Short
		Dim iPnt2 As Short
		Dim iPosit As Short
		Dim iLen As Integer

		strTemp = Space(MAX_SECTION_SIZE)
		ReDim Table(0, 0)
		Table(0, 0) = ""
		ReDim Table2(0, 0)
		Table2(0, 0) = ""

		lRet = GetPrivateProfileSection(Section, strTemp, MAX_SECTION_SIZE, strFilename)
		iLen = lRet
		iPnt = 0
		'For Redim+Preserve tables only the las index can be changed
		If Left(strTemp, 2) = Chr(0) & Chr(0) Then
			ReDim Table2(1, 1)
			Table2(0, 0) = ""
			Table2(0, 1) = ""
			Table2(1, 0) = ""
			Table2(1, 1) = ""
			GetAllIniSettings = Table2
			Exit Function
		End If
		
		If (lRet = 0) Then
			ReDim Table2(1, 1)
			Table2(0, 0) = ""
			Table2(0, 1) = ""
			Table2(1, 0) = ""
			Table2(1, 1) = ""
			GetAllIniSettings = Table2
			Exit Function
		End If

		iPosit = 1
		'count keys in table
		Do While iPosit > 0
			iPosit = InStr(iPosit + 1, strTemp, Chr(0))
			iPnt = iPnt + 1
		Loop
		ReDim Table(1, iPnt)
		iPosit = 0
		iPnt = 0
		Do While Left(strTemp, 1) <> Chr(0)
			iPosit = InStr(strTemp, "=")
			Table(0, iPnt) = Left(strTemp, iPosit - 1)
			strTemp = Mid(strTemp, iPosit + 1)
			iPosit = InStr(strTemp, Chr(0))
			Table(1, iPnt) = Left(strTemp, iPosit - 1)
			strTemp = Mid(strTemp, iPosit + 1)
			iPnt = iPnt + 1
		Loop
		ReDim Table2(iPnt - 1, 1)

		For iPnt2 = 0 To iPnt - 1
			Table2(iPnt2, 0) = Table(0, iPnt2)
			Table2(iPnt2, 1) = Table(1, iPnt2)
		Next

		GetAllIniSettings = Table2
	End Function
	
	Public Function DeleteIniSetting(ByRef strFilename As String, ByRef Section As String, Optional ByRef Key As String = "") As Integer
		Dim lRet As Integer
		DeleteIniSetting = 0
		If Key = "" Then
			lRet = WritePrivateProfileString(Section, vbNullString, vbNullString, strFilename)
		Else
			lRet = WritePrivateProfileString(Section, Key, vbNullString, strFilename)
		End If
		DeleteIniSetting = lRet
	End Function
	
	Public Sub GetIniList(ByRef strFilename As String, ByRef strSection As String, ByRef objList As Object)
		Dim varAllIni As Object
		Dim i As Short
		Dim idxLRec As Integer
		Dim idxLFld As Integer
		Dim idxURec As Integer
		Dim idxUFld As Integer
		
		objList.Clear()
		varAllIni = GetAllIniSettings(strFilename, strSection)
		If IsNothing(varAllIni) Then Exit Sub
		
		idxLRec = LBound(varAllIni, 1)
		idxLFld = LBound(varAllIni, 2)
		idxURec = UBound(varAllIni, 1)
		idxUFld = UBound(varAllIni, 2)
		
		For i = idxLRec To idxURec
			objList.AddItem(varAllIni(i, 1))
		Next
		objList.ListIndex = objList.TopIndex
	End Sub
	
	'------------------------------------------------------
	' Function:   ValidFilename as string
	' arguments:  strFilename		 a filename
	' Result: ValidFilename returns a valid filename
	'-------------------------------------------------------
	Function ValidFilename(ByRef strFilename As String) As String
		Dim i As Short
		Dim strFname As String
		Dim legalChar As String
		ValidFilename = strFilename
		strFname = Trim(strFilename) 'remove white space
		' Remove for illegal characters
		legalChar = " !#$%&'()-0123456789@ABCDEFGHIJKLMNOPQRSTUVWXYZ^_`{}~.üäöÄÖÜß"
		For i = 1 To Len(strFname)
			If InStr(legalChar, UCase(Mid(strFname, i, 1))) = 0 Then
				Mid(strFname, i, 1) = " "
			End If
		Next i
		ValidFilename = strFname
	End Function
	
	Public Sub SaveApplicationSettings(ByRef frm As System.Windows.Forms.Form, Optional ByRef AppSet As Boolean = True, Optional ByRef CfgSet As Boolean = True, Optional ByRef ComSet As Boolean = True)
		Dim strFilename As String
		strFilename = My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".ini"
		
		If (AppSet) Then
			'Application Settings
			'			SaveToIni(strFilename, IniSectionApp, frm.Name & "_Top", CStr(VB6.PixelsToTwipsY(frm.Top)))
			'			SaveToIni(strFilename, IniSectionApp, frm.Name & "_Left", CStr(VB6.PixelsToTwipsX(frm.Left)))
		End If
		
		If (CfgSet) Then
			'Configuration Settings
			SaveToIni(strFilename, IniSectionApp, "SendCoarseFineGain", CStr(s.SendCoarseFineGain))
			SaveToIni(strFilename, IniSectionApp, "SCAEnabled", CStr(s.SCAEnabled))
		End If
		
		If (ComSet) Then
			'Communications Interface Settings
			SaveToIni(strFilename, IniSectionApp, "CurrentInterface", CStr(s.CurrentInterface))
			SaveToIni(strFilename, IniSectionApp, "ComPort", CStr(s.ComPort))
			SaveToIni(strFilename, IniSectionApp, "SockAddr", CStr(s.SockAddr))
			SaveToIni(strFilename, IniSectionApp, "cstrSockAddr", (s.cstrSockAddr))
			SaveToIni(strFilename, IniSectionApp, "InetPort", CStr(s.InetPort))
		End If
	End Sub
	
	Public Sub LoadApplicationSettings(ByRef frm As System.Windows.Forms.Form, Optional ByRef AppSet As Boolean = True, Optional ByRef CfgSet As Boolean = True, Optional ByRef ComSet As Boolean = True)
		Dim strFilename As String
		strFilename = My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".ini"
		If (AppSet) Then
			'Application Settings
			'			frm.Top = VB6.TwipsToPixelsY(CSng(GetFromIni(strFilename, IniSectionApp, frm.Name & "_Top", CStr(VB6.PixelsToTwipsY(frm.Top)))))
			'			frm.Left = VB6.TwipsToPixelsX(CSng(GetFromIni(strFilename, IniSectionApp, frm.Name & "_Left", CStr(VB6.PixelsToTwipsX(frm.Left)))))
		End If

		If (CfgSet) Then
			'Configuration Settings
			s.SendCoarseFineGain = CBool(GetFromIni(strFilename, IniSectionApp, "SendCoarseFineGain", CStr(False)))
			s.SCAEnabled = CBool(GetFromIni(strFilename, IniSectionApp, "SCAEnabled", CStr(False)))
		End If
		
		If (ComSet) Then
			'Communications Interface Settings
			s.CurrentInterface = CByte(GetFromIni(strFilename, IniSectionApp, "CurrentInterface", CStr(USB)))
			s.ComPort = CInt(GetFromIni(strFilename, IniSectionApp, "ComPort", CStr(1)))
			s.SockAddr = CInt(GetFromIni(strFilename, IniSectionApp, "SockAddr", CStr(&HC0A80064)))
			s.cstrSockAddr = GetFromIni(strFilename, IniSectionApp, "cstrSockAddr", "192.168.0.100")
			s.InetPort = CInt(GetFromIni(strFilename, IniSectionApp, "InetPort", CStr(10001)))
		End If
	End Sub
		
	'Saves all INI Settings - NOTE:OVERWRITES ENTIRE SECTION - Use with caution!
	'Useage - copy all section settings IN ORDER to Settings array, overwrites section settings
	'Settings Variant Type of String Array -> Setting(Index,0),Setting(Index,1) -> 0=key,1=data
	Public Sub SaveAllIniSettingsEx(ByRef strFilename As String, ByRef Section As String, ByRef Settings As Object)
		Dim lngRetVal As Integer
		Dim strData As String = ""
		Dim strSettings As String = ""
		Dim strNULL As String = ""
		Dim idxSetting As Integer
		Dim strSep As String = ""
		
		strNULL = Chr(0)
		If (Not IsArray(Settings)) Then Exit Sub
		If (UBound(Settings, 2) <> 1) Then Exit Sub
		
		For idxSetting = LBound(Settings, 1) To UBound(Settings, 1)
			strData = CStr(Settings(idxSetting, 1))
			If (InStr(strData, ";") > 0) Then
				strSep = ""
			Else
				strSep = ";"
			End If
			strSettings = strSettings & CStr(Settings(idxSetting, 0)) & "=" & strData & strSep & strNULL
		Next
		strSettings = strSettings & strNULL
		lngRetVal = WritePrivateProfileSection(Section, strSettings, strFilename)
	End Sub
	
	Public Sub SaveToIniEx(ByRef strFilename As String, ByRef strSection As String, ByRef strKey As String, ByRef strData As String)
		Dim lngRetVal As Integer
		lngRetVal = WritePrivateProfileString(strSection, strKey, strData, strFilename)
	End Sub
	
	Public Function GetFromIniEx(ByRef strFilename As String, ByRef strSection As String, ByRef strKey As String, Optional ByRef strDefault As String = "") As String
		Dim lngRetVal As Integer
		Dim lngBuffSz As Integer
		Dim strBuff As String
		
		strBuff = Space(255)
		lngBuffSz = 255
		lngRetVal = GetPrivateProfileString(strSection, strKey, strDefault, strBuff, lngBuffSz, strFilename)
		lngBuffSz = InStr(1, strBuff, Chr(0), CompareMethod.Binary)
		If lngBuffSz > 0 Then
			GetFromIniEx = Trim(Left(strBuff, lngBuffSz - 1))
		Else
			GetFromIniEx = strBuff
		End If
	End Function
	
	Public Function DeleteIniSettingEx(ByRef strFilename As String, ByRef Section As String, Optional ByRef Key As String = "") As Integer
		Dim lRet As Integer
		DeleteIniSettingEx = 0
		If Key = "" Then
			lRet = WritePrivateProfileString(Section, vbNullString, vbNullString, strFilename)
		Else
			lRet = WritePrivateProfileString(Section, Key, vbNullString, strFilename)
		End If
		DeleteIniSettingEx = lRet
	End Function

	Public Function GetAllIniSettingsEx(ByRef strFilename As String, ByRef Section As String) As Object
		Dim lRet As Integer
		Dim strTemp As String
		Dim Table(,) As String
		Dim Table2(,) As String
		Dim iPnt As Short
		Dim iPnt2 As Short
		Dim iPosit As Short

		ReDim Table(0, 0)
		Table(0, 0) = ""
		ReDim Table2(0, 0)
		Table2(0, 0) = ""

		strTemp = Space(MAX_SECTION_SIZE)
		lRet = GetPrivateProfileSection(Section, strTemp, MAX_SECTION_SIZE, strFilename)
		iPnt = 0
		'For Redim+Preserve tables only the las index can be changed
		If Left(strTemp, 2) = Chr(0) & Chr(0) Then
			GetAllIniSettingsEx = Table2
			Exit Function
		End If

		If (lRet = 0) Then
			GetAllIniSettingsEx = Table2
			Exit Function
		End If

		iPosit = 1
		'count keys in table
		Do While iPosit > 0
			iPosit = InStr(iPosit + 1, strTemp, Chr(0))
			iPnt = iPnt + 1
		Loop
		ReDim Table(1, iPnt)
		iPosit = 0
		iPnt = 0

		Do While Left(strTemp, 1) <> Chr(0)
			iPosit = InStr(strTemp, "=")
			Table(0, iPnt) = Left(strTemp, iPosit - 1)
			strTemp = Mid(strTemp, iPosit + 1)
			iPosit = InStr(strTemp, Chr(0))
			Table(1, iPnt) = Left(strTemp, iPosit - 1)
			strTemp = Mid(strTemp, iPosit + 1)
			iPnt = iPnt + 1
		Loop
		ReDim Table2(iPnt - 1, 1)

		For iPnt2 = 0 To iPnt - 1
			Table2(iPnt2, 0) = Table(0, iPnt2)
			Table2(iPnt2, 1) = Table(1, iPnt2)
		Next
		GetAllIniSettingsEx = Table2
	End Function
End Module
