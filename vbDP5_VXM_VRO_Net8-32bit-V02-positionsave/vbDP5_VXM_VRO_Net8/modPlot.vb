Option Strict Off
Option Explicit On

Module modPlot

    Public m_Bitmap As Bitmap
    Public m_Graphics As Graphics

    'converts vb type 24bit rgb (BlueGreenRed) to QBColor
    Public Function RgbToQbColor(ByRef PlotColor As Integer) As Object
        Dim PlotColorIndex As Short 'QBColors 0-15
        Select Case PlotColor
            Case &H0 : PlotColorIndex = 0 'Black
            Case &H800000 : PlotColorIndex = 1 'Blue
            Case &H8000 : PlotColorIndex = 2 'Green
            Case &H808000 : PlotColorIndex = 3 'Cyan
            Case &H80 : PlotColorIndex = 4 'Red
            Case &H800080 : PlotColorIndex = 5 'Magenta
            Case &H8080 : PlotColorIndex = 6 'Brown
            Case &HC0C0C0 : PlotColorIndex = 7 'White
            Case &H808080 : PlotColorIndex = 8 'Grey
            Case &HFF0000 : PlotColorIndex = 9 'Light Blue
            Case &HFF00 : PlotColorIndex = 10 'Light Green
            Case &HFFFF00 : PlotColorIndex = 11 'Light Cyan
            Case &HFF : PlotColorIndex = 12 'Light Red
            Case &HFF00FF : PlotColorIndex = 13 'Light Magenta
            Case &HFFFF : PlotColorIndex = 14 'Yellow
            Case &HFFFFFF : PlotColorIndex = 15 'Bright White
            Case Else
                PlotColorIndex = &H0S 'unknown color
        End Select
        PlotColorIndex = PlotColorIndex + 1
        If (PlotColorIndex >= 16) Then PlotColorIndex = 0
        RgbToQbColor = QBColor(PlotColorIndex)
    End Function

    Public Sub AllocateBitmap(ByRef picPlot As System.Windows.Forms.PictureBox)
        m_Bitmap = New Bitmap(picPlot.Width, picPlot.Height)    'create a new bitmap to plot spectrum data on
        m_Graphics = Graphics.FromImage(m_Bitmap)               'create a graphics device interface to draw on the bitmap
        picPlot.Image = m_Bitmap                                'display the bitmap in the picturebox
    End Sub

    Public Sub PicDrawLine(ByRef picPlot As System.Windows.Forms.PictureBox, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal FillColor As Color)
        Dim m_Pen As New Pen(FillColor, 1)
        Dim m_Height As Integer
        Dim m_Y1 As Single
        Dim m_Y2 As Single

        m_Height = picPlot.Height
        m_Y1 = picPlot.Height - Y1
        If (m_Y1 < 0.0) Then m_Y1 = 0.0
        m_Y2 = picPlot.Height - Y2
        If (m_Y2 < 0.0) Then m_Y2 = 0.0

        m_Graphics.DrawLine(m_Pen, X1, m_Y1, X2, m_Y2)
        picPlot.Image = m_Bitmap
        m_Pen.Dispose()
    End Sub

    Public Sub Plot_Spectrum(ByRef picMCAPlot As System.Windows.Forms.PictureBox, ByRef SPECTRUM As Spec, ByRef isLinear As Boolean, ByRef isLines As Boolean)
        Dim X As Short
        Dim Y As Short
        Dim MCAMax As Integer
        Dim MCAPeak As Integer
        Dim Buffer(64) As Object
        Dim MCAScale As Integer
        Dim PlotXMin As Integer
        Dim PlotXMax As Integer
        Dim LogScaleValue As Integer

        Dim xW As Double
        Dim yH As Double
        Dim xScale As Double
        Dim yScale As Double
        Dim plotColor As Color
        Dim bgColor As Color
        Dim x1 As Double
        Dim x2 As Double
        Dim y1 As Double
        Dim y2 As Double

        AllocateBitmap(picMCAPlot)
        bgColor = picMCAPlot.BackColor
        plotColor = Color.Red
        MCAMax = 0
        MCAPeak = 0
        For X = 0 To SPECTRUM.Channels - 1
            If SPECTRUM.DATA(X) > MCAMax Then
                MCAMax = SPECTRUM.DATA(X)
                MCAPeak = X
            End If
        Next X
        xW = picMCAPlot.Width : If (xW < 10) Then xW = 10
        xScale = xW / SPECTRUM.Channels
        yH = picMCAPlot.Height : If (yH < 10) Then yH = 10
        If isLinear Then
            MCAScale = Fix(MCAMax + 10 + (10.0# * System.Math.Log(MCAMax + 1)))
        Else
            LogScaleValue = Fix(System.Math.Log(MCAMax + 1) / System.Math.Log(10) + 1)
            MCAScale = 10 ^ LogScaleValue
        End If
        PlotXMin = 0
        PlotXMax = SPECTRUM.Channels - 1

        If isLinear Then
            yScale = yH / MCAScale
            x1 = PlotXMin * xScale
            y1 = 0 * yScale
            x2 = (PlotXMax + 1) * xScale
            y2 = MCAScale * yScale
            PicDrawLine(picMCAPlot, x1, y1, x2, y2, bgColor)
        Else
            yScale = yH / (System.Math.Log(MCAScale))
            x1 = PlotXMin * xScale
            y1 = 0 * yScale
            x2 = (PlotXMax + 1) * xScale
            y2 = System.Math.Log(MCAScale) * yScale
            PicDrawLine(picMCAPlot, x1, y1, x2, y2, bgColor)
        End If

        If isLinear Then
            x1 = PlotXMin * xScale
            y1 = 0 * yScale
            x2 = (PlotXMax + 1) * xScale
            y2 = MCAScale * yScale
            PicDrawLine(picMCAPlot, x1, y1, x2, y2, bgColor)
            For X = 1 To 9
                x1 = PlotXMin * xScale
                y1 = MCAScale * X / 10 * yScale
                x2 = (PlotXMax + 1) * xScale
                y2 = (MCAScale * X / 10) * yScale
                PicDrawLine(picMCAPlot, x1, y1, x2, y2, Color.White)
            Next X
        Else
            x1 = PlotXMin * xScale
            y1 = 0
            x2 = (PlotXMax + 1) * xScale
            y2 = System.Math.Log(MCAScale) * yScale
            PicDrawLine(picMCAPlot, x1, y1, x2, y2, bgColor)
            For Y = 1 To LogScaleValue
                For X = 1 To 9
                    x1 = PlotXMin * xScale
                    y1 = System.Math.Log(MCAScale * X / (10 ^ Y)) * yScale
                    x2 = (PlotXMax + 1) * xScale
                    y2 = System.Math.Log(MCAScale * X / (10 ^ Y)) * yScale
                    PicDrawLine(picMCAPlot, x1, y1, x2, y2, Color.LightGray)
                Next X
            Next Y
        End If

        If isLines Then
            For X = PlotXMin To PlotXMax - 1
                If isLinear Then
                    x1 = X * xScale
                    y1 = SPECTRUM.DATA(X) * yScale
                    x2 = (X + 1) * xScale
                    y2 = SPECTRUM.DATA(X + 1) * yScale
                    PicDrawLine(picMCAPlot, x1, y1, x2, y2, plotColor)
                Else
                    x1 = X * xScale
                    y1 = System.Math.Log(1.0 + SPECTRUM.DATA(X)) * yScale
                    x2 = (X + 1) * xScale
                    y2 = System.Math.Log(1.0 + SPECTRUM.DATA(X + 1)) * yScale
                    PicDrawLine(picMCAPlot, x1, y1, x2, y2, plotColor)
                End If
            Next X
        Else
            For X = PlotXMin To PlotXMax
                If isLinear Then
                    x1 = X * xScale
                    y1 = 0
                    x2 = (X + 0.5) * xScale
                    y2 = SPECTRUM.DATA(X) * yScale
                    PicDrawLine(picMCAPlot, x1, y1, x2, y2, plotColor)
                Else
                    x1 = X * xScale
                    y1 = 0
                    x2 = (X + 0.5) * xScale
                    y2 = System.Math.Log(1.0 + SPECTRUM.DATA(X)) * yScale
                    PicDrawLine(picMCAPlot, x1, y1, x2, y2, plotColor)
                End If
            Next X
        End If
    End Sub

    Public Sub Plot_Scope(ByRef picScope As System.Windows.Forms.PictureBox, ByRef Scope() As Byte, ByRef PlotColor As Color)
        Dim X As Short
        Dim xW As Double
        Dim yH As Double
        Dim xScale As Double
        Dim yScale As Double
        Dim scopeColor As Color
        Dim x1 As Double
        Dim x2 As Double
        Dim y1 As Double
        Dim y2 As Double

        AllocateBitmap(picScope)
        xW = picScope.Width : If (xW < 10) Then xW = 10
        yH = picScope.Height : If (yH < 10) Then yH = 10
        xScale = xW / 2047.0
        yScale = yH / 255.0
        scopeColor = picScope.BackColor
        PicDrawLine(picScope, 0, 0, 2047 * xScale, 255 * yScale, scopeColor)
        For X = 0 To 2046
            x1 = X * xScale
            y1 = Scope(X) * yScale
            x2 = (X + 1) * xScale
            y2 = Scope(X + 1) * yScale
            PicDrawLine(picScope, x1, y1, x2, y2, PlotColor)
        Next X
    End Sub

    Public Sub SaveSpectrum(ByRef SPECTRUM As Spec, ByRef strStatus As String, ByRef strcfg As String, ByRef strTag As String, ByRef strDescription As String)
        Dim idxData As Integer
        Dim iFile As Short
        Dim strData As String = ""
        Dim saveFileDialogOut As New SaveFileDialog()
        saveFileDialogOut.Filter = dlgMCA_Filter
        If saveFileDialogOut.ShowDialog = DialogResult.OK Then
            iFile = FreeFile()
            FileOpen(iFile, saveFileDialogOut.FileName, OpenMode.Output)
            PrintLine(iFile, "<<PMCA SPECTRUM>>")
            PrintLine(iFile, "TAG - " & strTag)
            PrintLine(iFile, "DESCRIPTION - " & strDescription)

            Select Case SPECTRUM.Channels
                Case 256
                    PrintLine(iFile, "GAIN - 0")
                Case 512
                    PrintLine(iFile, "GAIN - 1")
                Case 1024
                    PrintLine(iFile, "GAIN - 2")
                Case 2048
                    PrintLine(iFile, "GAIN - 3")
                Case 4096
                    PrintLine(iFile, "GAIN - 4")
                Case 8192
                    PrintLine(iFile, "GAIN - 5")
                Case Else
                    PrintLine(iFile, "GAIN - 2")
            End Select

            PrintLine(iFile, "THRESHOLD - 0")
            PrintLine(iFile, "LIVE_MODE - 0")
            PrintLine(iFile, "PRESET_TIME - 0")
            PrintLine(iFile, "LIVE_TIME - 0")

            PrintLine(iFile, "REAL_TIME - " & STATUS.AccumulationTime)
            'NOTE: your acquisition start time goes here
            PrintLine(iFile, "START_TIME - " & Today.ToString("MM/dd/yyyy") & " " & TimeOfDay.ToString("HH:mm:ss"))
            PrintLine(iFile, "SERIAL_NUMBER - " & STATUS.SerialNumber)

            PrintLine(iFile, "<<DATA>>")
            For idxData = 0 To SPECTRUM.Channels - 1
                strData = SPECTRUM.DATA(idxData).ToString
                strData.Trim()
                PrintLine(iFile, strData)
            Next idxData
            PrintLine(iFile, "<<END>>")
            PrintLine(iFile, "<<DP5 CONFIGURATION>>")
            PrintLine(iFile, strcfg)
            PrintLine(iFile, "<<DP5 CONFIGURATION END>>")
            PrintLine(iFile, "<<DPP STATUS>>")
            If (strStatus = "") Then
                PrintLine(iFile, "Device Type: DP5")
                PrintLine(iFile, "Serial Number: 0")
            Else
                PrintLine(iFile, strStatus)
            End If
            PrintLine(iFile, "<<DPP STATUS END>>")
            FileClose(iFile)
        End If
    End Sub

    'This batch version will save without the save dialog and will be used in loop
    Public Sub SaveSpectrumBacth(ByRef SPECTRUM As Spec, ByRef strStatus As String, ByRef strcfg As String, ByRef strTag As String, ByRef strDescription As String, ByRef outfilename As String)
        Dim idxData As Integer
        Dim iFile As Short
        Dim strData As String = ""
        Dim saveFileDialogOut As New SaveFileDialog()
        iFile = FreeFile()
        FileOpen(iFile, outfilename, OpenMode.Output) 'automated version
        PrintLine(iFile, "<<PMCA SPECTRUM>>")
        PrintLine(iFile, "TAG - " & strTag)
        PrintLine(iFile, "DESCRIPTION - " & strDescription)

        Select Case SPECTRUM.Channels
            Case 256
                PrintLine(iFile, "GAIN - 0")
            Case 512
                PrintLine(iFile, "GAIN - 1")
            Case 1024
                PrintLine(iFile, "GAIN - 2")
            Case 2048
                PrintLine(iFile, "GAIN - 3")
            Case 4096
                PrintLine(iFile, "GAIN - 4")
            Case 8192
                PrintLine(iFile, "GAIN - 5")
            Case Else
                PrintLine(iFile, "GAIN - 2")
        End Select

        PrintLine(iFile, "THRESHOLD - 0")
        PrintLine(iFile, "LIVE_MODE - 0")
        PrintLine(iFile, "PRESET_TIME - 0")
        PrintLine(iFile, "LIVE_TIME - 0")

        PrintLine(iFile, "REAL_TIME - " & STATUS.AccumulationTime)
        'NOTE: your acquisition start time goes here
        PrintLine(iFile, "START_TIME - " & Today.ToString("MM/dd/yyyy") & " " & TimeOfDay.ToString("HH:mm:ss"))
        PrintLine(iFile, "SERIAL_NUMBER - " & STATUS.SerialNumber)

        PrintLine(iFile, "<<DATA>>")
        For idxData = 0 To SPECTRUM.Channels - 1
            strData = SPECTRUM.DATA(idxData).ToString
            strData.Trim()
            PrintLine(iFile, strData)
        Next idxData
        PrintLine(iFile, "<<END>>")
        PrintLine(iFile, "<<DP5 CONFIGURATION>>")
        PrintLine(iFile, strcfg)
        PrintLine(iFile, "<<DP5 CONFIGURATION END>>")
        PrintLine(iFile, "<<DPP STATUS>>")
        If (strStatus = "") Then
            PrintLine(iFile, "Device Type: DP5")
            PrintLine(iFile, "Serial Number: 0")
        Else
            PrintLine(iFile, strStatus)
        End If
        PrintLine(iFile, "<<DPP STATUS END>>")
        FileClose(iFile)
    End Sub

    Public Sub WriteTextFile(ByRef strFilename As String, ByRef strData As String, Optional ByRef OverWrite As Boolean = False)
        Dim intFile As Short

        intFile = FreeFile()
        If (OverWrite) Then
            FileOpen(intFile, strFilename, OpenMode.Output)
        Else
            FileOpen(intFile, strFilename, OpenMode.Append)
        End If
        PrintLine(intFile, strData)
        FileClose(intFile)
    End Sub

End Module
