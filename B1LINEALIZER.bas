Attribute VB_Name = "B1LINEALIZER"
Option Explicit

' ==================================================================================
' B1LINEALIZER - Legacy Report Parser
' Ported from Python to VBA
' ==================================================================================

Sub ParseReport()
    Dim ws As Worksheet
    Dim fso As Object, ts As Object
    Dim filePath As Variant
    Dim line As String
    Dim lines() As String
    Dim lineCount As Long, arrSize As Long
    Dim i As Long
    
    ' State Machine Variables
    Dim inSkipMode As Boolean
    Dim dashCount As Long
    Dim currentCard As String, currentName As String
    
    ' Record Buffers
    Dim pendingLine1 As String
    Dim hasPendingLine1 As Boolean
    Dim indentLen As Long
    
    ' Output Array (Buffer)
    Dim outputData() As Variant
    Dim outRow As Long
    
    ' Field Variables for Extraction
    Dim rawL1 As String, rawL2 As String
    Dim cleanL1 As String
    Dim rsVal As String
    
    ' 1. Select File
    filePath = Application.GetOpenFilename("Text Files (*.txt), *.txt", , "Select Report File")
    If filePath = False Then Exit Sub
    
    ' 2. Setup Output Sheet
    Set ws = ActiveSheet
    ws.Cells.Clear
    ws.Cells.NumberFormat = "@"
    
    ' Write Headers
    Dim headers As Variant
    headers = Array("TARJETA", "NOMBRE", "OPERAC", "RS", "MOVIM", "MONEDA ORIGINAL", "IMPORTE ORIGINAL", _
                    "MONEDA VISA", "IMPORT VISA", "MONEDA AFECTADO", "IMPORTE AFECTADO", "TIPO CUENTA", _
                    "CUENTA AFECTADA", "FECOPE", "HORA", "FBASE1", "EXPIRACION", _
                    "TERMINAL", "TIPO", "IDENTIFICACION", "ESTABLECIMIENTO", "CIUDAD", "PAIS", _
                    "BIN ADQUIR.", "PIN", "VIS.REFER", "TRNX", "CAVV", "POS.C.CODE")
    ws.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    
    ' 3. Read File
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1) ' 1 = ForReading
    
    ' Pre-allocate buffer (Expand if needed, assumption 50k lines)
    arrSize = 50000
    ReDim outputData(1 To arrSize, 1 To UBound(headers) + 1)
    outRow = 0
    
    inSkipMode = False
    dashCount = 0
    hasPendingLine1 = False
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Parsing file..."
    
    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        Dim stripped As String
        stripped = Trim(line)
        
        ' --- A. Page Header Skip Logic ---
        ' Enter skip mode on *****
        If InStr(stripped, "*****") > 0 Then
            inSkipMode = True
            dashCount = 0
        End If
        
        ' Exit skip mode on 2nd -----
        If inSkipMode And InStr(stripped, "-----") > 0 Then
            dashCount = dashCount + 1
            If dashCount >= 2 Then
                inSkipMode = False
                GoTo NextLine ' Skip the dash line itself
            End If
        End If
        
        If inSkipMode Then GoTo NextLine
        If stripped = "" Then GoTo NextLine
        
        ' --- B. Card Header Identification ---
        If Left(stripped, 9) = "- TARJETA" Then
            ' Extract Card and Name
            ' Format: - TARJETA 123456 LINEA NOMBRE TEST USER
            ' Simple parsing by splitting
            ' We assume Card is the 3rd word, Name is everything after "NOMBRE"
            
            Dim parts() As String
            parts = Split(Application.Trim(stripped), " ")
            If UBound(parts) >= 2 Then
                currentCard = parts(2)
            End If
            
            Dim namePos As Long
            namePos = InStr(stripped, "NOMBRE")
            If namePos > 0 Then
                currentName = Trim(Mid(stripped, namePos + 6))
            End If
            
            GoTo NextLine
        End If
        
        ' --- C. Separator Filter ---
        If Left(stripped, 5) = "-----" Or Left(stripped, 5) = "*****" Then GoTo NextLine
        If currentCard = "" Then GoTo NextLine ' Skip valid data if no card seen yet
        
        ' --- D. Data Row Processing (Pairs) ---
        If Not hasPendingLine1 Then
            ' This is LINE 1
            pendingLine1 = line
            hasPendingLine1 = True
            
            ' Calculate Indentation (Coupling Logic)
            ' Count leading spaces
            indentLen = Len(line) - Len(LTrim(line))
        Else
            ' This is LINE 2
            rawL1 = pendingLine1
            rawL2 = line
            
            ' Apply Indentation Logic
            ' Strip exact same indent from Line 2 as Line 1
            ' Mid is 1-based. Start at indentLen + 1
            cleanL1 = Mid(rawL1, indentLen + 1)
            
            Dim cleanL2 As String
            If Len(rawL2) > indentLen Then
                cleanL2 = Mid(rawL2, indentLen + 1)
            Else
                cleanL2 = ""
            End If
            
            ' --- E. Extraction & Validation ---
            
            ' Validate RS (Position 9-10 in 1-based string, length 2)
            ' Python (8,10) -> VBA Mid(cleanL1, 9, 2)
            rsVal = Trim(Mid(cleanL1, 9, 2))
            
            If IsNumeric(rsVal) Then
                ' Valid Record, Parse and Store
                outRow = outRow + 1
                If outRow > arrSize Then
                    arrSize = arrSize + 10000
                    ReDim Preserve outputData(1 To arrSize, 1 To UBound(headers) + 1)
                End If
                
                ' Common Info
                outputData(outRow, 1) = currentCard ' Force Text
                outputData(outRow, 2) = currentName
                
                ' Line 1 Fields (Python 0-based start, Length) -> VBA Mid(str, Start+1, Length)
                outputData(outRow, 3) = Trim(Mid(cleanL1, 1, 6))   ' OPERAC
                outputData(outRow, 4) = rsVal                      ' RS
                outputData(outRow, 5) = Trim(Mid(cleanL1, 13, 5))  ' MOVIM
                outputData(outRow, 6) = Trim(Mid(cleanL1, 20, 3))  ' MONEDA ORG
                outputData(outRow, 7) = CleanImporte(Mid(cleanL1, 23, 15)) ' IMPORTE ORG
                outputData(outRow, 8) = Trim(Mid(cleanL1, 38, 3))  ' MONEDA VISA
                outputData(outRow, 9) = CleanImporte(Mid(cleanL1, 41, 15)) ' IMPORTE VISA
                outputData(outRow, 10) = Trim(Mid(cleanL1, 56, 3)) ' MONEDA AFEC
                outputData(outRow, 11) = CleanImporte(Mid(cleanL1, 59, 15)) ' IMPORTE AFEC
                outputData(outRow, 12) = Trim(Mid(cleanL1, 74, 4)) ' TIPO CUENTA
                outputData(outRow, 13) = Trim(Mid(cleanL1, 78, 20)) ' CUENTA AFECTADA
                outputData(outRow, 14) = Trim(Mid(cleanL1, 98, 9))  ' FECOPE
                outputData(outRow, 15) = Trim(Mid(cleanL1, 107, 7)) ' HORA
                outputData(outRow, 16) = Trim(Mid(cleanL1, 114, 9)) ' FBASE1
                outputData(outRow, 17) = Trim(Mid(cleanL1, 123, 6)) ' EXPIRACION
                
                ' Line 2 Fields
                outputData(outRow, 18) = Trim(Mid(cleanL2, 1, 12))  ' TERMINAL
                outputData(outRow, 19) = Trim(Mid(cleanL2, 13, 5))  ' TIPO
                outputData(outRow, 20) = Trim(Mid(cleanL2, 18, 15)) ' IDENTIFICACION
                outputData(outRow, 21) = Trim(Mid(cleanL2, 33, 26)) ' ESTABLECIMIENTO
                outputData(outRow, 22) = Trim(Mid(cleanL2, 59, 14)) ' CIUDAD
                outputData(outRow, 23) = Trim(Mid(cleanL2, 73, 6))  ' PAIS
                outputData(outRow, 24) = Trim(Mid(cleanL2, 79, 13)) ' BIN ADQUIR
                outputData(outRow, 25) = Trim(Mid(cleanL2, 92, 5))  ' PIN
                outputData(outRow, 26) = Trim(Mid(cleanL2, 97, 12)) ' VIS.REFER
                outputData(outRow, 27) = Trim(Mid(cleanL2, 109, 5)) ' TRNX
                outputData(outRow, 28) = Trim(Mid(cleanL2, 114, 6)) ' CAVV
                outputData(outRow, 29) = Trim(Mid(cleanL2, 120, 21)) ' POS.C.CODE
                
            End If
            
            hasPendingLine1 = False ' Reset pair
        End If
        
NextLine:
    Loop
    
    ts.Close
    
    ' 4. Dump to Sheet
    If outRow > 0 Then
        ws.Range("A2").Resize(outRow, UBound(headers) + 1).Value = outputData
        MsgBox "Success! Parsed " & outRow & " records.", vbInformation
    Else
        MsgBox "No records found.", vbExclamation
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Function CleanImporte(val As String) As String
    ' Remove non-numeric chars except . and -
    Dim res As String
    Dim c As String
    Dim i As Long
    
    For i = 1 To Len(val)
        c = Mid(val, i, 1)
        If IsNumeric(c) Or c = "." Or c = "-" Then
            res = res & c
        End If
    Next i
    
    If res = "" Then
        CleanImporte = "0"
    Else
        CleanImporte = res
    End If
End Function
