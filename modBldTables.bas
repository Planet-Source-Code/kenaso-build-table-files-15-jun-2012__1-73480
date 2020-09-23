Attribute VB_Name = "modBldTables"
' ***************************************************************************
' Routine:       modBldTables
'
' Description:   Builds input tables used by GOST or Skipjack encryption
'                classes.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 16-Aug-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 16-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Fixed formatting bug in GostTables() routine
' 16-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated CenterReportText() routine
' 03-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Rewrote SkipjackTables() routine
'              - Tweaked WhirlpoolTables() routine
'              - Added reference to clsKeyEdit.cls for all three table
'                routines.
' 23-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Updated number of mixing iterations available
'              - Updated report titles
' 20-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Updated number of mixing iterations available
'              - Updated documentation
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const MODULE_NAME   As String = "modBldTables"
  Private Const TARGET_FOLDER As String = "Tables"
  Private Const MIN           As Long = 100   ' Minimum number of mixing loops
  Private Const MAX           As Long = 200   ' Maximum number of mixing loops
  Private Const MAX_BYTE      As Long = 256   ' Maximum number of bytes available (0-255)
  Private Const LINE_LENGTH   As Long = 74    ' Center data within this range
  Private Const SPACES        As Long = 100   ' Blank spaces for output record

' ***************************************************************************
' Enumerations
' ***************************************************************************
  Public Enum enumCASE_STRUCTURES
      eHorizontal_01   ' 0
      eHorizontal_02   ' 1
      eHorizontal_04   ' 2
      eHorizontal_05   ' 3
      eHorizontal_10   ' 4
      eVertical_02     ' 5
      eVertical_04     ' 6
      eVertical_05     ' 7
      eVertical_10     ' 8
      eVertical_20     ' 9
  End Enum

Public Sub GostTables(ByRef frmName As Form, _
             Optional ByVal lngMaxTables As Long = 1, _
             Optional ByVal lngCaseStruct As enumCASE_STRUCTURES = eHorizontal_01, _
             Optional ByVal strOptionalTitle As String = vbNullString)
                             
    Dim strTemp     As String
    Dim strRecord   As String
    Dim strOutput   As String
    Dim strPrefix   As String
    Dim strTarget   As String
    Dim hFile       As Long   ' File handle
    Dim lngRow      As Long
    Dim lngIndex    As Long
    Dim lngCaseCnt  As Long
    Dim lngPointer  As Long
    Dim lngMixCount As Long
    Dim lngTableCnt As Long
    Dim abytData()  As Byte
    Dim avntCase    As Variant
    Dim objPrng     As cPrng
    Dim objKeyEdit  As cKeyEdit
    
    Const RPT_TITLE    As String = "Gost S-box Table Sets"
    Const TARGET_FILE  As String = "Gost_tbl.txt"
    Const ROUTINE_NAME As String = "GostTables"

    On Error GoTo GostTables_Error

    If lngMaxTables < 1 Then
        InfoMsg "Number of tables to create must be greater than zero." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Sub
    End If
    
    ' Format path of destination folder
    strTarget = QualifyPath(App.Path) & TARGET_FOLDER
    
    ' Create folder if does not exist
    If Not IsPathValid(strTarget) Then
        MkDir strTarget
    End If
    
    ' Format complete path of destination file
    strTarget = QualifyPath(strTarget) & TARGET_FILE
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo GostTables_CleanUp
    End If
    
    Screen.MousePointer = vbHourglass   ' Change mouse pointer to hourglass
    DoEvents
    
    avntCase = Empty    ' Always start with empty variants
    Erase abytData()    ' Always start with empty arrays
    
    ReDim abytData(16)  ' Size array
    lngRow = 0          ' Init counters
    lngCaseCnt = 0
    lngTableCnt = 1
    strOutput = vbNullString  ' start with empty output record
    
    '                                  |--------------- 26th position (Insert row number)
    strPrefix = Space$(16) & "avntData(x) = Array("   ' Prepare prefix for case data
    avntCase = LoadCaseStatements(lngCaseStruct)      ' Load case statement formats
    
    DoEvents
    CloseAllFiles                         ' Close any open files
    hFile = FreeFile                      ' Get first free file handle
    Open strTarget For Output As #hFile   ' Create an empty receiving file
    
    Set objPrng = New cPrng               ' Instantiate class objects
    Set objKeyEdit = New cKeyEdit
    
    With objKeyEdit
        
        strOutput = strOutput & String$(LINE_LENGTH, 61) & vbNewLine
        strOutput = strOutput & .CenterReportText(LINE_LENGTH, RPT_TITLE, _
                                                  Format$(Now(), "d MMM yyyy"), _
                                                  FormatDateTime(Now(), vbLongTime)) & vbNewLine
        
        ' Insert optional title line if there is some data
        If Len(Trim$(strOptionalTitle)) > 0 Then
            strOutput = strOutput & .CenterReportText(LINE_LENGTH, strOptionalTitle) & vbNewLine
        End If
        
    End With
    
    Set objKeyEdit = Nothing   ' Free class object from memory (no longer needed)
        
    strOutput = strOutput & String$(LINE_LENGTH, 61) & vbNewLine
    strOutput = strOutput & "GOST S-Box values are 0 to 15 in random order with no duplicates." & vbNewLine
    strOutput = strOutput & IIf(lngMaxTables = 1, "Below is ", "Below are ") & CStr(lngMaxTables)
    strOutput = strOutput & IIf(lngMaxTables = 1, " table set.", " table sets from which to choose.")
    strOutput = strOutput & vbNewLine & vbNewLine
    
    strOutput = strOutput & "Each data line:" & vbNewLine
    strOutput = strOutput & Space$(4) & "1.  Consists of values (0-15) with no duplicates" & vbNewLine
    strOutput = strOutput & Space$(4) & "2.  Mixed 100-200 iterations using Knuth Shuffle" & vbNewLine
    strOutput = strOutput & String$(LINE_LENGTH, 61) & vbNewLine
    
    Print #hFile, strOutput   ' Write title data to output file
    strOutput = vbNullString  ' Verify output record is empty

    ' Print table heading
    Print #hFile, "Table set no. " & CStr(lngTableCnt)
    Print #hFile, Space$(4) & "Select Case mlngKeyLength"
    Print #hFile, Space$(11) & "Case " & CStr(avntCase(lngCaseCnt))
    
    frmName.lblCount(0).Caption = CStr(lngTableCnt)   ' Update table count on form
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo GostTables_CleanUp
    End If
    
    ' Load data array (0-15). Convert to
    ' smaller memory storage (1 byte <- 4 bytes)
    For lngIndex = 0 To 15
        abytData(lngIndex) = CByte(lngIndex)
    Next lngIndex
    
    With objPrng
        
        Rnd -1               ' Reset VB Random Number Generator
        Randomize .RndSeed   ' Reseed VB Random Number Generator
        
        ' Start with mixed data
        lngMixCount = Val(Int(Rnd() * (MAX - MIN + 1)) + MIN)   ' Create mixing number (100-200)
        .ReshuffleData abytData(), lngMixCount                  ' Mix array data
        
        Do
            DoEvents
            lngMixCount = Val(Int(Rnd() * (MAX - MIN + 1)) + MIN)   ' Create mixing number (100-200)
            .ReshuffleData abytData(), lngMixCount                  ' Mix array data
            lngPointer = 1                                          ' Set starting position for output record
            strTemp = vbNullString                                  ' Start with empty holding areas
            strRecord = vbNullString
            strOutput = Space$(55)                                  ' Preload output record with blank spaces
            
            ' Start formatting output string
            ' Ex:  "9, 10, 4, 5, 2, 13, 1, 0, 7, 11, 3, 12, 14, 15, 8, 6,  "
            For lngIndex = 0 To 15
                        
                strTemp = CStr(abytData(lngIndex)) & ", "            ' Append comma and blank space
                Mid$(strOutput, lngPointer, Len(strTemp)) = strTemp  ' Insert into output string
                lngPointer = lngPointer + Len(strTemp)               ' Increment string position
                strTemp = vbNullString                               ' Verify holding area is empty
                
            Next lngIndex
            
            ' Capture required data without
            ' trailing comma and blank space
            ' Ex:  "9, 10, 4, 5, 2, 13, 1, 0, 7, 11, 3, 12, 14, 15, 8, 6"
            strOutput = Left$(strOutput, 52)
                        
            ' Final format    |--- Row number
            ' Ex:   "avntData(0) = Array(9, 10, 4, 5, 2, 13, 1, 0, 7, 11, 3, 12, 14, 15, 8, 6)"
            Mid$(strPrefix, 26, 1) = CStr(lngRow)     ' Insert row number (0-7)
            strRecord = strPrefix & strOutput & ")"   ' Concatenate prefix, data and trailing parenthesis
            
            Print #hFile, strRecord   ' Write to output file
            strOutput = vbNullString  ' Verify output string is empty
            strRecord = vbNullString  ' Verify record string is empty
            lngRow = lngRow + 1       ' Increment row counter
        
            ' See if eight rows of
            ' data have been created
            If lngRow = 8 Then
            
                lngRow = 0                   ' Reset row counter
                lngCaseCnt = lngCaseCnt + 1  ' Increment case statement counter
                
                If lngCaseCnt > UBound(avntCase) Then
                    
                    Print #hFile, Space$(4) & "End Select"   ' write end of table data
                    Print #hFile, ""                         ' write blank line
                
                    ' See if required number of
                    ' tables have been created
                    Select Case lngTableCnt
                    
                           Case lngMaxTables   ' Time to leave
                                Print #hFile, ""                      ' write blank line
                                Print #hFile, "--- End of file ---"   ' Last file entry
                                Screen.MousePointer = vbDefault       ' Return mouse pointer to normal
                                Exit Do                               ' exit Do..Loop
                    
                           Case Else   ' Create another table
                                lngCaseCnt = 0                                    ' Reset case statement counter
                                lngTableCnt = lngTableCnt + 1                     ' Increment table counter
                                frmName.lblCount(0).Caption = CStr(lngTableCnt)   ' Update table count on form
                        
                                ' Write title data for new table
                                Print #hFile, "Table no. " & CStr(lngTableCnt)
                                Print #hFile, Space$(4) & "Select Case mlngKeyLength"
                                Print #hFile, Space$(11) & "Case " & CStr(avntCase(lngCaseCnt))
                    End Select
                    
                Else
                    ' Write next Case statement to file
                    Print #hFile, ""
                    Print #hFile, Space$(11) & "Case " & CStr(avntCase(lngCaseCnt))
                End If
            
            End If
            
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do
            End If
    
        Loop
    End With
    
GostTables_CleanUp:
    CloseAllFiles      ' Close any open files
    Erase abytData()   ' Always empty arrays when not needed
    avntCase = Empty   ' Always empty variants when not needed
    
    If Not objPrng Is Nothing Then
        objPrng.StopProcessing = gblnStopProcessing  ' Set abort flag
    End If
    
    Set objPrng = Nothing    ' Free class objects from memory
    Set objKeyEdit = Nothing
            
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        hFile = FreeFile                      ' Get first free file handle
        Open strTarget For Output As #hFile   ' empty output file
        Print #hFile, vbNewLine & "An error occurred or user opted to STOP processing"
        Close #hFile                          ' Close file
    End If
    
    Screen.MousePointer = vbDefault  ' Return mouse pointer to normal
    DoEvents
    
    On Error GoTo 0                  ' nullify this error trap
    Exit Sub

GostTables_Error:
    Screen.MousePointer = vbDefault   ' Return mouse pointer to normal
    DoEvents
    
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume GostTables_CleanUp

End Sub

Public Sub SkipjackTables(ByRef frmName As Form, _
                 Optional ByVal lngMaxTables As Long = 1, _
                 Optional ByVal lngCaseStruct As enumCASE_STRUCTURES = eHorizontal_01, _
                 Optional ByVal strOptionalTitle As String = vbNullString)
    
    Dim hFile       As Long   ' File handle
    Dim lngRow      As Long
    Dim lngLoop     As Long
    Dim lngIndex    As Long
    Dim lngTotal    As Long
    Dim lngCaseCnt  As Long
    Dim lngPointer  As Long
    Dim lngMixCount As Long
    Dim lngTableCnt As Long
    Dim strPrefix   As String
    Dim strOutput   As String
    Dim strRecord   As String
    Dim strTarget   As String
    Dim astrData()  As String
    Dim avntCase    As Variant
    Dim objPrng     As cPrng
    Dim objKeyEdit  As cKeyEdit
    
    Const RPT_TITLE    As String = "Skipjack Input Sets"
    Const TARGET_FILE  As String = "Skipjack_tbl.txt"
    Const ROUTINE_NAME As String = "SkipjackTables"

    On Error GoTo SkipjackTables_Error

    If lngMaxTables < 1 Then
        InfoMsg "Number of tables to create must be greater than zero." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Sub
    End If
    
    ' Format path of destination folder
    strTarget = QualifyPath(App.Path) & TARGET_FOLDER
    
    ' Create folder if does not exist
    If Not IsPathValid(strTarget) Then
        MkDir strTarget
    End If
    
    ' Format complete path of destination file
    strTarget = QualifyPath(strTarget) & TARGET_FILE
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo SkipjackTables_CleanUp
    End If
    
    Screen.MousePointer = vbHourglass   ' Change mouse pointer to hourglass
    DoEvents
        
    avntCase = Empty           ' Always start with empty variants
    Erase astrData()           ' Always start with empty arrays
        
    ReDim astrData(MAX_BYTE)   ' Size data array
    lngTotal = 0               ' total number of mixing iterations
    lngCaseCnt = 0             ' Set case statement counter
    lngTableCnt = 1            ' Set table counter
    strOutput = vbNullString   ' start with empty output record
    strRecord = vbNullString   ' start with empty data record
    
    strPrefix = Space$(16) & "strData = strData & " & Chr$(34)   ' Prepare prefix for case data
    avntCase = LoadCaseStatements(lngCaseStruct)                 ' Load case statement formats
    
    DoEvents
    CloseAllFiles                         ' Close any open files
    hFile = FreeFile                      ' Capture first free file handle
    Open strTarget For Output As #hFile   ' Create an empty file
    
    Set objPrng = New cPrng               ' Instantiate class objects
    Set objKeyEdit = New cKeyEdit
    
    With objKeyEdit
       
       strOutput = strOutput & String$(LINE_LENGTH, 61) & vbNewLine
       strOutput = strOutput & .CenterReportText(LINE_LENGTH, RPT_TITLE, _
                                                 Format$(Now(), "d MMM yyyy"), _
                                                 FormatDateTime(Now(), vbLongTime)) & vbNewLine
       
       ' Insert optional title line if there is some data
       If Len(Trim$(strOptionalTitle)) > 0 Then
           strOutput = strOutput & .CenterReportText(LINE_LENGTH, strOptionalTitle) & vbNewLine
       End If
       
    End With
    
    Set objKeyEdit = Nothing   ' Free class object from memory (no longer needed)
       
    strOutput = strOutput & String$(LINE_LENGTH, 61) & vbNewLine
    strOutput = strOutput & "Skipjack input sets for alternate loading.  "
    strOutput = strOutput & IIf(lngMaxTables = 1, "Below is ", "Below are ") & CStr(lngMaxTables)
    strOutput = strOutput & IIf(lngMaxTables = 1, " table set.", " table sets from") & vbNewLine
    strOutput = strOutput & IIf(lngMaxTables = 1, "", "which to choose." & vbNewLine)
    
    Print #hFile, strOutput   ' Write title data to output file
    strOutput = vbNullString  ' Verify output record is empty
    
    strOutput = strOutput & "Each case statement:" & vbNewLine
    strOutput = strOutput & Space$(4) & "1.  Consists of all ASCII values (0-255) in two" & vbNewLine
    strOutput = strOutput & Space$(8) & "character hex equivalent with no duplicates" & vbNewLine
    strOutput = strOutput & Space$(4) & "2.  Mixed 100-600 iterations using Knuth Shuffle" & vbNewLine
    strOutput = strOutput & String$(LINE_LENGTH, 61) & vbNewLine
    
    Print #hFile, strOutput   ' Write title data to output file
    strOutput = vbNullString  ' Verify output record is empty
    
    ' Print table heading
    Print #hFile, "Table set no. " & CStr(lngTableCnt)
    Print #hFile, Space$(4) & "Select Case mlngKeyLength"
    
    frmName.lblCount(1).Caption = CStr(lngTableCnt)   ' Update table count on form
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo SkipjackTables_CleanUp
    End If
        
    ' Load array with ASCII values 0-255
    ' converted to its hex equivalent
    For lngIndex = 0 To MAX_BYTE - 1
        astrData(lngIndex) = UCase$(Right$("00" & Hex$(lngIndex), 2))
    Next lngIndex
    
    With objPrng
        
        Rnd -1               ' Reset VB Random Number Generator
        Randomize .RndSeed   ' Reseed VB Random Number Generator
        
        ' Start with mixed data
        lngMixCount = Val(Int(Rnd() * (MAX - MIN + 1)) + MIN)   ' Create mixing number (100-200)
        .ReshuffleData astrData(), lngMixCount                  ' Mix array data
        
        Do
            lngLoop = Val(Int(Rnd() * 3) + 1)   ' number of loops (1-3)
            
            '----------------------------------------------------------------
            ' Uncomment next 3 Debug statements to display results in the
            ' immediate window.  Press CTRL+G to open the immediate window.
            ' Semi-colon at end of debug line means to append to this line.
            ' Same concept used to print to a sequential text file.
            '
            ' Ex:  Loops: 1  Mix count: 168 = 168
            '      Loops: 2  Mix count: 194 144 = 338
            '      Loops: 3  Mix count: 187 145 197 = 529
            '
        'Debug.Print "Loops:" & Format$(lngLoop, "@@") & "  Mix count:";
            For lngIndex = 1 To lngLoop
                
                lngMixCount = Val(Int(Rnd() * (MAX - MIN + 1)) + MIN)   ' Create mixing number (100-200)
        'Debug.Print Format$(lngMixCount, "@@@@");
                .ReshuffleData astrData(), lngMixCount                  ' Mix array data
                lngTotal = lngTotal + lngMixCount                       ' Increment total accumulator
            
            Next lngIndex
            
        'Debug.Print " = " & CStr(lngTotal)    ' Print total and start a new line
            lngTotal = 0
            '----------------------------------------------------------------
            
            ' Ex:  Case 128, 160, 192, 224
            Print #hFile, Space$(11) & "Case " & avntCase(lngCaseCnt)
            
            strRecord = Space$(770)   ' Preload data record with spaces
            lngPointer = 1            ' Set starting pointer for data record
                         
            ' load data record completely
            For lngIndex = 0 To (MAX_BYTE - 1)
                Mid$(strRecord, lngPointer, 2) = astrData(lngIndex)
                lngPointer = lngPointer + 3    ' Increment pointer (2 chars + blank space)
            Next lngIndex
            
            lngPointer = 1   ' Reset starting pointer
            
            ' load output data rows
            For lngIndex = 1 To 8
                
                strOutput = vbNullString   ' Verify output record is empty
                
                ' Add prefix data and append one double quote to output record
                ' ex:  strData = strData & "A0 E1 7C 36 C1 38 A5 ... D6 C6 7A 2D CB 1C AF "
                strOutput = strPrefix & Mid$(strRecord, lngPointer, 96) & Chr$(34)
                lngPointer = lngPointer + 96   ' Increment pointer
                Print #hFile, strOutput        ' Write data to output file
                                
            Next lngIndex
            
            strOutput = vbNullString   ' Verify output string is empty
            strRecord = vbNullString   ' Verify record string is empty
                            
            ' Test for end of case statements
            If lngCaseCnt = UBound(avntCase) Then
                                
                Print #hFile, Space$(4) & "End Select"   ' write end of table data
                Print #hFile, ""                         ' write blank line
                
                ' See if required number of
                ' tables have been created
                Select Case lngTableCnt
                
                       Case lngMaxTables   ' Time to leave
                            Print #hFile, ""                      ' write blank line
                            Print #hFile, "--- End of file ---"   ' Last file entry
                            Screen.MousePointer = vbDefault       ' Return mouse pointer to normal
                            Exit Do                               ' exit Do..Loop
                
                       Case Else   ' Create another table
                            lngCaseCnt = 0                                    ' Reset case statement counter
                            lngTableCnt = lngTableCnt + 1                     ' Increment table counter
                            frmName.lblCount(1).Caption = CStr(lngTableCnt)   ' Update table count on form
                    
                            ' Write title data for new table
                            Print #hFile, "Table set no. " & CStr(lngTableCnt)
                            Print #hFile, Space$(4) & "Select Case mlngKeyLength"
                End Select

            Else
                lngCaseCnt = lngCaseCnt + 1   ' Increment case statement count
                Print #hFile, ""              ' write blank line
            End If
        
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do
            End If
    
        Loop
    End With
    
SkipjackTables_CleanUp:
    CloseAllFiles      ' Close any open files
    Erase astrData()   ' Always empty arrays when not needed
    avntCase = Empty   ' Always empty variants when not needed
    
    If Not objPrng Is Nothing Then
        objPrng.StopProcessing = gblnStopProcessing  ' Set abort flag
    End If
    
    Set objPrng = Nothing    ' Free class objects from memory
    Set objKeyEdit = Nothing
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        hFile = FreeFile                      ' Get first free file handle
        Open strTarget For Output As #hFile   ' empty output file
        Print #hFile, vbNewLine & "An error occurred or user opted to STOP processing"
        Close #hFile                          ' Close file
    End If
    
    Screen.MousePointer = vbDefault  ' Return mouse pointer to normal
    DoEvents
    
    On Error GoTo 0                  ' nullify this error trap
    Exit Sub
    
SkipjackTables_Error:
    Screen.MousePointer = vbDefault  ' Return mouse pointer to normal
    DoEvents
    
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume SkipjackTables_CleanUp

End Sub

Public Sub WhirlpoolTables(ByRef frmName As Form, _
                  Optional ByVal lngMaxTables As Long = 1, _
                  Optional ByVal strOptionalTitle As String = vbNullString)
    
    Dim hFile       As Long   ' File handle
    Dim lngRow      As Long
    Dim lngLoop     As Long
    Dim lngIndex    As Long
    Dim lngTotal    As Long
    Dim lngCaseCnt  As Long
    Dim lngPointer  As Long
    Dim lngMixCount As Long
    Dim lngTableCnt As Long
    Dim strPrefix   As String
    Dim strOutput   As String
    Dim strRecord   As String
    Dim strTarget   As String
    Dim astrData()  As String
    Dim astrCase()  As String
    Dim objPrng     As cPrng
    Dim objKeyEdit  As cKeyEdit
    
    Const RPT_TITLE    As String = "Whirlpool Input Sets"
    Const TARGET_FILE  As String = "Whirlpool_tbl.txt"
    Const ROUTINE_NAME As String = "WhirlpoolTables"

    On Error GoTo WhirlpoolTables_Error

    If lngMaxTables < 1 Then
        InfoMsg "Number of tables to create must be greater than zero." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Sub
    End If
     
    ' Format path of destination folder
    strTarget = QualifyPath(App.Path) & TARGET_FOLDER
    
    ' Create folder if does not exist
    If Not IsPathValid(strTarget) Then
        MkDir strTarget
    End If
    
    ' Format complete path of destination file
    strTarget = QualifyPath(strTarget) & TARGET_FILE
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo WhirlpoolTables_CleanUp
    End If
    
    Screen.MousePointer = vbHourglass   ' Change mouse pointer to hourglass
    DoEvents
    
    ReDim astrCase(0 To 2)    ' Size Case array
    ReDim astrData(MAX_BYTE)  ' Size data array
    
    lngTotal = 0              ' total number of mixing iterations
    lngCaseCnt = 0            ' Set case statement counter
    lngTableCnt = 1           ' Set table counter
    strOutput = vbNullString  ' start with empty output record
    strRecord = vbNullString  ' start with empty data record
    
    strPrefix = Space$(16) & "strData = strData & " & Chr$(34)  ' Prepare prefix for case data
    
    astrCase(0) = "eWHIRLPOOL224"   ' Load case statements
    astrCase(1) = "eWHIRLPOOL256"
    astrCase(2) = "eWHIRLPOOL384"
    
    DoEvents
    CloseAllFiles                         ' Close any open files
    hFile = FreeFile                      ' Capture first free file handle
    Open strTarget For Output As #hFile   ' Create an empty file
    
    Set objPrng = New cPrng               ' Instantiate class objects
    Set objKeyEdit = New cKeyEdit
    
    With objKeyEdit
       
       strOutput = strOutput & String$(LINE_LENGTH, 61) & vbNewLine
       strOutput = strOutput & .CenterReportText(LINE_LENGTH, RPT_TITLE, _
                                                 Format$(Now(), "d MMM yyyy"), _
                                                 FormatDateTime(Now(), vbLongTime)) & vbNewLine
       
       ' Insert optional title line if there is some data
       If Len(Trim$(strOptionalTitle)) > 0 Then
           strOutput = strOutput & .CenterReportText(LINE_LENGTH, strOptionalTitle) & vbNewLine
       End If
       
    End With
    
    Set objKeyEdit = Nothing   ' Free class object from memory (no longer needed)
       
    strOutput = strOutput & String$(LINE_LENGTH, 61) & vbNewLine
    strOutput = strOutput & "Whirlpool input sets for alternate loading.  "
    strOutput = strOutput & IIf(lngMaxTables = 1, "Below is ", "Below are ") & CStr(lngMaxTables)
    strOutput = strOutput & IIf(lngMaxTables = 1, " table set.", " table sets from") & vbNewLine
    strOutput = strOutput & IIf(lngMaxTables = 1, "", "which to choose." & vbNewLine)
    
    Print #hFile, strOutput   ' Write title data to output file
    strOutput = vbNullString  ' Verify output record is empty
    
    strOutput = strOutput & "Each case statement:" & vbNewLine
    strOutput = strOutput & Space$(4) & "1.  Consists of all ASCII values (0-255) in two" & vbNewLine
    strOutput = strOutput & Space$(8) & "character hex equivalent with no duplicates" & vbNewLine
    strOutput = strOutput & Space$(4) & "2.  Mixed 100-600 iterations using Knuth Shuffle" & vbNewLine
    strOutput = strOutput & String$(LINE_LENGTH, 61) & vbNewLine
    
    Print #hFile, strOutput   ' Write title data to output file
    strOutput = vbNullString  ' Verify output record is empty
    
    ' Print table heading
    Print #hFile, "Table set no. " & CStr(lngTableCnt)
    Print #hFile, Space$(4) & "Select Case mlngHashMethod"
    
    frmName.lblCount(2).Caption = CStr(lngTableCnt)   ' Update table count on form
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo WhirlpoolTables_CleanUp
    End If
        
    ' Load array with ASCII values 0-255
    ' converted to its hex equivalent
    For lngIndex = 0 To MAX_BYTE - 1
        astrData(lngIndex) = UCase$(Right$("00" & Hex$(lngIndex), 2))
    Next lngIndex
    
    With objPrng
        
        Rnd -1               ' Reset VB Random Number Generator
        Randomize .RndSeed   ' Reseed VB Random Number Generator
        
        ' Start with mixed data
        lngMixCount = Val(Int(Rnd() * (MAX - MIN + 1)) + MIN)   ' Create mixing number (100-200)
        .ReshuffleData astrData(), lngMixCount                  ' Mix array data
        
        Do
            lngLoop = Val(Int(Rnd() * 3) + 1)   ' number of loops (1-3)
            
            '----------------------------------------------------------------
            ' Uncomment next 3 Debug statements to display results in the
            ' immediate window.  Press CTRL+G to open the immediate window.
            ' Semi-colon at end of debug line means to append to this line.
            ' Same concept used to print to a sequential text file.
            '
            ' Ex:  Loops: 1  Mix count: 168 = 168
            '      Loops: 2  Mix count: 194 144 = 338
            '      Loops: 3  Mix count: 187 145 197 = 529
            '
        'Debug.Print "Loops:" & Format$(lngLoop, "@@") & "  Mix count:";
            For lngIndex = 1 To lngLoop
                
                lngMixCount = Val(Int(Rnd() * (MAX - MIN + 1)) + MIN)   ' Create mixing number (100-200)
        'Debug.Print Format$(lngMixCount, "@@@@");
                .ReshuffleData astrData(), lngMixCount                  ' Mix array data
                lngTotal = lngTotal + lngMixCount                       ' Increment total accumulator
            
            Next lngIndex
            
        'Debug.Print " = " & CStr(lngTotal)    ' Print total and start a new line
            lngTotal = 0
            '----------------------------------------------------------------
            
            ' Ex:  Case eWHIRLPOOL224
            Print #hFile, Space$(11) & "Case " & astrCase(lngCaseCnt)
            
            strRecord = Space$(770)   ' Preload data record with spaces
            lngPointer = 1            ' Set starting pointer for data record
                         
            ' load Temp record completely
            For lngIndex = 0 To (MAX_BYTE - 1)
                Mid$(strRecord, lngPointer, 2) = astrData(lngIndex)
                lngPointer = lngPointer + 3    ' Increment pointer (2 chars + blank space)
            Next lngIndex
            
            lngPointer = 1   ' Reset starting pointer
            
            ' load output data rows
            For lngIndex = 1 To 8
                
                strOutput = vbNullString   ' Verify output record is empty
                
                ' Add prefix data and append one double quote to output record
                ' ex:  strData = strData & "A0 E1 7C 36 C1 38 A5 ... D6 C6 7A 2D CB 1C AF "
                strOutput = strPrefix & Mid$(strRecord, lngPointer, 96) & Chr$(34)
                lngPointer = lngPointer + 96   ' Increment pointer
                Print #hFile, strOutput        ' Write data to output file
                                
            Next lngIndex
            
            strOutput = vbNullString   ' Verify output string is empty
            strRecord = vbNullString   ' Verify record string is empty
                                                        
            ' Test for end of case statements in Case array
            If lngCaseCnt = UBound(astrCase) Then
                                
                ' THIS MUST NOT BE MODIFIED!
                ' Last case statement is Whirlpool's default data
                Print #hFile, ""                         ' write blank line
                Print #hFile, Space$(11) & "Case eWHIRLPOOL512   ' Original data - DO NOT MODIFY"
                Print #hFile, strPrefix & "18 23 C6 E8 87 B8 01 4F 36 A6 D2 F5 79 6F 91 52 60 BC 9B 8E A3 0C 7B 35 1D E0 D7 C2 2E 4B FE 57 " & Chr$(34)
                Print #hFile, strPrefix & "15 77 37 E5 9F F0 4A DA 58 C9 29 0A B1 A0 6B 85 BD 5D 10 F4 CB 3E 05 67 E4 27 41 8B A7 7D 95 D8 " & Chr$(34)
                Print #hFile, strPrefix & "FB EE 7C 66 DD 17 47 9E CA 2D BF 07 AD 5A 83 33 63 02 AA 71 C8 19 49 D9 F2 E3 5B 88 9A 26 32 B0 " & Chr$(34)
                Print #hFile, strPrefix & "E9 0F D5 80 BE CD 34 48 FF 7A 90 5F 20 68 1A AE B4 54 93 22 64 F1 73 12 40 08 C3 EC DB A1 8D 3D " & Chr$(34)
                Print #hFile, strPrefix & "97 00 CF 2B 76 82 D6 1B B5 AF 6A 50 45 F3 30 EF 3F 55 A2 EA 65 BA 2F C0 DE 1C FD 4D 92 75 06 8A " & Chr$(34)
                Print #hFile, strPrefix & "B2 E6 0E 1F 62 D4 A8 96 F9 C5 25 59 84 72 39 4C 5E 78 38 8C D1 A5 E2 61 B3 21 9C 1E 43 C7 FC 04 " & Chr$(34)
                Print #hFile, strPrefix & "51 99 6D 0D FA DF 7E 24 3B AB CE 11 8F 4E B7 EB 3C 81 94 F7 B9 13 2C D3 E7 6E C4 03 56 44 7F A9 " & Chr$(34)
                Print #hFile, strPrefix & "2A BB C1 53 DC 0B 9D 6C 31 74 F6 46 AC 89 14 E1 16 3A 69 09 70 B6 D0 ED CC 42 98 A4 28 5C F8 86 " & Chr$(34)
                Print #hFile, Space$(4) & "End Select"   ' write end of table data
                Print #hFile, ""                         ' write blank line
                
                ' See if required number of
                ' tables have been created
                Select Case lngTableCnt
                
                       Case lngMaxTables   ' Time to leave
                            Print #hFile, ""                      ' write blank line
                            Print #hFile, "--- End of file ---"   ' Last file entry
                            Screen.MousePointer = vbDefault       ' Return mouse pointer to normal
                            Exit Do                               ' exit Do..Loop
                
                       Case Else   ' Create another table
                            lngCaseCnt = 0                                    ' Reset case statement counter
                            lngTableCnt = lngTableCnt + 1                     ' Increment table counter
                            frmName.lblCount(2).Caption = CStr(lngTableCnt)   ' Update table count on form
                    
                            ' Write title data for new table
                            Print #hFile, "Table set no. " & CStr(lngTableCnt)
                            Print #hFile, Space$(4) & "Select Case mlngHashMethod"
                End Select

            Else
                lngCaseCnt = lngCaseCnt + 1   ' Increment case statement count
                Print #hFile, ""              ' Write blank line
            End If
        
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do
            End If
    
        Loop
    End With
    
WhirlpoolTables_CleanUp:
    CloseAllFiles      ' Close any open files
    Erase astrData()   ' Always empty arrays when not needed
    Erase astrCase()
    
    If Not objPrng Is Nothing Then
        objPrng.StopProcessing = gblnStopProcessing  ' Set abort flag
    End If
    
    Set objPrng = Nothing    ' Free class objects from memory
    Set objKeyEdit = Nothing

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        hFile = FreeFile                      ' Get first free file handle
        Open strTarget For Output As #hFile   ' empty output file
        Print #hFile, vbNewLine & "An error occurred or user opted to STOP processing"
        Close #hFile                          ' Close file
    End If
    
    Screen.MousePointer = vbDefault  ' Return mouse pointer to normal
    DoEvents
    
    On Error GoTo 0                  ' nullify this error trap
    Exit Sub
    
WhirlpoolTables_Error:
    Screen.MousePointer = vbDefault  ' Return mouse pointer to normal
    DoEvents
    
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume WhirlpoolTables_CleanUp

End Sub



' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function LoadCaseStatements(ByVal lngCaseStruct As enumCASE_STRUCTURES) As Variant
    
    ' Called by GostTables()
    '           SkipjackTables()
    
    LoadCaseStatements = Empty    ' Always start with empty variants
    
    ' Determine format of CASE statements
    Select Case lngCaseStruct
    
           ' Horizontal key length
           Case eHorizontal_01  ' One case statement
                LoadCaseStatements = Array("128, 160, 192, 224, 256, 288, 320, 352, 384, 416, _" & _
                                           vbNewLine & Space$(16) & _
                                           "448, 512, 576, 640, 704, 768, 832, 896, 960, 1024")
                                                
           Case eHorizontal_02  ' Two case statements
                LoadCaseStatements = Array("128, 160, 192, 224, 256, 288, 320, 352, 384, 416", _
                                           "448, 512, 576, 640, 704, 768, 832, 896, 960, 1024")
                                                
           Case eHorizontal_04  ' Four case statements
                LoadCaseStatements = Array("128, 160, 192, 224, 256", _
                                           "288, 320, 352, 384, 416", _
                                           "448, 512, 576, 640, 704", _
                                           "768, 832, 896, 960, 1024")
                                                
           Case eHorizontal_05  ' Five case statements
                LoadCaseStatements = Array("128, 160, 192, 224", _
                                           "256, 288, 320, 352", _
                                           "384, 416, 448, 512", _
                                           "576, 640, 704, 768", _
                                           "832, 896, 960, 1024")
                                                
           Case eHorizontal_10  ' Ten case statements
                LoadCaseStatements = Array("128, 160", _
                                           "192, 224", _
                                           "256, 288", _
                                           "320, 352", _
                                           "384, 416", _
                                           "448, 512", _
                                           "576, 640", _
                                           "704, 768", _
                                           "832, 896", _
                                           "960, 1024")
           ' Vertical key length
           Case eVertical_02    ' Two case statements
                LoadCaseStatements = Array("128, 192, 256, 320, 384, 448, 576, 704, 832, 960", _
                                           "160, 224, 288, 352, 416, 512, 640, 768, 896, 1024")
                                                
           Case eVertical_04    ' Four case statements
                LoadCaseStatements = Array("128, 256, 384, 576, 832", _
                                           "160, 288, 416, 640, 896", _
                                           "192, 320, 448, 704, 960", _
                                           "224, 352, 512, 768, 1024")
                                                
           Case eVertical_05    ' Five case statements
                LoadCaseStatements = Array("128, 288, 448, 768", _
                                           "160, 320, 512, 832", _
                                           "192, 352, 576, 896", _
                                           "224, 384, 640, 960", _
                                           "256, 416, 704, 1024")
                                                
           Case eVertical_10    ' Ten case statements
                LoadCaseStatements = Array("128, 448", _
                                           "160, 512", _
                                           "192, 576", _
                                           "224, 640", _
                                           "256, 704", _
                                           "288, 768", _
                                           "320, 832", _
                                           "352, 896", _
                                           "384, 960", _
                                           "416, 1024")
    
           Case eVertical_20    ' Twenty case statements
                LoadCaseStatements = Array("128", "192", "256", "320", "384", _
                                           "448", "576", "704", "832", "960", _
                                           "160", "224", "288", "352", "416", _
                                           "512", "640", "768", "896", "1024")
    End Select
    
End Function

