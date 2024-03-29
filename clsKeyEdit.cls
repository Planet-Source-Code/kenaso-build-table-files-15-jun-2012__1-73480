VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cKeyEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        cKeyEdit
'
' Description:   These are the common edit routines you will find in most
'                word processors.  (Copy, Cut, Paste)
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-JUL-1998  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 11-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated ProcessNumericOnly() routine flower box with an
'              additional code suggestion.
' 20-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Major update to CenterCaption() routine.
'              - Changed ErrorMsg call to InfoMsg call in CenterCaption()
'                and CenterReportText() routines.
' 25-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote CenterReportText() routine.
' 18-Feb-2011  Kenneth Ives  kenaso@tx.rr.com
'              Added NoCopyText() routine.
' 16-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated CenterReportText() routine
' 10-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug in CenterCaption() routine.
' 02-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed minor bugs.  See ProcessNumericOnly() and
'              ProcessAlphaNumeric() routines.
' ***************************************************************************
Option Explicit
  
' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME             As String = "clsKeyEdit"
  Private Const SPI_GETNONCLIENTMETRICS As Long = 41
  Private Const LOGPIXELSY              As Long = 90

' ***************************************************************************
' Type structures
' ***************************************************************************
  Private Type LogFont
      FontHeight          As Long
      FontWidth           As Long
      FontEscapement      As Long
      FontOrientation     As Long
      FontWeight          As Long
      FontItalic          As Byte
      FontUnderline       As Byte
      FontStrikeOut       As Byte
      FontCharSet         As Byte
      FontOutPrecision    As Byte
      FontClipPrecision   As Byte
      FontQuality         As Byte
      FontPitchAndFamily  As Byte
      FontFaceName        As String * 32
  End Type

  Private Type NONCLIENTMETRICS
      cbSize              As Long
      iBorderWidth        As Long
      iScrollWidth        As Long
      iScrollHeight       As Long
      iCaptionWidth       As Long
      iCaptionHeight      As Long
      LFCaptionFont       As LogFont
      iSMCaptionWidth     As Long
      iSMCaptionHeight    As Long
      LFSMCaptionFont     As LogFont
      iMenuWidth          As Long
      iMenuHeight         As Long
      LFMenuFont          As LogFont
      LFStatusFont        As LogFont
      LFMessageFont       As LogFont
  End Type

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' The GetSystemMetrics function retrieves various system metrics and
  ' system configuration settings.  System metrics are the dimensions
  ' (widths and heights) of Windows display elements. All dimensions
  ' retrieved by GetSystemMetrics are in pixels.
  Private Declare Function GetSystemMetrics Lib "user32" _
          (ByVal nIndex As Long) As Long

  ' The GetDeviceCaps function retrieves device-specific information
  ' about a specified device.
  Private Declare Function GetDeviceCaps Lib "gdi32" _
          (ByVal hDC As Long, ByVal nIndex As Long) As Long

  ' The SystemParametersInfo function queries or sets systemwide
  ' parameters. This function can also update the user profile while
  ' setting a parameter.  This function is intended for use with
  ' applications, such as Control Panel, that allow the user to
  ' customize the Windows environment.
  Private Declare Function SystemParametersInfo Lib "user32" _
          Alias "SystemParametersInfoA" (ByVal uAction As Long, _
          ByVal uParam As Long, lpvParam As Any, _
          ByVal fuWinIni As Long) As Long

  ' ZeroMemory is used for clearing contents of a type structure.
  Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" _
          (Destination As Any, ByVal Length As Long)


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       CenterCaption
'
' Description:   Centers a caption on a form.
'
' Parameters:    frmForm - Name of form whose caption is to be centered
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-OCT-2000  Tom Pydeski
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=12103&lngWId=1
' 20-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 10-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Exit routine if form caption is empty.
' ***************************************************************************
Public Sub CenterCaption(ByRef frmForm As Form)
                    
    Dim strCaption      As String
    Dim lngTwips        As Long
    Dim lngCtrlBox      As Long
    Dim lngTbarWidth    As Long
    Dim lngTbarHeight   As Long
    Dim lngCaptionWidth As Long
    
    On Error GoTo CenterCaption_Error
       
    Set frmForm.Font = GetTitleFont(frmForm)  ' Get font information
    strCaption = TrimStr(frmForm.Caption)     ' Capture form caption
    
    ' If no caption then leave
    If Len(strCaption) = 0 Then
        Exit Sub
    End If
    
    lngTbarHeight = GetSystemMetrics(4) * Screen.TwipsPerPixelX   ' height of window title bar area
    lngTbarWidth = GetSystemMetrics(30) * Screen.TwipsPerPixelX   ' width of button in window title bar area
    lngTwips = (frmForm.TextWidth(strCaption)) / Len(strCaption)  ' get average size of character in twips
    lngCtrlBox = (lngTbarWidth * 3) + lngTbarHeight               ' Calculate box space used on caption bar
    lngCaptionWidth = (frmForm.Width - lngCtrlBox) / lngTwips     ' Calculate total caption space
    strCaption = Space$(Fix(lngCaptionWidth / 1.5)) & strCaption  ' Format caption string
    
    frmForm.Caption = strCaption   ' Return modified caption
    
CenterCaption_CleanUp:
    On Error GoTo 0
    Exit Sub

CenterCaption_Error:
    Err.Clear
    Resume CenterCaption_CleanUp

End Sub

' ***************************************************************************
' Routine:       CenterReportText
'
' Description:   Center text on a line
'
' Parameters:    lngLineLength - Length of report line
'                strMiddle - Optional - String of data to be centered.
'                    (ex:  "My name and email")
'                    Default = vbNullString
'                strLeftSide - Optional - String of data to remain at left
'                    most end of output string.  (ex:  "25-Dec-2010")
'                    Default = vbNullString
'                strRightSide - Optional - String of data to remain at right
'                    most end of output string.  (ex:  "Page 1")
'                    Default = vbNullString
'
' Returns:       Formatted text
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 16-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              Added better parameter evaluation with error messages
' 11-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Made line length parameter mandatory
' ***************************************************************************
Public Function CenterReportText(ByVal lngLineLength As Long, _
                        Optional ByVal strMiddle As String = vbNullString, _
                        Optional ByVal strLeftSide As String = vbNullString, _
                        Optional ByVal strRightSide As String = vbNullString) As String
    
    Dim lngDataLength As Long
    
    Const ROUTINE_NAME As String = "CenterReportText"
    
    CenterReportText = vbNullString   ' Verify return string is empty
    
    ' If no line length, data cannot be centered
    If lngLineLength < 1 Then
        InfoMsg "Line length must be a positive number." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If

    ' Remove leading and trailing blank spaces
    strLeftSide = TrimStr(strLeftSide)
    strMiddle = TrimStr(strMiddle)
    strRightSide = TrimStr(strRightSide)

    ' Capture data length
    lngDataLength = Len(strLeftSide & strMiddle & strRightSide)
    
    Select Case lngDataLength

           Case Is < 1   ' If no data to process then leave
                InfoMsg "No data to process." & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                Exit Function

           Case is > lngLineLength   ' If too much data then leave
                InfoMsg "Line length must be equal to or greater than data length." & _
                        vbNewLine & "Line length:  " & CStr(lngLineLength) & _
                        vbNewLine & "Data length:  " & CStr(lngDataLength) & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                Exit Function
    End Select

    ' Add a blank space to beginning and end of
    ' middle string of data until line length
    ' requirement has been met or exceeded
    Do While Len(strMiddle) < lngLineLength
        strMiddle = Chr$(32) & strMiddle & Chr$(32)
    Loop
    
    ' Verify string length equals line length
    strMiddle = Left$(strMiddle, lngLineLength)
    
    ' If data is available then overlay left side
    If Len(strLeftSide) > 0 Then
        Mid$(strMiddle, 1, Len(strLeftSide)) = strLeftSide
    End If
    
    ' If data is available then overlay right side
    If Len(strRightSide) > 0 Then
        Mid$(strMiddle, (lngLineLength - Len(strRightSide)) + 1, Len(strRightSide)) = strRightSide
    End If
        
    ' Remove any excess trailing blanks because
    ' only leading blanks are needed to push data
    ' to middle of line.
    CenterReportText = RTrim$(strMiddle)
    
End Function

' ***************************************************************************
' Routine:       TextBoxKeyDown
'
' Description:   Processes a KeyDown() event for a textbox.  A combination
'                of keys that have been depressed in a TextBox
'
' Syntax:        Private Sub txtLength_KeyDown(KeyCode As Integer, Shift As Integer)
'                    ' Process any key combinations
'                    mobjKeyEdit.TextBoxKeyDown txtLength, KeyCode, Shift
'                End Sub
'
' Parameters:    ctlTextBox - TextBox control
'                intKeyCode - Possible special key being pressed (Ctl, Alt, etc)
'                intShift   - Possible shift key being pressed
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-APR-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub TextBoxKeyDown(ByRef ctlTextBox As TextBox, _
                          ByRef intKeyCode As Integer, _
                          ByRef intShift As Integer)

    Dim CtrlDown    As Integer
    Dim PressedKey  As Integer
    
    ' Initialize  variables
    CtrlDown = (intShift And vbCtrlMask) > 0      ' Define control key
    PressedKey = Asc(UCase$(Chr$(intKeyCode)))    ' Convert to uppercase
      
    ' Check to see if it is okay to make changes
    If CtrlDown And PressedKey = vbKeyX Then      ' Ctrl + X was pressed
        Edit_Cut ctlTextBox
        
    ElseIf CtrlDown And PressedKey = vbKeyA Then  ' Ctrl + A was pressed
        TextBoxFocus ctlTextBox
        
    ElseIf CtrlDown And PressedKey = vbKeyC Then  ' Ctrl + C was pressed
        Edit_Copy ctlTextBox
        
    ElseIf CtrlDown And PressedKey = vbKeyV Then  ' Ctrl + V was pressed
        Edit_Paste ctlTextBox
        
    ElseIf PressedKey = vbKeyDelete Then          ' Delete key was pressed
        Edit_Delete ctlTextBox
    End If

End Sub

' ***************************************************************************
' Routine:       TextBoxFocus
'
' Description:   Processes a GotFocus() event for a textbox.  Whenever the
'                textbox has been first entered then all the text within
'                will be highlighted.
'
' Syntax:        Private Sub txtLength_GotFocus()
'                    ' Highlight all the text in the box
'                    mobjKeyEdit.TextBoxFocus txtLength
'                End Sub
'
' Parameters:    ctlTextBox - TextBox control
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-APR-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub TextBoxFocus(ByRef ctlTextBox As TextBox)

    ' Called by TextBoxKeyDown()
    
    ' Highlight all the text in the box
    With ctlTextBox
         .SelStart = 0             ' Start with first character in TextBox
         .SelLength = Len(.Text)   ' To last character in TextBox
    End With
  
End Sub

' ***************************************************************************
' Routine:       NoCopyText
'
' Description:   Prevents text from being copied from a textbox.
'
' Syntax:        Private Sub txtLength_KeyDown(KeyCode As Integer, Shift As Integer)
'                    ' Protect data from being copied
'                    mobjKeyEdit.NoCopyText txtLength, KeyCode, Shift
'                End Sub
'
' Parameters:    ctlTextBox - TextBox control
'                intKeyCode - Possible special key being pressed (Ctl, Alt, etc)
'                intShift   - Possible shift key being pressed
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Feb-2011  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub NoCopyText(ByRef ctlTextBox As TextBox, _
                      ByRef intKeyCode As Integer, _
                      ByRef intShift As Integer)

    Dim CtrlDown   As Integer
    Dim PressedKey As Integer
    
    ' Define control key
    CtrlDown = (intShift And vbCtrlMask) > 0
    
    ' Convert keycode to uppercase then to decimal
    PressedKey = Asc(UCase$(Chr$(intKeyCode)))
      
    With ctlTextBox
        If (CtrlDown And PressedKey = vbKeyX) Or _
           (CtrlDown And PressedKey = vbKeyA) Or _
           (CtrlDown And PressedKey = vbKeyC) Or _
           (CtrlDown And PressedKey = vbKeyV) Then
            
            .SelStart = 0   ' Keep mouse pointer at beginning of data
            .SelLength = 0
            Clipboard.SetText vbNullString  ' capture an empty string
            Clipboard.Clear                 ' Empty clipboard
         End If
    End With
    
End Sub

' ***************************************************************************
' Routine:       ProcessNumericOnly
'
' Description:   Processes a KeyPress() event for a textbox.  Numeric only
'                data.  ASCII 9 (TAB), 13 (ENTER), 8 (BACKSPACE), 48-57 (0-9)
'
'
' Syntax:        Private Sub txtLength_KeyPress(KeyAscii As Integer)
'                    ' Save only numeric and the backspace character
'                    mobjKeyEdit.ProcessNumericOnly KeyAscii
'                End Sub
'
' Parameters:    intKeyASCII - key being pressed
'
' NOTE:          To prevent someone from pasting a non-numeric value
'                into this textbox, insert the code below into the
'                textbox_Change() event. Change "txtbox" to the name
'                of the textbox control.
'
'                ' Prevent user from pasting a non-numeric value
'                ' into this textbox
'                If Not IsNumeric(txtbox.Text) Then
'                    txtbox.Text = vbNullString
'                End If
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-APR-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 11-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added above note to prevent a user from pasting non-numeric
'              data into a textbox.  Got code from Masino Sinaga (04-Aug-2003)
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=47381&lngWId=1
' 02-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug.  ENTER key improperly recognized.  Now sends a
'              TAB key response.
' ***************************************************************************
Public Sub ProcessNumericOnly(ByRef intKeyASCII As Integer)
      
        Select Case intKeyASCII
               Case 9, 13
                    ' Tab key, ENTER key
                    intKeyASCII = 0
                    SendKeys "{TAB}"
                
               Case 8, 48 To 57
                    ' Backspace and numeric keys only
                
               Case Else  ' Everything else (invalid)
                    intKeyASCII = 0
        End Select
                              
End Sub

' ***************************************************************************
' Routine:       ProcessAlphaNumeric
'
' Description:   Processes a KeyPress() event for a textbox.  All keyboard
'                characters.  ASCII 32-126
'
' Syntax:        Private Sub txtLength_KeyPress(KeyAscii As Integer)
'                    ' Save only alphanumeric characters
'                    mobjKeyEdit.ProcessAlphaNumeric KeyAscii
'                End Sub
'
' Parameters:    intKeyASCII - key being pressed
'                blnAcceptSpaces - [Optional]
'                     TRUE=Accept blanks spaces as input (DEFAULT)
'                     FALSE=Do not aceept any blank spaces as input
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-APR-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 02-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug.  ENTER key improperly recognized.  Now enters a
'              Carriage Return and Linefeed.
' ***************************************************************************
Public Sub ProcessAlphaNumeric(ByRef intKeyASCII As Integer, _
                      Optional ByVal blnAcceptSpaces As Boolean = True)

    If blnAcceptSpaces Then
        Select Case intKeyASCII
               Case 9, 13
                    ' Tab key, ENTER key
                    intKeyASCII = 0
                    SendKeys "{TAB}"
                
               Case 8, 32 To 126
                    ' Backspace and other
                    ' valid data keys
                
               Case Else  ' Everything else (invalid)
                    intKeyASCII = 0
        End Select
    Else
        ' Do not accept blank spaces
        Select Case intKeyASCII
               Case 9, 13
                    ' Tab key, ENTER key
                    intKeyASCII = 0
                    SendKeys "{TAB}"
                
               Case 8, 33 To 126
                    ' Backspace and other
                    ' valid data keys
                
               Case Else  ' Everything else (invalid)
                    intKeyASCII = 0
        End Select
    End If

End Sub

' ***************************************************************************
' Routine:       ProcessAlphabetic
'
' Description:   Processes a KeyPress() event for a textbox.  Alphabetic
'                only data.  (A-Z, a-z)
'
' Syntax:        Private Sub txtLength_KeyPress(KeyAscii As Integer)
'                    ' Save only alphabetic only characters
'                    mobjKeyEdit.ProcessAlphabetic KeyAscii
'                End Sub
'
' Parameters:    intKeyASCII - key being pressed
'                blnUppercaseOnly - [Optional] - Allow only uppercase
'                    values to be entered.
'                    TRUE=Convert all letters to uppercase
'                    FALSE=Upper and lower case allowed (DEFAULT)
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-APR-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub ProcessAlphabetic(ByRef intKeyASCII As Integer, _
                    Optional ByVal blnUppercaseOnly As Boolean = False)
      
    If blnUppercaseOnly Then
        Select Case intKeyASCII
               Case 9, 13
                    ' Tab key, ENTER key
                    intKeyASCII = 0
                    SendKeys "{TAB}"
                
               Case 8, 65 To 90
                    ' Backspace and uppercase
                    ' alphabetic keys only
               
               Case 97 To 122
                    ' Convert lowercase to uppercase
                    intKeyASCII = intKeyASCII - 32
                    
               Case Else  ' Everything else (invalid)
                    intKeyASCII = 0
        End Select
    Else
        ' Case does not matter
        Select Case intKeyASCII
               Case 9, 13
                    ' Tab key, ENTER key
                    intKeyASCII = 0
                    SendKeys "{TAB}"
                
               Case 8, 65 To 90, 97 To 122
                    ' Backspace and alphabetic keys only

               Case Else  ' Everything else (invalid)
                    intKeyASCII = 0
        End Select
    End If
    
End Sub


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

' ***************************************************************************
' Routine:       GetTitleFont
'
' Description:   Captues the font information
'
' Parameters:    frmForm - Name of the form whose caption is to be centered
'
' Returns:       Complete type structure describing the font used on this form
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-OCT-2000  Tom Pydeski
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=12103&lngWId=1
' 16-APR-2001  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function GetTitleFont(ByRef frmForm As Form) As StdFont
    
    ' Called by CenterCaption()
    
    Dim typNCM        As NONCLIENTMETRICS
    Dim typLogFont    As LogFont
    Dim typTargetFont As Font
    
    On Error GoTo GetTitleFont_Error

    ZeroMemory typNCM, Len(typNCM)
    ZeroMemory typLogFont, Len(typLogFont)
    
    typNCM.cbSize = Len(typNCM)    ' initialize variables
    
    ' Make the API to get the windows position
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, 0&, typNCM, 0&
    
    ' See if there are any fonts
    If typNCM.iCaptionHeight = 0 Then
        ' If no fonts involved then set to zero
        typLogFont.FontHeight = 0
    Else
        ' save the height of the caption font
        typLogFont = typNCM.LFCaptionFont
    End If
    
    Set typTargetFont = New StdFont
    
    With typTargetFont
        .Charset = typLogFont.FontCharSet
        .Weight = typLogFont.FontWeight
        .Name = typLogFont.FontFaceName
        .Strikethrough = typLogFont.FontStrikeOut
        .Underline = typLogFont.FontUnderline
        .Italic = typLogFont.FontItalic
        .Bold = (typLogFont.FontWeight = 700)
        .Size = -(typLogFont.FontHeight * (72 / GetDeviceCaps(frmForm.hDC, LOGPIXELSY)))
    End With
    
    Set GetTitleFont = typTargetFont

GetTitleFont_CleanUp:
    On Error GoTo 0
    Exit Function

GetTitleFont_Error:
    ErrorMsg MODULE_NAME, "GetTitleFont", Err.Description
    Resume GetTitleFont_CleanUp

End Function

' ***************************************************************************
' Routine:       Edit_Copy
'
' Description:   Copy highlighted text to the clipboard. See Keydown event
'                for the text boxes to see an example of the code calling
'                this routine.
'
' Special Logic: When the user highlights text with the cursor and presses
'                CTRL+C to perform a copy function.  The highlighted text
'                is then loaded into the clipboard.
'
' Parameters:    ctlTextBox - TextBox control
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-APR-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub Edit_Copy(ByRef ctlTextBox As TextBox)

    ' Called by TextBoxKeyDown()
    
    Clipboard.Clear                       ' Clear clipboard contents
    Clipboard.SetText ctlTextBox.SelText  ' Load clipboard with highlighted text
  
End Sub

' ***************************************************************************
' Routine:       Edit_Cut
'
' Description:   Copy highlighted text to the clipboard and then remove it
'                from the text box. See Keydown event for the text boxes to
'                see an example of the code calling this routine.
'
' Special Logic: When the user highlights text with the cursor and presses
'                CTRL+X to perform a cutting function.  The highlighted text
'                is then moved to the clipboard.
'
' Parameters:    ctlTextBox - TextBox control
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-APR-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub Edit_Cut(ByRef ctlTextBox As TextBox)

    ' Called by TextBoxKeyDown()
    
    Clipboard.Clear                        ' Clear clipboard contents
    Clipboard.SetText ctlTextBox.SelText   ' Load clipboard with highlighted text
    ctlTextBox.SelText = vbNullString                ' Empty TextBox contents
  
End Sub

' ***************************************************************************
' Routine:       Edit_Delete
'
' Description:   Copy highlighted text to the clipboard and then remove it
'                from the text box. See Keydown event for the text boxes to
'                see an example of the code calling this routine.
'
' Special Logic: When the user highlights text with the cursor and presses
'                CTRL+X to perform a cutting function.  The highlighted text
'                is then moved to the clipboard and the clipboard is emptied
'
' Parameters:    ctlTextBox - TextBox control
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-APR-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub Edit_Delete(ByRef ctlTextBox As TextBox)

    ' Called by TextBoxKeyDown()
    
    ctlTextBox.SelText = vbNullString  ' remove highlighted text from TextBox
  
End Sub

' ***************************************************************************
' Routine:       Edit_Paste
'
' Description:   Copy whatever text is being held in the clipboard and then
'                paste it in the text box. See Keydown event for the text
'                boxes to see an example of the code calling this routine.
'
' Parameters:    ctlTextBox - TextBox control
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-APR-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub Edit_Paste(ByRef ctlTextBox As TextBox)

    ' Called by TextBoxKeyDown()
    
    ctlTextBox.SelText = Clipboard.GetText()  ' unload clipboard into TextBox
  
End Sub


