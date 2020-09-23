VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Build Tables"
   ClientHeight    =   4380
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   5895
   HelpContextID   =   10
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5895
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   3150
      Top             =   3735
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "&Stop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   1
      Left            =   4545
      TabIndex        =   11
      Top             =   3645
      Width           =   555
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2805
      Left            =   90
      ScaleHeight     =   2775
      ScaleWidth      =   5655
      TabIndex        =   13
      Top             =   720
      Width           =   5685
      Begin VB.CheckBox chkFile 
         Caption         =   " Whirlpool Tables"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   225
         TabIndex        =   6
         Top             =   1395
         Width           =   1680
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1980
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "txtQty"
         Top             =   1395
         Width           =   510
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View Files"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4350
         TabIndex        =   8
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1980
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "txtQty"
         Top             =   945
         Width           =   510
      End
      Begin VB.ComboBox cboLayout 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   2610
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   945
         Width           =   1995
      End
      Begin VB.TextBox txtOptionalTitle 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   90
         TabIndex        =   9
         Text            =   "txtOptionalTitle"
         Top             =   2295
         Width           =   5490
      End
      Begin VB.ComboBox cboLayout 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   2610
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   450
         Width           =   1995
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1980
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "txtQty"
         Top             =   450
         Width           =   510
      End
      Begin VB.CheckBox chkFile 
         Caption         =   " Skipjack F-Tables"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   225
         TabIndex        =   3
         Top             =   945
         Width           =   1680
      End
      Begin VB.CheckBox chkFile 
         Caption         =   " GOST S-Boxes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   225
         TabIndex        =   0
         Top             =   450
         Width           =   1680
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4800
         TabIndex        =   24
         Top             =   1395
         Width           =   540
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4800
         TabIndex        =   23
         Top             =   945
         Width           =   540
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4800
         TabIndex        =   22
         Top             =   450
         Width           =   540
      End
      Begin VB.Label lblFiles 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Count"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   4875
         TabIndex        =   21
         Top             =   225
         Width           =   435
      End
      Begin VB.Label lblFiles 
         BackStyle       =   0  'Transparent
         Caption         =   "Optional title line in output file  ( Max 60 characters )"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   20
         Top             =   2025
         Width           =   3840
      End
      Begin VB.Label lblFiles 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Table layout"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   2700
         TabIndex        =   19
         Top             =   225
         Width           =   1560
      End
      Begin VB.Label lblFiles 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2025
         TabIndex        =   18
         Top             =   225
         Width           =   465
      End
      Begin VB.Label lblFiles 
         BackStyle       =   0  'Transparent
         Caption         =   "Files to be created"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   14
         Top             =   135
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   2
      Left            =   5175
      TabIndex        =   12
      Top             =   3645
      Width           =   555
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "&Go"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   0
      Left            =   4545
      TabIndex        =   10
      Top             =   3645
      Width           =   555
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kenneth Ives"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2393
      TabIndex        =   17
      Top             =   405
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Build Table Files"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1785
      TabIndex        =   16
      Top             =   90
      Width           =   2310
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "lblDisclaimer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   135
      TabIndex        =   15
      Top             =   3735
      Width           =   3855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Routine:       frmMain
'
' Description:   Build GOST S-Box table sets and Skipjack F-Table sets.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-Sep-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MIN_QTY As Long = 1
  Private Const MAX_QTY As Long = 99

' ***************************************************************************
' Module Variables
'
' Variable name:     mblnGost
' Naming standard:   m bln Gost
'                    - --- ---------
'                    |  |   |_______ Variable subname
'                    |  |___________ Data type (Boolean)
'                    |______________ Module level designator
'
' ***************************************************************************
  Private mstrFile      As String
  Private mblnGost      As Boolean
  Private mblnSkipjack  As Boolean
  Private mblnWhirlpool As Boolean
  
Private Sub LoadComboboxes()
    
    Dim lngIndex As Long
    
    With frmMain
        For lngIndex = 0 To (.cboLayout.Count - 1)
        
            .txtQty(lngIndex).Text = IIf(lngIndex = 0, gstrGostQty, gstrSkipQty)
        
            With .cboLayout(lngIndex)
                .Clear           ' Verify combobox is empty
                .AddItem "Horizontal 1 line"
                .AddItem "Horizontal 2 lines"
                .AddItem "Horizontal 4 lines"
                .AddItem "Horizontal 5 lines"
                .AddItem "Horizontal 10 lines"
                .AddItem "Vertical 2 lines"
                .AddItem "Vertical 4 lines"
                .AddItem "Vertical 5 lines"
                .AddItem "Vertical 10 lines"
                .AddItem "Vertical 20 lines"
                .ListIndex = IIf(lngIndex = 0, Val(gstrGostCase), Val(gstrSkipCase))
            End With
        
        Next lngIndex
    
        .txtQty(2).Text = gstrWhirlQty
    
    End With

End Sub

Private Sub SetControls()

    DoEvents
    With frmMain
        .picFrame.Enabled = False
        .cmdChoice(0).Visible = False  ' Go button
        .cmdChoice(0).Enabled = False
        .cmdChoice(1).Enabled = True   ' Stop button
        .cmdChoice(1).Visible = True
        .cmdChoice(2).Enabled = False  ' Exit button
    End With
                
End Sub

Private Sub ResetControls()

    DoEvents
    With frmMain
        .picFrame.Enabled = True
        .cmdChoice(0).Enabled = True   ' Go button
        .cmdChoice(0).Visible = True
        .cmdChoice(1).Visible = False  ' Stop button
        .cmdChoice(1).Enabled = False
        .cmdChoice(2).Enabled = True   ' Exit button
    End With
    
End Sub

Private Function IsGoodData(ByVal lngIndex As Long) As Boolean

    ' Test here instead of Lost_Focus event
    ' so user is not forced to enter data
    ' if attempting to exit application
    
    IsGoodData = False   ' Preset to FALSE
    
    If Not IsNumeric(txtQty(lngIndex).Text) Then
        
        InfoMsg "Quantity must be numeric." & vbNewLine & _
                "Min=" & CStr(MIN_QTY) & "    Max=" & CStr(MAX_QTY)
        txtQty(lngIndex).SetFocus   ' Highlight quantity textbox
    
    ElseIf Val(txtQty(lngIndex).Text) < MIN_QTY Then
        
        InfoMsg "Quantity must be greater than zero." & vbNewLine & _
                "Min=" & CStr(MIN_QTY) & "    Max=" & CStr(MAX_QTY)
        txtQty(lngIndex).SetFocus   ' Highlight quantity textbox
    
    ElseIf Val(txtQty(lngIndex).Text) > MAX_QTY Then
        
        InfoMsg "Quantity exceeds maximum range." & vbNewLine & _
                "Min=" & CStr(MIN_QTY) & "    Max=" & CStr(MAX_QTY)
        txtQty(lngIndex).SetFocus   ' Highlight quantity textbox
    
    ElseIf Len(txtOptionalTitle.Text) > 60 Then
        
        InfoMsg "Optional title line cannot exceed sixty   " & vbNewLine & _
                "characters to include blank spaces."
        txtOptionalTitle.SetFocus   ' Highlight optional data textbox
    
    Else
        
        IsGoodData = True                             ' Set flag for good data
        gstrOptTitle = Trim$(txtOptionalTitle.Text)   ' Save optional title line
        
        Select Case lngIndex
               Case 0   ' Save Gost data
                    gstrGostQty = Trim$(txtQty(0).Text)           ' Save quantity
                    gstrGostCase = CStr(cboLayout(0).ListIndex)   ' Save case statement format
                        
               Case 1   ' Save Skipjack data
                    gstrSkipQty = Trim$(txtQty(1).Text)           ' Save quantity
                    gstrSkipCase = CStr(cboLayout(1).ListIndex)   ' Save case statement format
                        
               Case 2   ' Save Whirlpool data
                    gstrWhirlQty = Trim$(txtQty(2).Text)           ' Save quantity
        End Select
    End If
    
End Function

Private Sub chkFile_Click(Index As Integer)
    
    With frmMain
        .lblCount(0).Caption = vbNullString   ' Empty accumulator boxes
        .lblCount(1).Caption = vbNullString
        .lblCount(2).Caption = vbNullString
        
        Select Case Index
        
               Case 0  ' Gost checkbox
                    mblnGost = Not mblnGost                         ' Toggle boolean flag
                    .txtQty(0).Enabled = mblnGost                   ' Enable/Disable textbox
                    .cboLayout(0).Enabled = mblnGost                ' Enable/Disable combobox
                    gstrGostChk = CStr(Abs(CInt(mblnGost)))         ' Save checkbox value
                    
               Case 1  ' Skipjack checkbox
                    mblnSkipjack = Not mblnSkipjack                 ' Toggle boolean flag
                    .txtQty(1).Enabled = mblnSkipjack               ' Enable/Disable textbox
                    .cboLayout(1).Enabled = mblnSkipjack            ' Enable/Disable combobox
                    gstrSkipChk = CStr(Abs(CInt(mblnSkipjack)))     ' Save checkbox value
                    
               Case 2  ' Whirlpool checkbox
                    mblnWhirlpool = Not mblnWhirlpool               ' Toggle boolean flag
                    .txtQty(2).Enabled = mblnWhirlpool              ' Enable/Disable textbox
                    gstrWhirlChk = CStr(Abs(CInt(mblnWhirlpool)))   ' Save checkbox value
        End Select
    End With
        
End Sub

Private Sub cmdChoice_Click(Index As Integer)

    Dim strTarget As String
    
    Const TARGET_FOLDER = "Tables"
    
    Select Case Index
            
           Case 0  ' GO button
           
                gblnStopProcessing = False
                DoEvents
                
                lblCount(0).Caption = vbNullString   ' Empty accumulator boxes
                lblCount(1).Caption = vbNullString
                lblCount(2).Caption = vbNullString
                
                ' At least one checkbox must be active
                If Not mblnGost And _
                   Not mblnSkipjack And _
                   Not mblnWhirlpool Then
                    
                    InfoMsg "Cannot identify which file to create."
                    Exit Sub
                End If
                                
                SetControls
                strTarget = QualifyPath(App.Path) & TARGET_FOLDER
                            
                ' Create GOST table sets
                If mblnGost Then
                    
                    ' Evaluate textbox data
                    If IsGoodData(0) Then
                                        
                        ' Create file with table sets
                        GostTables frmMain, _
                                   CLng(txtQty(0).Text), _
                                   cboLayout(0).ListIndex, _
                                   txtOptionalTitle.Text
                    
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            ResetControls
                            Exit Sub
                        End If
                        
                    End If
                End If
                
                ' Create Skipjack table sets
                If mblnSkipjack Then
                                                            
                    ' Evaluate textbox data
                    If IsGoodData(1) Then
                    
                        ' Create file with table sets
                        SkipjackTables frmMain, _
                                       CLng(txtQty(1).Text), _
                                       cboLayout(1).ListIndex, _
                                       txtOptionalTitle.Text
                    
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            ResetControls
                            Exit Sub
                        End If
                        
                    End If
                End If
                
                ' Create Whirlpool table sets
                If mblnWhirlpool Then
                                                            
                    ' Evaluate textbox data
                    If IsGoodData(2) Then
                    
                        ' Create file with table sets
                        WhirlpoolTables frmMain, _
                                        CLng(txtQty(2).Text), _
                                        txtOptionalTitle.Text
                    
                        ' An error occurred or user opted to STOP processing
                        DoEvents
                        If gblnStopProcessing Then
                            ResetControls
                            Exit Sub
                        End If
                        
                    End If
                End If
                
                ' Finish messages
                If mblnGost And (Not mblnSkipjack) And (Not mblnWhirlpool) Then
                        
                    InfoMsg "Finished creating " & txtQty(0).Text & _
                            " tables to be used in Gost cipher in folder" & Space$(3) & _
                            vbNewLine & vbNewLine & strTarget
                                            
                ElseIf mblnSkipjack And (Not mblnGost) And (Not mblnWhirlpool) Then
                        
                    InfoMsg "Finished creating " & txtQty(1).Text & _
                            " tables to be used in Skipjack cipher in folder" & Space$(3) & _
                            vbNewLine & vbNewLine & strTarget
                
                ElseIf mblnWhirlpool And (Not mblnGost) And (Not mblnSkipjack) Then
                        
                    InfoMsg "Finished creating " & txtQty(2).Text & _
                            " tables to be used in Whirlpool hash in folder" & Space$(3) & _
                            vbNewLine & vbNewLine & strTarget
                
                ElseIf mblnGost And mblnSkipjack And (Not mblnWhirlpool) Then
                        
                    InfoMsg "Finished creating tables for" & vbNewLine & vbNewLine & _
                            "Gost cipher     -" & Format$(txtQty(0).Text, "@@@") & vbNewLine & _
                            "Skipjack cipher -" & Format$(txtQty(1).Text, "@@@") & _
                            vbNewLine & vbNewLine & strTarget

                ElseIf mblnGost And mblnWhirlpool And (Not mblnSkipjack) Then
                        
                    InfoMsg "Finished creating tables for" & vbNewLine & vbNewLine & _
                            "Gost cipher     -" & Format$(txtQty(0).Text, "@@@") & vbNewLine & _
                            "Whirlpool hash  -" & Format$(txtQty(2).Text, "@@@") & _
                            vbNewLine & vbNewLine & strTarget

                ElseIf mblnSkipjack And mblnWhirlpool And (Not mblnGost) Then
                        
                    InfoMsg "Finished creating tables for" & vbNewLine & vbNewLine & _
                            "Skipjack cipher -" & Format$(txtQty(1).Text, "@@@") & vbNewLine & _
                            "Whirlpool hash  -" & Format$(txtQty(2).Text, "@@@") & _
                            vbNewLine & vbNewLine & strTarget

                ElseIf mblnGost And mblnSkipjack And mblnWhirlpool Then
                        
                    InfoMsg "Finished creating tables for" & vbNewLine & vbNewLine & _
                            "Gost cipher     -" & Format$(txtQty(0).Text, "@@@") & vbNewLine & _
                            "Skipjack cipher -" & Format$(txtQty(1).Text, "@@@") & vbNewLine & _
                            "Whirlpool hash  -" & Format$(txtQty(2).Text, "@@@") & _
                            vbNewLine & vbNewLine & strTarget
                End If
                
                ResetControls  ' Reset buttons when finished
           
           Case 1  ' Stop button
                gblnStopProcessing = True  ' Reset boolean flag
                ResetControls              ' Reset buttons when finished
                DoEvents                   ' Allow time for notification process
                
           Case Else
                DoEvents
                gblnStopProcessing = True  ' Reset boolean flag
                
                ' Save GOST quantity
                If Val(Trim$(txtQty(0).Text)) >= 0 Then
                    gstrGostQty = Trim$(txtQty(0).Text)
                Else
                    gstrGostQty = "0"
                End If
                
                gstrGostCase = CStr(cboLayout(0).ListIndex)   ' Save case statement format
                        
                ' Save Skipjack quantity
                If Val(Trim$(txtQty(1).Text)) >= 0 Then
                    gstrSkipQty = Trim$(txtQty(1).Text)
                Else
                    gstrSkipQty = "0"
                End If
                
                gstrSkipCase = CStr(cboLayout(1).ListIndex)   ' Save case statement format
                        
                ' Save Whilpool quantity
                If Val(Trim$(txtQty(2).Text)) >= 0 Then
                    gstrWhirlQty = Trim$(txtQty(2).Text)
                Else
                    gstrWhirlQty = "0"
                End If
                
                TerminateProgram   ' End this application
    End Select
    
End Sub

' ***************************************************************************
' Routine:       cmdView_Click
'
' Description:   Opens File Open dialog box so user can browse for a
'                former report file.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Jan-2011  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Sub cmdView_Click()

    On Error GoTo ErrHandler
    
    cmDialog.CancelError = True  ' Set Cancel to True.
    mstrFile = vbNullString
  
    ' Setup and display the "FILE OPEN" dialog box
    With cmDialog
         .Flags = cdlOFNHideReadOnly Or _
                  cdlOFNExplorer Or _
                  cdlOFNLongNames Or _
                  cdlOFNFileMustExist
         
         .InitDir = QualifyPath(App.Path) & "Tables"
         .FileName = vbNullString
         
         ' Set filters
         .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
         .FilterIndex = 1   ' Specify default filter
         
         .ShowOpen          ' Display the Open dialog box
    End With
    
    ' Capture name of selected file
    mstrFile = TrimStr(cmDialog.FileName)
    
    If Len(mstrFile) > 0 Then
        DisplayFile mstrFile, frmMain   ' Review this file
    End If
    
    Exit Sub
    
ErrHandler:
    ' User pressed the Cancel button
    mstrFile = vbNullString
    Exit Sub

End Sub

Private Sub Form_Load()

    LoadComboboxes                              ' Fill comboboxes
    mblnGost = Not CBool(Val(gstrGostChk))      ' Set opposite flags
    mblnSkipjack = Not CBool(Val(gstrSkipChk))
    mblnWhirlpool = Not CBool(Val(gstrWhirlChk))
    
    With frmMain
        .Caption = PGM_NAME & gstrVersion
        .lblDisclaimer.Caption = "This is a freeware product." & vbNewLine & _
                                 "No warranties or guarantees implied or intended."
        .lblCount(0).Caption = vbNullString
        .lblCount(1).Caption = vbNullString
        .lblCount(2).Caption = vbNullString
        
        ' Update flag settings
        If gstrGostChk = "0" Then
            .chkFile(0).Value = vbUnchecked
            chkFile_Click 0
        Else
            .chkFile(0).Value = vbChecked
        End If
        
        If gstrSkipChk = "0" Then
            .chkFile(1).Value = vbUnchecked
            chkFile_Click 1
        Else
            .chkFile(1).Value = vbChecked
        End If
        
        If gstrWhirlChk = "0" Then
            .chkFile(2).Value = vbUnchecked
            chkFile_Click 2
        Else
            .chkFile(2).Value = vbChecked
        End If
        
        .txtOptionalTitle.Text = gstrOptTitle   ' Optional title line
        
        .cmdChoice(0).Enabled = True    ' GO button
        .cmdChoice(0).Visible = True
        .cmdChoice(1).Visible = False   ' STOP button
        .cmdChoice(1).Enabled = False
        .cmdChoice(2).Visible = True    ' EXIT button
        
        ' Center form on screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Show
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        TerminateProgram   ' "X" selected in upper right corner
    End If
    
End Sub

Private Sub lblAuthor_Click()
    SendEmail   ' Send email to author of this application
End Sub

Private Sub txtOptionalTitle_GotFocus()

    ' Highlight all data in textbox
    With txtOptionalTitle
         .SelStart = 0             ' start with first char
         .SelLength = Len(.Text)   ' to end of data string
    End With
  
End Sub

Private Sub txtOptionalTitle_LostFocus()
    
    ' Remove leading and trailing blank spaces
    txtOptionalTitle.Text = Trim$(txtOptionalTitle.Text)
    
    If Len(txtOptionalTitle.Text) = 0 Then
        txtOptionalTitle.Text = vbNullString  ' Verify textbox is empty
    End If
    
End Sub

Private Sub txtQty_GotFocus(Index As Integer)

    lblCount(0).Caption = vbNullString
    lblCount(1).Caption = vbNullString
    lblCount(2).Caption = vbNullString
    
    ' Highlight all data in textbox
    With txtQty(Index)
         .SelStart = 0             ' start with first char
         .SelLength = Len(.Text)   ' to end of data string
    End With
  
End Sub

Private Sub txtQty_KeyPress(Index As Integer, KeyAscii As Integer)

    ' Evaluate data as it is entered into textbox
    Select Case KeyAscii
           Case 9             ' Tab key
                KeyAscii = 0
                SendKeys "{TAB}"
           Case 13            ' Enter key (no bell sound)
                KeyAscii = 0
           Case 8, 48 To 57   ' backspace & numeric keys only
                ' good data
           Case Else          ' everything else
                KeyAscii = 0
    End Select
                              
End Sub

