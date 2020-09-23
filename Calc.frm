VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCalc 
   Caption         =   "Calculator-2"
   ClientHeight    =   3705
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3705
   HelpContextID   =   10
   Icon            =   "Calc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMemValue 
      Caption         =   "Memory Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1680
      TabIndex        =   32
      Top             =   3000
      WhatsThisHelpID =   10
      Width           =   1935
      Begin VB.Label lblMemoryVal 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1700
      End
   End
   Begin VB.CommandButton Number 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   0
      Left            =   690
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2580
      Width           =   525
   End
   Begin VB.CommandButton cmdOff 
      Caption         =   "Off"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Turn off (Exit Program)."
      Top             =   3075
      WhatsThisHelpID =   10
      Width           =   525
   End
   Begin VB.TextBox txtMemory 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   5400
      TabIndex        =   28
      Text            =   "0"
      Top             =   75
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdNull 
      Caption         =   "Null"
      Height          =   425
      Left            =   5400
      TabIndex        =   0
      Top             =   2520
      Width           =   525
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Number 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   7
      Left            =   690
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1110
      Width           =   525
   End
   Begin VB.CommandButton Number 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   8
      Left            =   1290
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1110
      Width           =   525
   End
   Begin VB.CommandButton Number 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   9
      Left            =   1890
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1110
      Width           =   525
   End
   Begin VB.CommandButton Operator 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   3
      Left            =   2490
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Divide."
      Top             =   1110
      WhatsThisHelpID =   10
      Width           =   525
   End
   Begin VB.CommandButton cmdSquareRoot 
      Caption         =   "sqrt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Left            =   3090
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Calculate Square Root."
      Top             =   1110
      WhatsThisHelpID =   10
      Width           =   525
   End
   Begin VB.CommandButton Number 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   4
      Left            =   690
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1590
      Width           =   525
   End
   Begin VB.CommandButton Number 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   5
      Left            =   1290
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1590
      Width           =   525
   End
   Begin VB.CommandButton Number 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   6
      Left            =   1890
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1590
      Width           =   525
   End
   Begin VB.CommandButton Operator 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   2
      Left            =   2490
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Multiply."
      Top             =   1590
      WhatsThisHelpID =   10
      Width           =   525
   End
   Begin VB.CommandButton Percent 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Left            =   3090
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Convert to percentage value."
      Top             =   1590
      WhatsThisHelpID =   10
      Width           =   525
   End
   Begin VB.CommandButton Number 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   1
      Left            =   690
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2085
      Width           =   525
   End
   Begin VB.CommandButton Number 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   2
      Left            =   1290
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2085
      Width           =   525
   End
   Begin VB.CommandButton Number 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   3
      Left            =   1890
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2085
      Width           =   525
   End
   Begin VB.CommandButton Operator 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   1
      Left            =   2490
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subtract."
      Top             =   2085
      WhatsThisHelpID =   10
      Width           =   525
   End
   Begin VB.CommandButton cmdRecip 
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Left            =   3090
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Reciprocal of number."
      Top             =   2085
      WhatsThisHelpID =   10
      Width           =   525
   End
   Begin VB.CommandButton cmdPlusMinus 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2580
      Width           =   525
   End
   Begin VB.CommandButton Decimal 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2580
      Width           =   525
   End
   Begin VB.CommandButton Operator 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   0
      Left            =   2490
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Add."
      Top             =   2580
      WhatsThisHelpID =   10
      Width           =   525
   End
   Begin VB.CommandButton Operator 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   4
      Left            =   3090
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Equals."
      Top             =   2580
      WhatsThisHelpID =   10
      Width           =   525
   End
   Begin VB.CommandButton cmdMem 
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   0
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Memory Add."
      Top             =   2580
      WhatsThisHelpID =   10
      Width           =   525
   End
   Begin VB.CommandButton cmdMem 
      Caption         =   "MS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   1
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Memory Store."
      Top             =   2085
      WhatsThisHelpID =   10
      Width           =   525
   End
   Begin VB.CommandButton cmdMem 
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   2
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Memory Recall."
      Top             =   1590
      WhatsThisHelpID =   10
      Width           =   525
   End
   Begin VB.CommandButton cmdMem 
      Caption         =   "MC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   3
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Memory Clear."
      Top             =   1110
      WhatsThisHelpID =   10
      Width           =   525
   End
   Begin VB.TextBox txtMem 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   425
      Left            =   135
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   575
      Width           =   425
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0C0C0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   2755
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Clear."
      Top             =   575
      WhatsThisHelpID =   10
      Width           =   840
   End
   Begin VB.CommandButton cmdClearEntry 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   1840
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Clear Entry."
      Top             =   575
      WhatsThisHelpID =   10
      Width           =   840
   End
   Begin VB.CommandButton cmdBackSpace 
      Caption         =   "Backspace"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   690
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Delete last entry."
      Top             =   575
      WhatsThisHelpID =   10
      Width           =   1075
   End
   Begin VB.Label lblScreen 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   90
      TabIndex        =   29
      Top             =   75
      Width           =   3510
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      HelpContextID   =   210
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      HelpContextID   =   220
      Begin VB.Menu mnuViewStandard 
         Caption         =   "S&tandard"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewScientific 
         Caption         =   "&Scientific"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      HelpContextID   =   230
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Help Topics"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Calculator"
      End
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************
' Calculator-2 Application                              *
' Author: Robert Gross, Ian Williams, Microsoft, and    *
'         Nagalla Anil Choudary and others.             *
' Last modified: October 16, 2001                       *
'********************************************************

Dim i As Integer            ' Variable for different commandbuttons.
Dim Op1, Op2                ' Previously input operand.
Dim DecimalFlag As Integer  ' Decimal point present yet?
Dim NumOps As Integer       ' Number of operands.
Dim LastInput               ' Indicate type of last keypress event.
Dim OpFlag                  ' Indicate pending operation.
Dim TempReadout
Public memo As Double       ' For storing numbers in Memory
Dim prev As Double          ' For storing numbers in Memory

Dim LastInput_1 As Integer          ' Variable for code from Ian Williams Memory functions
Dim ShortenScreenRunning As Boolean ' Variable for code from Ian Williams Memory functions
Dim Mem As Double                   ' Varaible for code from Ian Williams Memory functions
Dim NewScreen As Boolean            ' Variable for code from Ian Williams Memory functions
Dim Length As Integer               ' Variable for code from Ian Williams Memory functions
Dim DecimalPoint As Integer         ' Variable for code from Ian Williams
Dim NumCalcs As Integer             ' Variable for code from Ian Williams
Dim Calc1 As Double                 ' Variable for code from Ian Williams
Dim Calc2 As Double                 ' Variable for code from Ian Williams
Dim Calculations                    ' Variable for code from Ian Williams
Dim TempDisplay As Double           ' Variable for code from Ian Williams

Private Sub cmdBackSpace_Click()

'Code from Ian Williams via Devx discussion group
'sent via email to me on October 14, 2001
Dim strCap As String
With lblScreen
    strCap = .Caption
If Len(strCap) And strCap <> "0." Then
    If Right$(strCap, 1) = "." Then
        strCap = Left$(strCap, Len(strCap) - 2) & "."
    Else
        strCap = Left$(strCap, Len(strCap) - 1)
    End If
    Else
        strCap = "0."
    End If
    If strCap = "." Then strCap = "0."
    .Caption = strCap
End With
'Send focus to the cmdNull command button so that the command
'button just pressed doesn't have the focus-trying to make the
'command buttons behave like the Microsoft Calculator.
    cmdNull.SetFocus

End Sub

Private Sub cmdNull_Click()
'Call the code to do a Operation
Call Operator_Click(4)
End Sub

Private Sub cmdOff_Click()
'Shut Calculator off
Unload Me
End Sub

Private Sub cmdPlusMinus_Click()
'Switch sign of the value
On Error GoTo ErrHandler
lblScreen.Caption = -Val(lblScreen.Caption)

'Send focus to the cmdNull command button so that the command
'button just pressed doesn't have the focus-trying to make the
'command buttons behave like the Microsoft Calculator.
    cmdNull.SetFocus
ErrHandler:
Exit Sub
End Sub

Private Sub cmdRecip_Click()
On Error GoTo ErrHandler

'Code from Nagalla Anil Chodary naccalc application
    Dim temp As Long
        With lblScreen
            temp = Val(.Caption)
                If temp <> 0 Then
                    .Caption = Str(1 / temp)
                    .Caption = Format(.Caption, "#0.0#########")
                Else
                    .Caption = "Error: Positive Infinity."
                End If
        End With
'Send focus to the cmdNull command button so that the command
'button just pressed doesn't have the focus-trying to make the
'command buttons behave like the Microsoft Calculator.
    cmdNull.SetFocus
ErrHandler:
Exit Sub
End Sub

Private Sub cmdSquareRoot_Click()
On Error GoTo ErrHandler

'Code from Nagalla Anil Chodary naccalc application
    With lblScreen
        If .Caption < 0 Then
            Exit Sub
        End If
        .Caption = Str(Sqr(Val(.Caption)))
        .Caption = Format(.Caption, "#0.0#########")
    End With
'Send focus to the cmdNull command button so that the command
'button just pressed doesn't have the focus-trying to make the
'command buttons behave like the Microsoft Calculator.
    cmdNull.SetFocus
ErrHandler:
Exit Sub
End Sub

Private Sub Decimal_Click()
    If LastInput_1 <> 2 Then
        lblScreen.Caption = Format(0, "0.")
    End If
    DecimalPoint = True
    LastInput_1 = 2

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = Asc("0") Then
        i = 0
         Number(i).SetFocus
         Number_Click (i)
         
    ElseIf KeyAscii = Asc("1") Then
        i = 1
          Number(i).SetFocus
          Number_Click (i)
          
    ElseIf KeyAscii = Asc("2") Then
        i = 2
          Number(i).SetFocus
          Number_Click (i)
        
    ElseIf KeyAscii = Asc("3") Then
        i = 3
          Number(i).SetFocus
          Number_Click (i)
          
    ElseIf KeyAscii = Asc("4") Then
        i = 4
          Number(i).SetFocus
          Number_Click (i)
          
    ElseIf KeyAscii = Asc("5") Then
        i = 5
          Number(i).SetFocus
          Number_Click (i)
          
    ElseIf KeyAscii = Asc("6") Then
        i = 6
          Number(i).SetFocus
          Number_Click (i)
          
    ElseIf KeyAscii = Asc("7") Then
        i = 7
          Number(i).SetFocus
          Number_Click (i)
          
    ElseIf KeyAscii = Asc("8") Then
        i = 8
          Number(i).SetFocus
          Number_Click (i)
          
    ElseIf KeyAscii = Asc("9") Then
        i = 9
          Number(i).SetFocus
          Number_Click (i)

    ElseIf KeyAscii = Asc("+") Then
        i = 0
            Operator(i).SetFocus
            Operator_Click (i)
          
    ElseIf KeyAscii = Asc("-") Then
        i = 1
            Operator(i).SetFocus
            Operator_Click (i)
          
    ElseIf KeyAscii = Asc("*") Then
        i = 2
            Operator(i).SetFocus
            Operator_Click (i)
          
    ElseIf KeyAscii = Asc("/") Then
        i = 3
            Operator(i).SetFocus
            Operator_Click (i)
    
    ElseIf KeyAscii = Asc("=") Then
        i = 4
            Operator(i).SetFocus
            Operator_Click (i)

    End If

If KeyAscii = vbKeyBack Then Call cmdBackSpace_Click

End Sub

Private Sub Form_Load()

'Code from Ian Williams Calculator Application
    NewScreen = True
    DecimalPoint = False
    NumCalcs = 0
    LastInput_1 = 0
    lblScreen.Caption = Format(0, "0.")
    txtMemory.Text = Format(0, "0.")
    ShortenScreenRunning = False
'End of Code from Ian Williams Calculator Application

    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
    mnuEditPaste.Enabled = False        'Disable Paste because there
                                        'isn't anything to Paste yet!
    mnuViewScientific.Enabled = False   'Disable Scientific until I
                                        'code this.

'Set color of Caption on the Operation buttons (e.g.,+,-,*)
For i = i To 4
SetButtonForeColor Operator(i), vbRed
Next i

SetButtonForeColor cmdBackSpace, vbRed
SetButtonForeColor cmdClearEntry, vbRed
SetButtonForeColor cmdClear, vbRed
SetButtonForeColor cmdMem(0), vbRed
SetButtonForeColor cmdMem(1), vbRed
SetButtonForeColor cmdMem(2), vbRed
SetButtonForeColor cmdMem(3), vbRed

i = 0

For i = i To 9
SetButtonForeColor Number(i), vbBlue
Next i

SetButtonForeColor Decimal, vbBlue
SetButtonForeColor cmdPlusMinus, vbBlue
SetButtonForeColor cmdSquareRoot, vbBlue
SetButtonForeColor Percent, vbBlue
SetButtonForeColor cmdRecip, vbBlue
SetButtonForeColor cmdOff, vbYellow

End Sub

Private Sub Form_Unload(Cancel As Integer)

For i = i To 4
UnsetButtonForeColor Operator(i)
Next i

UnsetButtonForeColor cmdBackSpace
UnsetButtonForeColor cmdClearEntry
UnsetButtonForeColor cmdClear
UnsetButtonForeColor cmdMem(0)
UnsetButtonForeColor cmdMem(1)
UnsetButtonForeColor cmdMem(2)
UnsetButtonForeColor cmdMem(3)

i = 0
For i = i To 9
UnsetButtonForeColor Number(i)
Next i

UnsetButtonForeColor Decimal
UnsetButtonForeColor cmdPlusMinus
UnsetButtonForeColor cmdSquareRoot
UnsetButtonForeColor Percent
UnsetButtonForeColor cmdRecip
UnsetButtonForeColor cmdOff

End Sub

Private Sub mnuEditPaste_Click()
    lblScreen.Caption = ""
    lblScreen.Caption = Clipboard.GetText
    LastInput = "NUMS"
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show
End Sub
' Click event procedure for C (Clear) key.
' Reset the lblScreen and initializes variables.
Private Sub cmdClear_Click()
    lblScreen.Caption = Format(0, "0.")
    Op1 = 0
    Op2 = 0

'Send focus to the cmdNull command button so that the command
'button just pressed doesn't have the focus-trying to make the
'command buttons behave like the Microsoft Calculator.
    cmdNull.SetFocus

End Sub

' Click event procedure for CE (Clear Entry) key.
Private Sub cmdClearEntry_Click()
    lblScreen.Caption = Format(0, "0.")
    DecimalFlag = False
    LastInput = "CE"
'Send focus to the cmdNull command button so that the command
'button just pressed doesn't have the focus-trying to make the
'command buttons behave like the Microsoft Calculator.
    cmdNull.SetFocus

End Sub

Private Sub mnuHelpHelp_Click()
    ' Open the Help File
    With CommonDialog1
        .HelpFile = App.Path & "\" & "Calculator-2.hlp"
        .HelpCommand = cdlHelpContents
        .ShowHelp
    End With

ErrHandler:
'Error occurred
Exit Sub

End Sub

Private Sub Operator_Click(Index As Integer)
'This code copied from Ian William's Calculator Application
    TempDisplay = lblScreen.Caption 'number in display is move to a Variable (temporarily)
    If LastInput_1 = 2 Then
        NumCalcs = NumCalcs + 1
    End If
    Select Case NumCalcs
        Case 1
            Calc1 = lblScreen
        Case Else
            Calc2 = TempDisplay
                Select Case Calculations
                    Case "="
                        Calc1 = CDbl(Calc2)
                    Case "+"
                        Calc1 = Calc1 + Calc2
                    Case "-"
                        Calc1 = Calc1 - Calc2
                    Case "*"
                        Calc1 = Calc1 * Calc2
                    Case "/"
                    If Calc2 = 0 Then
                        lblScreen.Caption = "Error: Positive Infinity."
                        Exit Sub
                    Else
                        Calc1 = Calc1 / Calc2
                    End If
                End Select
                lblScreen = Calc1
                NumCalcs = 1
    End Select
    Calculations = Operator(Index).Caption
    NewScreen = True
'Send focus to the cmdNull command button so that the command
'button just pressed doesn't have the focus-trying to make the
'command buttons behave like the Microsoft Calculator.
    cmdNull.SetFocus

End Sub

' Click event procedure for percent key (%).
' Compute and lblScreen a percentage of the first operand.
Private Sub Percent_Click()
On Error GoTo ErrHandler

    lblScreen.Caption = lblScreen.Caption / 100
    LastInput = "Ops"
    OpFlag = "%"
    NumOps = NumOps + 1
    DecimalFlag = True
'Send focus to the cmdNull command button so that the command
'button just pressed doesn't have the focus-trying to make the
'command buttons behave like the Microsoft Calculator.
    cmdNull.SetFocus
ErrHandler:
Exit Sub
End Sub

Private Sub mnuEditCopy_Click()
        Clipboard.Clear
        Clipboard.SetText lblScreen.Caption
'Enable Paste now
mnuEditPaste.Enabled = True

End Sub
Private Sub cmdMem_Click(Index As Integer)

'Code for saving numbers into memory and recalling them
'Following code copied from Ian Williams Calculator Application
    Select Case Index
        Case 0 'M+ - Represents adding to the Memory value
            Mem = Mem + Val(lblScreen.Caption)
            txtMemory = lblScreen.Caption
            MemFunc
        Case 1 'MS - represents Memory Store
            Mem = Val(lblScreen.Caption)
            MemFunc
        Case 2 'MR
            lblScreen.Caption = Mem
            LastInput_1 = 2
            MemFunc
        Case 3 'MC
            txtMemory.Text = "0."
            Mem = 0
            lblScreen.Caption = txtMemory.Text
            MemFunc
    End Select
'Send focus to the cmdNull command button so that the command
'button just pressed doesn't have the focus-trying to make the
'command buttons behave like the Microsoft Calculator.
    cmdNull.SetFocus

End Sub
Private Sub MemFunc()
'This code copied from Ian Williams Calculator Application
    Length = 25
    NewScreen = True
    txtMemory.Text = Mem
    If txtMemory Like "*.*" Then Length = Length + 1
    If txtMemory Like "*-*" Then Length = Length + 1
    txtMemory = Mid(txtMemory, 1, Length)
    lblMemoryVal = txtMemory.Text
    lblMemoryVal = Format(lblMemoryVal, "#0.0##########")

    If Mem = 0 Then
        txtMem.Text = ""
        lblMemoryVal = ""
    Else
       txtMem.Text = "M"
    End If
End Sub
Sub ShortenScreen()
    If ShortenScreenRunning Then Exit Sub
    ShortenScreenRunning = True
    
    Length = 26
    If lblScreen Like "*-*" Then Length = Length + 1
    If lblScreen Like "*." Then Length = Length - 1
    
    lblScreen = Mid(lblScreen, 1, Length)
    If Not lblScreen Like "*.*" Then
        lblScreen = lblScreen + "."
    End If
    
    ShortenScreenRunning = False
End Sub

Private Sub Number_Click(Index As Integer)

'Code from Ian Williams Calculator Application
    If NewScreen = True Then    'If the last input is mpt equal to any number (if it is 0), then goto next line.
        lblScreen.Caption = ""    'Sets the screen to have 0. on it.
        NewScreen = False
    End If
    If lblScreen = "0." Then
        lblScreen = ""
    End If
    If lblScreen.Caption = "Error" Then       'If the display has "Error" on it, then
        lblScreen.Caption = Format(0, "0.")   'the display will be set to display 0.
    End If
    If DecimalPoint = True Then
        lblScreen.Caption = lblScreen.Caption & Number(Index).Caption   'if there is a decimal point, then the program adds the numbers to the end of the number after the decimal point.
    Else
        lblScreen.Caption = Left(lblScreen.Caption, InStr(lblScreen.Caption, ".") - 1) & Number(Index).Caption & "."  'Inserts a number between the number and the decimal point. (it adds the number then puts a decimal point at the end)
    End If
    LastInput_1 = 2 'LastInput is a number. (allows for calculations)

'Send focus to the cmdNull command button so that the command
'button just pressed doesn't have the focus-trying to make the
'command buttons behave like the Microsoft Calculator.
    cmdNull.SetFocus

End Sub
Private Sub lblScreen_Change()
    ShortenScreen 'Calls the ShortenScreen Sub
End Sub

