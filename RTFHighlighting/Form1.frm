VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RTF Tricks"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Also..."
      Height          =   855
      Left            =   5160
      TabIndex        =   36
      Top             =   1800
      Width           =   2895
      Begin VB.CheckBox ChWordWrap 
         Caption         =   "WordWrap"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   480
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox ChHook 
         Caption         =   "Remove Selection Rectangles"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Mouse"
      Height          =   1815
      Left            =   3000
      TabIndex        =   30
      Top             =   1800
      Width           =   2055
      Begin VB.CheckBox ChHideCaret 
         Caption         =   "Hide Caret"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox ChSelectCharacter 
         Caption         =   "Select Character"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox ChCaret 
         Caption         =   "Move Caret"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblCharacter 
         AutoSize        =   -1  'True
         Caption         =   "Charater: "
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   1530
         Width           =   690
      End
      Begin VB.Label lblCharacterNumber 
         AutoSize        =   -1  'True
         Caption         =   "Charater Pos: 0"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   1290
         Width           =   1095
      End
      Begin VB.Label lblLineNumber 
         AutoSize        =   -1  'True
         Caption         =   "Line: 0"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   1050
         Width           =   480
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Highlight Word"
      Height          =   1815
      Left            =   120
      TabIndex        =   23
      Top             =   1800
      Width           =   2775
      Begin VB.TextBox txtWord 
         Height          =   285
         Left            =   240
         TabIndex        =   29
         Text            =   "HighLight"
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton OptWord 
         Caption         =   "All ocurrences"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton OptWord 
         Caption         =   "First found"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdHighlightWord 
         Caption         =   "Highlight"
         Height          =   255
         Left            =   1560
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cboWord 
         Height          =   315
         ItemData        =   "Form1.frx":014A
         Left            =   1320
         List            =   "Form1.frx":0160
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Highlight Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   900
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Reset"
      Height          =   975
      Left            =   5160
      TabIndex        =   20
      Top             =   2640
      Width           =   2895
      Begin VB.CommandButton cmdResetColors 
         Caption         =   "Remove Syntax Colors"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdUnHighlight 
         Caption         =   "Remove All Highlighting"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Syntax Colors"
      Height          =   1575
      Left            =   5160
      TabIndex        =   11
      Top             =   120
      Width           =   2895
      Begin VB.ComboBox cboSyntax 
         Height          =   315
         Index           =   2
         ItemData        =   "Form1.frx":018A
         Left            =   1440
         List            =   "Form1.frx":01A0
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdSyntax 
         Caption         =   "VB KeyWords"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox cboSyntax 
         Height          =   315
         Index           =   1
         ItemData        =   "Form1.frx":01CA
         Left            =   1440
         List            =   "Form1.frx":01E0
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdSyntax 
         Caption         =   "Strings"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cboSyntax 
         Height          =   315
         Index           =   0
         ItemData        =   "Form1.frx":020A
         Left            =   1440
         List            =   "Form1.frx":0220
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdSyntax 
         Caption         =   "Comments"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtComments 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   12
         Text            =   "'"
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "CHR"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   1080
         TabIndex        =   19
         Top             =   220
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clever Selection"
      Height          =   1575
      Left            =   3000
      TabIndex        =   7
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select Below"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select Above"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select All"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Highlighting Selected Text"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton cmdHighlight 
         Caption         =   "Highlight"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton OptManualHighlight 
         Caption         =   "Manual"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptAutoHighlight 
         Caption         =   "AutoHighlight on Mouse_Up"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cboHLColor 
         Height          =   315
         ItemData        =   "Form1.frx":024A
         Left            =   1320
         List            =   "Form1.frx":0260
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Highlight Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   420
         Width           =   1095
      End
   End
   Begin RichTextLib.RichTextBox RTF 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      FileName        =   "C:\Documents and Settings\Administrator\Desktop\ffff.txt"
      TextRTF         =   $"Form1.frx":028A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - November 2002
'If you got it elsewhere - they stole it from PSC.

'Please visit our website at www.psst.com.au

Option Explicit
Dim Onlyloading As Boolean
Dim ConcealCaret As Boolean
Private Sub cboHLColor_Click()
    'change the back/forecolor according to the selected index
    cboHLColor.BackColor = GetHLColor(cboHLColor)
    If Not Onlyloading Then RTF.SetFocus
End Sub



Private Sub cboSyntax_Click(Index As Integer)
    'change the back/forecolor according to the selected index
    cboSyntax(Index).BackColor = GetHLColor(cboSyntax(Index))
    If Not Onlyloading Then RTF.SetFocus

End Sub

Private Sub cboWord_Click()
    'change the back/forecolor according to the selected index
    cboWord.BackColor = GetHLColor(cboWord)
    If Not Onlyloading Then RTF.SetFocus

End Sub

Private Sub ChHideCaret_Click()
    ConcealCaret = CBool(ChHideCaret.Value)
    If ConcealCaret Then
        HideCaret RTF.hwnd
    Else
        RTF.SetFocus
    End If
End Sub

Private Sub ChHook_Click()
    'hook/unhook the appropriate controls to remove/add
    'the focus rectangle
    Dim ctl As Control
    For Each ctl In Me.Controls
        If (TypeOf ctl Is CheckBox) Or (TypeOf ctl Is CommandButton) Or (TypeOf ctl Is OptionButton) Then
            If CBool(ChHook.Value) Then
                Hook ctl.hwnd
                RTF.SetFocus
            Else
                Unhook ctl.hwnd
            End If
        End If
    Next

End Sub

Private Sub ChSelectCharacter_Click()
    'select the character under the mouse
    'change the mousepointer to an arrow or it will flash
    'when you move it
    If CBool(ChSelectCharacter.Value) Then
        RTF.MousePointer = rtfArrow
    Else
        RTF.MousePointer = rtfDefault
    End If
End Sub

Private Sub ChWordWrap_Click()
    RTF.RightMargin = IIf(CBool(ChWordWrap.Value), 0, Screen.Width)
End Sub

Private Sub cmdHighlight_Click()
    'Wouldn't it be great if we could just do this...
    '    RTF.SelHighlightColor = vbGreen
    '    RTF.SelHighlight = True
    HighLightSelection Me, RTF, GetHLColor(cboHLColor)
End Sub

Private Sub cmdHighlightWord_Click()
    If Len(Trim(txtWord.Text)) = 0 Then Exit Sub
    HighLightWord Me, RTF, txtWord.Text, GetHLColor(cboWord), OptWord(1).Value

End Sub

Private Sub cmdResetColors_Click()
    Dim CurPos As Long
    LockWindowUpdate Me.hwnd
    CurPos = GetCurrentPosition(RTF)
    SelectAll Me, RTF, True
    RTF.SelColor = vbBlack
    RTF.SelLength = 0
    SetScrollPos Me, RTF, CurPos, True
    LockWindowUpdate 0
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Select Case Index
        Case 0: SelectAll Me, RTF
        Case 1: SelectAbove Me, RTF
        Case 2: SelectBelow Me, RTF
    End Select
End Sub

Private Sub cmdSyntax_Click(Index As Integer)
    Select Case Index
        Case 0
            HighLightComments Me, RTF, GetHLColor(cboSyntax(Index)), txtComments.Text
        Case 1
            StringColor Me, RTF, GetHLColor(cboSyntax(Index))
        Case 2
            KeyColor Me, RTF, GetHLColor(cboSyntax(Index)), GetVBKeyWords
    End Select
End Sub

Private Sub cmdUnHighlight_Click()
    'Wouldn't it be great if we could just do this...
    '    RTF.SelHighlight = False
    UnHighLight Me, RTF, , True
End Sub


Private Sub Form_Load()
    Onlyloading = True
    cboHLColor.ListIndex = 0
    cboSyntax(0).ListIndex = 2
    cboSyntax(1).ListIndex = 1
    cboSyntax(2).ListIndex = 5
    cboWord.ListIndex = 3
End Sub

Private Sub Form_Paint()
    Onlyloading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ctl As Control
    For Each ctl In Me.Controls
        If (TypeOf ctl Is CheckBox) Or (TypeOf ctl Is CommandButton) Or (TypeOf ctl Is OptionButton) Then
            Unhook ctl.hwnd
        End If
    Next
End Sub

Private Sub OptAutoHighlight_Click()
    cmdHighlight.Enabled = OptManualHighlight.Value
End Sub

Private Sub OptManualHighlight_Click()
    cmdHighlight.Enabled = OptManualHighlight.Value
End Sub

Private Sub RTF_GotFocus()
    'Hide the Caret
    If ConcealCaret Then HideCaret RTF.hwnd

End Sub

Private Sub RTF_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Hide the Caret
    If ConcealCaret Then HideCaret RTF.hwnd

End Sub

Private Sub RTF_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim P As POINTL, CurPos As Long, tmpChar As String
    On Error Resume Next
    'Hide the Caret
    If ConcealCaret Then HideCaret RTF.hwnd
    'Current position
    P.x = x / Screen.TwipsPerPixelX
    P.y = y / Screen.TwipsPerPixelY
    'Use the API to discover the character number
    CurPos = SendMessageP(RTF.hwnd, EM_CHARFROMPOS, 0, P)
    lblCharacterNumber.Caption = "Charater Pos: " & CurPos
    'The richtextbox will tell us the line number
    lblLineNumber.Caption = "Line: " & RTF.GetLineFromChar(CurPos)
    'What character is it?
    tmpChar = Mid(RTF.Text, CurPos + 1, 1)
    lblCharacter.Caption = "Charater: " & IIf(tmpChar = vbCr, "vbCr", tmpChar)
    If CBool(ChCaret.Value) Then
        'move the Caret as required
        RTF.SetFocus
        RTF.SelStart = CurPos
    End If
    If CBool(ChSelectCharacter.Value) Then
        'Select the character
        RTF.SetFocus
        RTF.SelStart = CurPos
        RTF.SelLength = 1
    End If
End Sub

Private Sub RTF_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Hide the Caret
    If ConcealCaret Then HideCaret RTF.hwnd
    If OptAutoHighlight.Value Then HighLightSelection Me, RTF, GetHLColor(cboHLColor)
End Sub

Public Function GetHLColor(mcbo As ComboBox) As Long
    'Return a color according to the listindex
    'and adjust the forecolor of the combobox accordingly
    Select Case mcbo.ListIndex
        Case 0: GetHLColor = vbYellow: mcbo.ForeColor = vbBlack
        Case 1: GetHLColor = vbRed: mcbo.ForeColor = vbWhite
        Case 2: GetHLColor = vbGreen: mcbo.ForeColor = vbBlack
        Case 3: GetHLColor = vbBlue: mcbo.ForeColor = vbWhite
        Case 4: GetHLColor = vbCyan: mcbo.ForeColor = vbBlack
        Case 5: GetHLColor = RGB(0, 0, 128): mcbo.ForeColor = vbWhite
    End Select
End Function
