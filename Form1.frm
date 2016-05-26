VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4695
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox txtout 
      Height          =   2895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   360
      Width           =   4695
   End
   Begin VB.TextBox txtsbh 
      Height          =   270
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pattern As String, matches() As String, matchoffset As Long
Private bhmatcher As New bh

Private Sub Form_Load()
    bhmatcher.init
    txtsbh_Change
    txtout.selstart = Len(txtout.Text)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtsbh.Width = Me.ScaleWidth
    txtout.Height = Me.ScaleHeight - txtout.Top
    txtout.Width = Me.ScaleWidth
End Sub

Private Sub txtsbh_Change()
    Dim strinfo As String
    Dim xhi As Long
    matches = bhmatcher.match(pattern, matchoffset)
    If matches(0) = "" Then
        If matchoffset <> 0 Then matchoffset = 0
        pattern = Left(pattern, Len(pattern) - 1)
        Beep
        updatesbh
    Else
        For xhi = 0 To UBound(matches)
            strinfo = strinfo & Mid$("12345QWERT", xhi + 1, 1) & matches(xhi)
        Next xhi
        Me.Caption = strinfo
    End If
End Sub

Private Sub txtsbh_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sbh As String
    sbh = pattern
    Debug.Print KeyCode
    Select Case KeyCode
    Case vbKeyI
        sbh = sbh & "1"
        matchoffset = 0
    Case vbKeyO
        sbh = sbh & "2"
        matchoffset = 0
    Case vbKeyP
        sbh = sbh & "3"
        matchoffset = 0
    Case vbKeyK
        sbh = sbh & "4"
        matchoffset = 0
    Case vbKeyL
        sbh = sbh & "5"
        matchoffset = 0
    Case 186
        sbh = sbh & "6"
        matchoffset = 0
    Case 188
        sbh = sbh & "7"
        matchoffset = 0
    Case 190
        sbh = sbh & "8"
        matchoffset = 0
    Case 191
        sbh = sbh & "9"
        matchoffset = 0
    Case 222
        sbh = sbh & "0"
        matchoffset = 0
    Case vbKeyBack
        If sbh <> "" Then
            sbh = Left$(sbh, Len(sbh) - 1)
            matchoffset = 0
        ElseIf txtout.selstart <> 0 Then
            Dim selstart As Long
            txtout.selstart = txtout.selstart - 1
            selstart = txtout.selstart
            txtout.Text = Left$(txtout.Text, txtout.selstart) & Right$(txtout.Text, Len(txtout.Text) - txtout.selstart - 1)
            txtout.selstart = selstart
        End If
    Case vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKeyQ, vbKeyW, vbKeyE, vbKeyR, vbKeyT
        Dim idx As Long
        Select Case KeyCode
            Case vbKey1: idx = 0
            Case vbKey2: idx = 1
            Case vbKey3: idx = 2
            Case vbKey4: idx = 3
            Case vbKey5: idx = 4
            Case vbKeyQ: idx = 5
            Case vbKeyW: idx = 6
            Case vbKeyE: idx = 7
            Case vbKeyR: idx = 8
            Case vbKeyT: idx = 9
        End Select
        txtout.SelText = matches(idx)
        txtout.SelLength = 0
        sbh = ""
        matchoffset = 0
    Case vbKeyLeft
        If txtout.selstart <> 0 Then txtout.selstart = txtout.selstart - 1
    Case vbKeyRight
        If txtout.selstart <> Len(txtout.Text) Then txtout.selstart = txtout.selstart + 1
    Case vbKeyEscape
        sbh = ""
    Case vbKeyJ
        matchoffset = matchoffset + bhmatcher.nmatch
        txtsbh_Change
    Case vbKeyU
        If matchoffset > 0 Then matchoffset = matchoffset - bhmatcher.nmatch
        txtsbh_Change
    Case vbKeyReturn
        txtout.SelText = vbCrLf
        txtout.SelLength = 0
    Case Else
    
    End Select
    pattern = sbh
    updatesbh
End Sub

Private Sub updatesbh()
    Dim sp As String
    sp = pattern
    sp = Replace(sp, "1", "Ò»")
    sp = Replace(sp, "2", "Ø­")
    sp = Replace(sp, "3", "Ø¯")
    sp = Replace(sp, "4", "Ø¼")
    sp = Replace(sp, "5", "ÒÒ")
    sp = Replace(sp, "6", "¶þ")
    sp = Replace(sp, "7", "Ê®")
    sp = Replace(sp, "8", "Ú¢")
    sp = Replace(sp, "9", "Ùï")
    sp = Replace(sp, "0", "¡£")
    
    txtsbh.Text = sp
    txtsbh.selstart = Len(sp)
End Sub
