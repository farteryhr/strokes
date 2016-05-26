VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private libsbh() As String, libhz() As String
Public nmatch As Long

Public Function init()
    Dim inline As String, asplit() As String
    Dim colsbh As New Collection, colhz As New Collection
    Dim xhi As Long, xhm As Long, xhv As Variant
    Dim arrm() As Byte
    Open App.Path + "\code.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, inline
        If inline = "" Then GoTo continue
        asplit = Split(inline, " ")
        If UBound(asplit) <> 1 Then GoTo continue
        
        ReDim arrm(Len(asplit(0)) - 1)
        For xhm = 0 To UBound(arrm)
            arrm(xhm) = CByte(Mid$(asplit(0), xhm + 1, 1))
        Next xhm
        
        colsbh.Add arrm
        colhz.Add asplit(1)
continue:
    Loop
    ReDim libsbh(colsbh.Count - 1), libhz(colhz.Count - 1)
    xhi = 0
    For Each xhv In colsbh
        libsbh(xhi) = xhv
        xhi = xhi + 1
    Next xhv
    xhi = 0
    For Each xhv In colhz
        libhz(xhi) = xhv
        xhi = xhi + 1
    Next xhv
    
    nmatch = 10
End Function

Public Function match(sbh As String, offset As Long) As String()
    Dim xhi As Long, imatch As Long
    Dim xhp As Long, arrp() As Byte, arrm() As Byte
    Dim bmatch As Boolean
    Dim res() As String
    ReDim res(nmatch - 1)
    imatch = -offset
    If sbh <> "" Then
        ReDim arrp(Len(sbh) - 1)
        For xhp = 0 To UBound(arrp)
            arrp(xhp) = CByte(Mid$(sbh, xhp + 1, 1))
        Next xhp
    End If
    For xhi = 0 To UBound(libsbh)
        bmatch = False
        arrm = libsbh(xhi)
        If sbh = "" Then bmatch = True Else bmatch = matchin(arrm, arrp)
        If bmatch Then
            If imatch >= 0 Then res(imatch) = libhz(xhi)
            imatch = imatch + 1
            If imatch = nmatch Then Exit For
        End If
    Next xhi
    match = res
End Function

Private Function matchin(arrm() As Byte, arrp() As Byte) As Boolean
    Dim xhm As Long, xhp As Long

    xhm = 0
    For xhp = 0 To UBound(arrp)
        If arrp(xhp) = 0 Then
            If xhm > UBound(arrm) Then GoTo win Else GoTo die
        ElseIf arrp(xhp) <= 5 Then
            If xhm > UBound(arrm) Then GoTo die
            If arrm(xhm) <> arrp(xhp) Then GoTo die
            xhm = xhm + 1
        Else
            If xhm + 1 > UBound(arrm) Then GoTo die
            Select Case arrp(xhp)
            Case 6
                If arrm(xhm) <> arrm(xhm + 1) Then GoTo die
            Case 7
                If Not ( _
                    (arrm(xhm) = 1 And arrm(xhm + 1) = 2) Or _
                    (arrm(xhm) = 2 And arrm(xhm + 1) = 1) Or _
                    (arrm(xhm) = 3 And arrm(xhm + 1) = 4) Or _
                    (arrm(xhm) = 4 And arrm(xhm + 1) = 3) _
                ) Then GoTo die
            Case 8
                If Not ( _
                    (arrm(xhm) = 2 And arrm(xhm + 1) = 5) Or _
                    (arrm(xhm) = 4 And arrm(xhm + 1) = 5) _
                ) Then GoTo die
            Case 9
                If Not (arrm(xhm) = 4 And arrm(xhm + 1) = 1) Then GoTo die
            End Select
            xhm = xhm + 2
        End If
    Next xhp
win:
    matchin = True
    Exit Function
die:
    matchin = False
End Function
