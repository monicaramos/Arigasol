Attribute VB_Name = "HoverButton"
Option Explicit

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef _
lpPoint As POINTAPI) As Long

Private Declare Function WindowFromPoint Lib "user32.dll" ( _
     ByVal xPoint As Long, _
     ByVal yPoint As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

' +-+- temp +-+-

' +-+-+-+-+-+-+-

Dim HooverActive As Boolean

Public Sub MakeHover(ByVal Button As Object)
    Dim PntApi As POINTAPI
    
    If HooverActive Then Exit Sub
    
    HooverActive = True
    SetEffect Button
    Do
        'get mousepointers position...
        If GetCursorPos(PntApi) <> 0 Then
            'wait until Mouse isnt over the button...
            If WindowFromPoint(PntApi.X, PntApi.Y) = Button.hwnd Then
            'If WindowFromPoint(PntApi.X, PntApi.Y) = Button.Picture.Handle Then
                'let the operating system process other events...
                DoEvents
            Else
                'if mouse isnt on the button any more exit this endless-loop...
                Exit Do
            End If
        End If
    Loop
    
    ' +-+- temp ++-
    
    ' +-+-+-+-+--++-
    
    HooverActive = False
    ResetEffect Button
End Sub



'Feel free to change the following lines
'to create your individuel hoovereffect...

Private Sub SetEffect(Button As Object)
    'because command-buttons do not support forecolor...
    On Error Resume Next
    
'    With Button
'
'        .BackColor = RGB(200, 200, 255)
'        .ForeColor = vbRed
'        .FontBold = True
'        .MouseIcon = LoadResPicture(101, vbResCursor)
'        .MousePointer = 99
'
'    End With
    Button.Picture = LoadPicture(App.Path & "\Images\bus\cua_selec.gif")
End Sub

Private Sub ResetEffect(Button As Object)
    'because command-buttons do not support forecolor...
    On Error Resume Next
    
'    With Button
'
'        .BackColor = vbMenuBar
'        .ForeColor = vbMenuText
'        .FontBold = False
'        .MousePointer = 0
'
'    End With
    'Button.Picture = LoadPicture("")
    Button.Picture = LoadPicture(App.Path & "\Images\bus\cua_selec.gif")
End Sub



