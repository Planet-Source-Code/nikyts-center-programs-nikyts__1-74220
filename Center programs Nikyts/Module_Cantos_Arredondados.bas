Attribute VB_Name = "Module_Cantos_Arredondados"
'Option Explicit
Private Const PI = 22 / 7
Private Const RADIUS = 10
Private Const WM_SYSCOMMAND = &H112

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

'constants required by Shell_NotifyIcon API call:
Const NIM_ADD = &H0
Const NIM_MODIFY = &H1
Const NIM_DELETE = &H2
Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4
Const WM_MOUSEMOVE = &H200

Private Sub drawRoundRect(Optional ht As Single = 60, Optional wd As Single = 60, Optional ThreeD As Boolean = True, Optional left As Single = 100, Optional top As Single = 100, Optional r As Single = RADIUS)
    On Error Resume Next
    Dim linw As Single ' line width
    Dim linh As Single ' line height
    Dim X As Single ' x & y center coordinates of circle
    Dim Y As Single
    Dim col1 As Long, col2 As Long 'colors
    Dim prevdw As Integer
    linw = wd - (2 * r)
    linh = ht - (2 * r)
    X = left + r
    Y = top + r
    If ThreeD Then
        prevdw = DrawWidth
        DrawWidth = 2
        col1 = &HFFFFFF
        col2 = RGB(95, 95, 95)
    Else
        col1 = ForeColor
        col2 = col1
    End If
    
'    Line (X, top)-(X + linw, top), col1                  'top line
'    Circle (X, Y), r, col1, DegToRad(90), DegToRad(180)  'top left curve
'    Line (left, Y)-(left, Y + linh), col1                'left line
'    Circle (X, Y + linh), r, col1, DegToRad(180), DegToRad(225) 'left bottom curve 1
'    Circle (X, Y + linh), r, col2, DegToRad(225), DegToRad(270) 'left bottom curve 2
'    Line (X, top + ht)-(X + linw, top + ht), col2 ' bottom line
'    Circle (X + linw, Y + linh), r, col2, DegToRad(270), DegToRad(0)  'right bottom curve
'    Line (left + wd, Y)-(left + wd, Y + linh), col2    'right line
'    Circle (X + linw, Y), r, col2, DegToRad(0), DegToRad(45)  'right top curve1
'    Circle (X + linw, Y), r, col1, DegToRad(45), DegToRad(90)  'right top curve2
    DrawWidth = prevdw
End Sub

Private Function DegToRad(Deg As Single)
    ' PI radians = 180 deg
    ' 1 deg = PI / 180 rad
    DegToRad = Deg * (PI / 180)
End Function

Public Sub Arredondar_Cantos_do_Form(nome_form As Form, existe_frame_botoes As Boolean)
    'Arredondar os campos (top) do formulário
    Dim regn As Long
    nome_form.AutoRedraw = True
    regn = CreateRoundRectRgn(0, 0, nome_form.ScaleWidth, nome_form.Height, 7, 7)
    SetWindowRgn nome_form.hwnd, regn, True
    nome_form.Refresh
End Sub
