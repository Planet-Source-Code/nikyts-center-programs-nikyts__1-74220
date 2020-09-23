Attribute VB_Name = "Module_Geral"
Option Explicit

Public Const ULW_ALPHA = &H2
Public Const DIB_RGB_COLORS As Long = 0
Public Const AC_SRC_ALPHA As Long = &H1
Public Const AC_SRC_OVER = &H0
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE As Long = -20
Public Const HWND_TOPMOST As Long = -1
Public Const SWP_NOSIZE As Long = &H1
Public Const DEFAULT_QUALITY = 0
Public Const DEFAULT_PITCH = 0
Public Const DEFAULT_CHARSET = 1
Public Const OUT_DEFAULT_PRECIS = 0

Public Type BITMAPINFOHEADER
    Size As Long
    Width As Long
    Height As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    SizeImage As Long
    XPelsPerMeter As Long
    YPelsPerMeter As Long
    ClrUsed As Long
    ClrImportant As Long
End Type

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Public Enum GDIPLUS_ALIGNMENT
   StringAlignmentNear = 0
   StringAlignmentCenter = 1
   StringAlignmentFar = 2
End Enum

Public Enum GDIPLUS_COLORS
   AliceBlue = &HFFF0F8FF
   AntiqueWhite = &HFFFAEBD7
   Aqua = &HFF00FFFF
   Aquamarine = &HFF7FFFD4
   Azure = &HFFF0FFFF
   Beige = &HFFF5F5DC
   Bisque = &HFFFFE4C4
   Black = &HFF000000
   BlanchedAlmond = &HFFFFEBCD
   Blue = &HFF0000FF
   BlueViolet = &HFF8A2BE2
   Brown = &HFFA52A2A
   BurlyWood = &HFFDEB887
   CadetBlue = &HFF5F9EA0
   Chartreuse = &HFF7FFF00
   Chocolate = &HFFD2691E
   Coral = &HFFFF7F50
   CornflowerBlue = &HFF6495ED
   Cornsilk = &HFFFFF8DC
   Crimson = &HFFDC143C
   Cyan = &HFF00FFFF
   DarkBlue = &HFF00008B
   DarkBrown = &HFF804040
   DarkCyan = &HFF008B8B
   DarkGoldenrod = &HFFB8860B
   DarkGray = &HFFA9A9A9
   DarkGreen = &HFF006400
   DarkKhaki = &HFFBDB76B
   DarkMagenta = &HFF8B008B
   DarkOliveGreen = &HFF556B2F
   DarkOrange = &HFFFF8C00
   DarkOrchid = &HFF9932CC
   DarkRed = &HFF8B0000
   DarkSalmon = &HFFE9967A
   DarkSeaGreen = &HFF8FBC8B
   DarkSlateBlue = &HFF483D8B
   DarkSlateGray = &HFF2F4F4F
   DarkTurquoise = &HFF00CED1
   DarkViolet = &HFF9400D3
   DeepPink = &HFFFF1493
   DeepSkyBlue = &HFF00BFFF
   DimGray = &HFF696969
   DodgerBlue = &HFF1E90FF
   Firebrick = &HFFB22222
   FloralWhite = &HFFFFFAF0
   ForestGreen = &HFF228B22
   Fuchsia = &HFFFF00FF
   Gainsboro = &HFFDCDCDC
   GhostWhite = &HFFF8F8FF
   Gold = &HFFFFD700
   Goldenrod = &HFFDAA520
   Gray = &HFF808080
   Green = &HFF008000
   GreenYellow = &HFFADFF2F
   Honeydew = &HFFF0FFF0
   HotPink = &HFFFF69B4
   IndianRed = &HFFCD5C5C
   Indigo = &HFF4B0082
   Ivory = &HFFFFFFF0
   Khaki = &HFFF0E68C
   Lavender = &HFFE6E6FA
   LavenderBlush = &HFFFFF0F5
   LawnGreen = &HFF7CFC00
   LemonChiffon = &HFFFFFACD
   LightBlue = &HFFADD8E6
   LightCoral = &HFFF08080
   LightCyan = &HFFE0FFFF
   LightGoldenrodYellow = &HFFFAFAD2
   LightGray = &HFFD3D3D3
   LightGreen = &HFF90EE90
   LightPink = &HFFFFB6C1
   LightSalmon = &HFFFFA07A
   LightSeaGreen = &HFF20B2AA
   LightSkyBlue = &HFF87CEFA
   LightSlateGray = &HFF778899
   LightSteelBlue = &HFFB0C4DE
   LightYellow = &HFFFFFFE0
   Lime = &HFF00FF00
   LimeGreen = &HFF32CD32
   Linen = &HFFFAF0E6
   Magenta = &HFFFF00FF
   Maroon = &HFF800000
   MediumAquamarine = &HFF66CDAA
   MediumBlue = &HFF0000CD
   MediumOrchid = &HFFBA55D3
   MediumPurple = &HFF9370DB
   MediumSeaGreen = &HFF3CB371
   MediumSlateBlue = &HFF7B68EE
   MediumSpringGreen = &HFF00FA9A
   MediumTurquoise = &HFF48D1CC
   MediumVioletRed = &HFFC71585
   MidnightBlue = &HFF191970
   MintCream = &HFFF5FFFA
   MistyRose = &HFFFFE4E1
   Moccasin = &HFFFFE4B5
   NavajoWhite = &HFFFFDEAD
   Navy = &HFF000080
   OldLace = &HFFFDF5E6
   Olive = &HFF808000
   OliveDrab = &HFF6B8E23
   Orange = &HFFFFA500
   OrangeRed = &HFFFF4500
   Orchid = &HFFDA70D6
   PaleGoldenrod = &HFFEEE8AA
   PaleGreen = &HFF98FB98
   PaleTurquoise = &HFFAFEEEE
   PaleVioletRed = &HFFDB7093
   PapayaWhip = &HFFFFEFD5
   PeachPuff = &HFFFFDAB9
   Peru = &HFFCD853F
   Pink = &HFFFFC0CB
   Plum = &HFFDDA0DD
   PowderBlue = &HFFB0E0E6
   Purple = &HFF800080
   Red = &HFFFF0000
   RosyBrown = &HFFBC8F8F
   RoyalBlue = &HFF4169E1
   SaddleBrown = &HFF8B4513
   Salmon = &HFFFA8072
   SandyBrown = &HFFF4A460
   SeaGreen = &HFF2E8B57
   SeaShell = &HFFFFF5EE
   Sienna = &HFFA0522D
   Silver = &HFFC0C0C0
   SkyBlue = &HFF87CEEB
   SlateBlue = &HFF6A5ACD
   SlateGray = &HFF708090
   Snow = &HFFFFFAFA
   SpringGreen = &HFF00FF7F
   SteelBlue = &HFF4682B4
   Tan = &HFFD2B48C
   Teal = &HFF008080
   Thistle = &HFFD8BFD8
   Tomato = &HFFFF6347
   Transparent = &HFFFFFF
   Turquoise = &HFF40E0D0
   Violet = &HFFEE82EE
   Wheat = &HFFF5DEB3
   White = &HFFFFFFFF
   WhiteSmoke = &HFFF5F5F5
   XPBlue = &HFF003CC7
   XPGradient = &HFFC6C5D7
   XPGoldDark = &HFFB08218
   XPGoldLight = &HFFFCF9C3
   Yellow = &HFFFFFF00
   YellowGreen = &HFF9ACD32
End Enum

Public Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum

Public Type GDIPLUS_STARTINPUT
    GDIPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Public Enum GDIPLUS_UNIT
    UnitWorld
    UnitDisplay
    UnitPixel
    UnitPoint
    UnitInch
    UnitDocument
    UnitMillimeter
End Enum

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECTF
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Public Enum TASKBAR_POSITION
    vbBottom
    vbleft
    vbright
    vbTop
End Enum

Public Type BITMAPINFO
    bmpHeader As BITMAPINFOHEADER
    bmpColors As RGBQUAD
End Type

Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GdipCreateFont Lib "gdiplus" (ByVal fontFamily As Long, ByVal emSize As Single, ByVal style As GDIPLUS_FONTSTYLE, ByVal UNIT As GDIPLUS_UNIT, createdfont As Long) As Long
Public Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal name As String, ByVal fontCollection As Long, fontFamily As Long) As Long
Public Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hdc As Long, GpGraphics As Long) As Long
Public Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As GDIPLUS_COLORS, brush As Long) As Long
Public Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Public Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As Long
Public Declare Function GdipDeleteFont Lib "gdiplus" (ByVal curFont As Long) As Long
Public Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Public Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal graphics As Long) As Long
Public Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As Long) As Long
Public Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As Long) As Long
Public Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal graphics As Long, ByVal Img As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Public Declare Function GdipDrawString Lib "gdiplus" (ByVal graphics As Long, ByVal str As String, ByVal length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, ByVal brush As Long) As Long
Public Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As Long, Height As Long) As Long
Public Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As Long, Width As Long) As Long
Public Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal FileName As Long, GpImage As Long) As Long
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As Long
Public Declare Function GdiplusStartup Lib "gdiplus.dll" (Token As Long, gdipInput As GDIPLUS_STARTINPUT, GdiplusStartupOutput As Long) As Long
Public Declare Function GdipReleaseDC Lib "gdiplus.dll" (ByVal graphics As Long, ByVal hdc As Long) As Long
Public Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal graphics As Long, ByVal InterMode As Long) As Long
Public Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal align As GDIPLUS_ALIGNMENT) As Long
Public Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal align As GDIPLUS_ALIGNMENT) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long


'-----------------------------------------------------------------------------------------------------------
'API's para mover o formulários e carregar os skins
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
'Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
'Public Type POINTAPI
'        X As Long
'        Y As Long
'End Type
Public Const SRCCOPY = &HCC0020
Global iTPPY As Long
Global iTPPX As Long

'Variável das msgboxs
Public Resposta As String

'API para o procedimento alway's on top
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Variáveis para poder mover o formulário
Dim bMoveFrom As Boolean, LastPoint As POINTAPI

'Variável para indicar o caminho do ficheiro das opções do programa
Public Localizacao_Ficheiro_Preferencias As String
Public Localizacao_Ficheiro_Lingua As String
'Public Localizacao_Ficheiro_Skin As String

Sub Main()
    'Verificar se já existe algum instância da aplicação
    'If App.PrevInstance = True Then End
    Localizacao_Ficheiro_Preferencias = App.Path & "\Options\Properties.ini"
    'Localizacao_Ficheiro_Skin = App.Path & "\Skins\" & Form_Preferencias.Text_Skin.Text & "\Style.ini"
    
    On Error Resume Next
    'Verificar se o programa já foi instalado
    Dim Programa_Instalado As String: Programa_Instalado = ReadINI("Settings", "Installed_Program", Localizacao_Ficheiro_Preferencias)
    If Programa_Instalado = "False" Then
        Form_Setup.Show
    Else
        Form_Instalar.Show
    End If
End Sub

Public Sub Mensagem_de_Aviso(Aviso As String, Mensagem As String)
    'Procedimento para mostrar uma mensagem de aviso
    With Form_Mensagem
        If Aviso = "Information" Then
            .Pic_Mensagem.Picture = Form_Skin.Icon_Info.Picture
            .Botao_Ok.Visible = True
        ElseIf Aviso = "Error" Then
            .Pic_Mensagem.Picture = Form_Skin.Icon_Error.Picture
            .Botao_Ok.Visible = True
        ElseIf Aviso = "Question" Then
            .Pic_Mensagem.Picture = Form_Skin.Icon_Quest.Picture
            .Botao_Sim.Visible = True
            .Botao_Nao.Visible = True
        End If

        .Label_Mensagem.Caption = Mensagem
        .Show vbModal
    End With
End Sub


Public Sub Ajustar_ChecBox(Pic_CheckBox As PictureBox, CheckBox As CheckBox)
    'Procedimento para ajustar as checboxs do formulário
    With Pic_CheckBox
        .Height = CheckBox.Height
        .Width = CheckBox.Height
    End With
End Sub

Public Sub Ajustar_Option(Pic_Option As PictureBox)
    'Procedimento para ajustar as checboxs do formulário
    With Pic_Option
        .Height = Form_Skin.Opcao_Normal.Height
        .Width = Form_Skin.Opcao_Normal.Width
    End With
End Sub

'Colocar o formulário por cima dos outros
Sub AlwaysOnTop(FrmID As Form, OnTop As Integer)
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
    If OnTop = -1 Then
        OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
    Else
        OnTop = SetWindowPos(FrmID.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
    End If
End Sub

Public Function ArquivoExiste(ByVal Caminho As String, Optional ByVal SomenteDiretorio As Boolean = False) As Boolean
    'Função para verificar se a pasta existe
    On Error Resume Next
    If SomenteDiretorio Then
        ArquivoExiste = GetAttr(Mid(Caminho, 1, InStrRev(Caminho, ""))) And vbDirectory
    Else
        ArquivoExiste = GetAttr(Caminho)
    End If
    On Error GoTo 0
End Function

Public Sub Ajustar_Formulario(Form As Form, Icon_Visivel As Boolean, Form_Ajustavel As Boolean, Frame_Centro_Visivel As Boolean, _
                                Frame_Botoes_Visivel As Boolean)
    'Procedimento para ajustar os componentes dos formulários
    If Form.WindowState = 1 Then Exit Sub
    With Form.Shape_Contorno
        .Height = Form.ScaleHeight
        .Top = 0
        .Width = Form.ScaleWidth
        .Left = 0
    End With
    
    With Form.Barra_ControlBox
        .Height = Form_Skin.Fundo_Barra_ControlBox.Height
        .Top = 0 ' 1
        .Width = Form.ScaleWidth '- 2
        .Left = 0 ' 1
    End With
    
    With Form.Fundo_Barra_ControlBox
        .Stretch = True
        .Top = 0
        .Width = Form.Barra_ControlBox.ScaleWidth
        .Left = 0
    End With

    With Form.Label_Titulo
        .Top = (Form.Barra_ControlBox.ScaleHeight - .Height) / 2
        If Icon_Visivel = False Then .Left = 10 Else: .Left = 26
    End With
    
    'Botões do controlbox
    Dim Ajustar_Botoes As String
    Ajustar_Botoes = "False" 'ReadINI("Dimensions", "Adjust_Button_ControlBox", Localizacao_Ficheiro_Skin)
    
    With Form.Botao_Fechar
        .Height = Form_Skin.Botao_Fechar.Height
        If Ajustar_Botoes = "False" Then
            .Top = (Form.Barra_ControlBox.ScaleHeight - .Height) / 2
        Else
            .Top = 0
        End If
        .Width = Form_Skin.Botao_Fechar.Width
        .Left = Form.Barra_ControlBox.Width - .Width - 6
    End With
    
    If Form_Ajustavel = True Then
        With Form.Botao_Maximizar
            .Top = Form.Botao_Fechar.Top
            If Ajustar_Botoes = "False" Then
                .Left = Form.Botao_Fechar.Left - .Width - 8
            Else
                .Left = Form.Botao_Fechar.Left - .Width
            End If
        End With
        
        With Form.Botao_Restaurar
            .Top = Form.Botao_Fechar.Top
            .Left = Form.Botao_Maximizar.Left
        End With
        
        With Form.Botao_Minimizar
            .Top = Form.Botao_Fechar.Top
            If Ajustar_Botoes = "False" Then
                .Left = Form.Botao_Maximizar.Left - .Width - 8
            Else
                .Left = Form.Botao_Maximizar.Left - .Width
            End If
        End With
    End If
    
    If Frame_Botoes_Visivel = True Then
        With Form.Frame_Botoes
            .Height = Form_Skin.Fundo_Frame_Botoes.Height
            .Top = Form.ScaleHeight - .ScaleHeight - 1
            .Width = Form.ScaleWidth - 2
            .Left = 1
        End With
        
        With Form.Fundo_Frame_Botoes
            .Stretch = True
            .Top = 0
            .Width = Form.Frame_Botoes.ScaleWidth
            .Left = 0
        End With
    End If
    
    If Frame_Centro_Visivel = True Then
        With Form.Frame_Centro
            .Height = Form.ScaleHeight - Form.Barra_ControlBox.ScaleHeight - Form.Frame_Botoes.ScaleHeight - 2
            .Top = Form.Barra_ControlBox.Top + Form.Barra_ControlBox.ScaleHeight
            .Width = Form.ScaleWidth - 20
            .Left = 10
        End With
        
        With Form.Shape_Centro
            .Top = 0
            .Height = Form.Frame_Centro.Height
            .Left = 0
            .Width = Form.Frame_Centro.Width
            .Visible = True
        End With
    End If
End Sub

Public Sub Ajustar_Botao(Form As Form, Nome_Botao As PictureBox, Nome_Label As Label, Botao_Esta_Na_Frame_Botoes As Boolean, Nome_Shape As Shape)
    'Procedimento para ajustar os botoes e respectivas labels
    With Nome_Botao
        .Height = Form_Skin.Botao_Form.Height
        .Width = Form_Skin.Botao_Form.Width
        If Botao_Esta_Na_Frame_Botoes = True Then
            .Top = (Form.Frame_Botoes.ScaleHeight - .ScaleHeight) / 2
        End If
    End With
    
    With Nome_Label
        .AutoSize = False
        .Alignment = vbCenter
        .Top = (Nome_Botao.ScaleHeight - .Height) / 2
        .Width = Nome_Botao.ScaleWidth
        .Left = 0
    End With
    
    With Nome_Shape
        .Height = Nome_Botao.ScaleHeight
        .Top = 0
        .Width = Nome_Botao.ScaleWidth
        .Left = 0
    End With
End Sub

Public Sub Ajustar_Caixa_Texto(Barra_TextBox As PictureBox, Nome_TextBox As TextBox, Nome_Shape As Shape, Caixa_Observacoes As Boolean)
    'Procedimento para ajustar as caixas de texto
    If Caixa_Observacoes = False Then
        With Barra_TextBox
            .Height = Form_Skin.Caixa_de_Texto.Height
            .Width = Form_Skin.Caixa_de_Texto.Width
        End With
    Else
        With Barra_TextBox
            .Height = Form_Skin.Caixa_de_Observacoes.Height
            .Width = Form_Skin.Caixa_de_Observacoes.Width
        End With
    End If
    
    With Nome_TextBox
        .Height = Barra_TextBox.ScaleHeight - 8 - 8
        .Top = (Barra_TextBox.ScaleHeight - .Height) / 2
        .Width = Barra_TextBox.ScaleWidth - 8 - 8
        .Left = 8
    End With
    
    With Nome_Shape
        .Height = Barra_TextBox.ScaleHeight
        .Top = 0
        .Width = Barra_TextBox.ScaleWidth
        .Left = 0
    End With
End Sub

Public Sub Ajustar_Caixa_Texto_Mini(Caixa_Texto As PictureBox, Nome_TextBox As TextBox, Nome_Shape As Shape)
    'Procedimento para ajustar as caixas de texto
    With Caixa_Texto
        .Height = Form_Skin.Caixa_de_Texto_Mini.Height
        .Width = Form_Skin.Caixa_de_Texto_Mini.Width
    End With
    
    With Nome_TextBox
        .Height = Caixa_Texto.ScaleHeight - 8 - 8
        .Top = (Caixa_Texto.ScaleHeight - .Height) / 2
        .Width = Caixa_Texto.ScaleWidth - 8 - 8
        .Left = 8
    End With
    
    With Nome_Shape
        .Height = Caixa_Texto.ScaleHeight
        .Top = 0
        .Width = Caixa_Texto.ScaleWidth
        .Left = 0
    End With
End Sub

Public Sub Ajustar_Caixa_Texto_Media(Caixa_Texto As PictureBox, Nome_TextBox As TextBox, Nome_Shape As Shape)
    'Procedimento para ajustar as caixas de texto
    With Caixa_Texto
        .Height = Form_Skin.TextBox_Intermediate.Height
        .Width = Form_Skin.TextBox_Intermediate.Width
    End With
    
    With Nome_TextBox
        .Height = Caixa_Texto.ScaleHeight - 8 - 8
        .Top = (Caixa_Texto.ScaleHeight - .Height) / 2
        .Width = Caixa_Texto.ScaleWidth - 8 - 8
        .Left = 8
    End With
    
    With Nome_Shape
        .Height = Caixa_Texto.ScaleHeight
        .Top = 0
        .Width = Caixa_Texto.ScaleWidth
        .Left = 0
    End With
End Sub

Public Sub Mover_Formulario(Form As Form)
    'Procedimento para poder mover o formulário
    If Form.WindowState = 0 Then
        Dim iDX As Long, iDY As Long
        Dim POINT As POINTAPI
        If Not bMoveFrom Then Exit Sub
        GetCursorPos POINT
        iDX& = (POINT.X - LastPoint.X) * iTPPX&
        iDY& = (POINT.Y - LastPoint.Y) * iTPPY&
        LastPoint.X = POINT.X
        LastPoint.Y = POINT.Y
        Form.Move Form.Left + iDX&, Form.Top + iDY&
    End If
End Sub

Public Sub Capturar_Posicao_Formulario(Form As Form)
    'Capturar a posição de x e y
    Dim POINT As POINTAPI
    GetCursorPos POINT
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    bMoveFrom = True
End Sub

Public Sub Largar_Formulario(Form As Form)
    'Largar o formulário para a posição final
    bMoveFrom = False
End Sub

Public Sub Ajustar_Formulario_com_Menu(Form As Form, Icon_Visivel As Boolean, Form_Ajustavel As Boolean, Frame_Centro_Visivel As Boolean, _
                                Frame_Botoes_Visivel As Boolean)
    'Procedimento para ajustar os componentes dos formulários
    If Form.WindowState = 1 Then Exit Sub
    With Form.Shape_Contorno
        .Height = Form.ScaleHeight
        .Top = 0
        .Width = Form.ScaleWidth
        .Left = 0
    End With
    
    With Form.Barra_ControlBox
        .Height = Form_Skin.Fundo_Barra_ControlBox.Height
        .Top = 0 ' 1
        .Width = Form.ScaleWidth '- 2
        .Left = 0 ' 1
    End With
    
    With Form.Fundo_Barra_ControlBox
        .Stretch = True
        .Top = 0
        .Width = Form.Barra_ControlBox.ScaleWidth
        .Left = 0
    End With

    With Form.Label_Titulo
        .Top = (Form.Barra_ControlBox.ScaleHeight - .Height) / 2
        If Icon_Visivel = False Then .Left = 10 Else: .Left = 26
    End With
    
    'Botões do controlbox
    Dim Ajustar_Botoes As String
    Ajustar_Botoes = "False" 'ReadINI("Dimensions", "Adjust_Button_ControlBox", Localizacao_Ficheiro_Skin)
    
    With Form.Botao_Fechar
        .Height = Form_Skin.Botao_Fechar.Height
        If Ajustar_Botoes = "False" Then
            .Top = (Form.Barra_ControlBox.ScaleHeight - .Height) / 2
        Else
            .Top = 0
        End If
        .Width = Form_Skin.Botao_Fechar.Width
        .Left = Form.Barra_ControlBox.Width - .Width - 6
    End With
    
    If Form_Ajustavel = True Then
        With Form.Botao_Maximizar
            .Top = Form.Botao_Fechar.Top
            If Ajustar_Botoes = "False" Then
                .Left = Form.Botao_Fechar.Left - .Width - 8
            Else
                .Left = Form.Botao_Fechar.Left - .Width
            End If
        End With
        
        With Form.Botao_Restaurar
            .Top = Form.Botao_Fechar.Top
            .Left = Form.Botao_Maximizar.Left
        End With
        
        With Form.Botao_Minimizar
            .Top = Form.Botao_Fechar.Top
            If Ajustar_Botoes = "False" Then
                .Left = Form.Botao_Maximizar.Left - .Width - 8
            Else
                .Left = Form.Botao_Maximizar.Left - .Width
            End If
        End With
        
        With Form.Botao_Tray
            .Top = Form.Botao_Fechar.Top
            If Ajustar_Botoes = "False" Then
                .Left = Form.Botao_Minimizar.Left - .Width - 8
            Else
                .Left = Form.Botao_Minimizar.Left - .Width
            End If
        End With
    End If
    
    If Frame_Botoes_Visivel = True Then
        With Form.Frame_Botoes
            .Height = Form_Skin.Fundo_Frame_Botoes.Height
            .Top = Form.ScaleHeight - .ScaleHeight - 1
            .Width = Form.ScaleWidth - 2
            .Left = 1
        End With
        
        With Form.Fundo_Frame_Botoes
            .Stretch = True
            .Top = 0
            .Width = Form.Frame_Botoes.ScaleWidth
            .Left = 0
        End With
    End If
    
    If Frame_Centro_Visivel = True Then
        With Form.Frame_Centro
            .Height = Form.ScaleHeight - Form.Barra_ControlBox.ScaleHeight - Form.Frame_Botoes.ScaleHeight - 2 '- Form_Skin.Bar_Menu.Height
            .Top = Form.Barra_ControlBox.Top + Form.Barra_ControlBox.ScaleHeight '+ Form_Skin.Bar_Menu.Height
            .Width = Form.ScaleWidth - 20
            .Left = 10
        End With
        
        With Form.Shape_Centro
            .Top = 0
            .Height = Form.Frame_Centro.Height
            .Left = 0
            .Width = Form.Frame_Centro.Width
            .Visible = False
        End With
    End If
End Sub

Public Function DataArq(ByVal sArq As String) As String
    'Função para verificar a data de criação dos ficheiros
    If Dir$(sArq) <> "" Then
        DataArq = FileDateTime(sArq)
    Else
        DataArq = "ERRO" 'Não foi possivel identificar a data de criação do ficheiro
    End If
End Function


