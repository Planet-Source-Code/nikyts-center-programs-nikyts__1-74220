VERSION 5.00
Begin VB.UserControl NCombo 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2385
   ScaleHeight     =   375
   ScaleWidth      =   2385
   Begin VB.PictureBox Pic_Fundo 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   1980
      TabIndex        =   0
      Top             =   0
      Width           =   1980
      Begin VB.PictureBox Pic_Seta 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1200
         ScaleHeight     =   375
         ScaleWidth      =   465
         TabIndex        =   2
         Top             =   0
         Width           =   465
         Begin VB.Image Image_Seta 
            Height          =   315
            Left            =   0
            Picture         =   "NCombo.ctx":0000
            Top             =   0
            Width           =   285
         End
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   2400
      Picture         =   "NCombo.ctx":052E
      Top             =   480
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   315
      Left            =   2760
      Picture         =   "NCombo.ctx":0A5C
      Top             =   480
      Width           =   285
   End
   Begin VB.Image Image_Caixa_Texto 
      Height          =   375
      Left            =   2400
      Picture         =   "NCombo.ctx":0F8A
      Top             =   0
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00C0C0C0&
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   2250
   End
End
Attribute VB_Name = "NCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum TamamlaS
    Hayýr
    Evet
End Enum

Public Enum StyleSNe
    IlkHarfBüyük
    HepsiKüçük
    HepsiBüyük
    Normal
End Enum

Private Const TamamlaSeç As Integer = Evet
Private Const StyleSeç As Integer = Normal

Private StyleSeçim As StyleSNe

Dim TamamlaSeçim As TamamlaS
Dim C_BenDe As Boolean
Dim HarfKim As Integer

Dim BorderStyle_Over As OLE_COLOR
Dim BorderStyle_Normal As OLE_COLOR
Dim AZemin As OLE_COLOR
Dim DZemin As OLE_COLOR
Dim XSeçim As Integer

Public Event CChange()
Public Event CKeyDown(KeyCode As Integer, Shift As Integer)
Public Event CKeyPress(KeyAscii As Integer)
Public Event CKeyUp(KeyCode As Integer, Shift As Integer)
Public Event CMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event CMouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event CMouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event CGotFocus()
Public Event CLostFocus()
Public Event CClick()
Public Event CDblClick()

Private Sub Combo1_Change()
Dim EnBaþ As String
    RaiseEvent CChange
    'HarfleriAyarla HarfKim
            Select Case Style
                Case HepsiKüçük
                    If C_BenDe = True Then
                        Combo1.Text = LCase(Combo1.Text)
                        SendKeys "^{END}"
                        
                    End If
                Case HepsiBüyük
                    If C_BenDe = True Then
                        Combo1.Text = UCase(Combo1.Text)
                        SendKeys "^{END}"
                        
                    End If
                Case IlkHarfBüyük
                    If C_BenDe = True Then
                        If Len(Combo1.Text) > 1 Then
                            EnBaþ = left(Combo1.Text, 1)
                            Combo1.Text = EnBaþ & LCase(Right(Combo1.Text, Len(Combo1.Text) - 1))
                        End If
                        SendKeys "^{END}"
                        
                    End If
            End Select
End Sub

Private Sub Combo1_Click()
    RaiseEvent CClick
End Sub

Private Sub Combo1_DblClick()
    RaiseEvent CDblClick
End Sub

Private Sub Combo1_GotFocus()
    RaiseEvent CGotFocus
    C_BenDe = True
    'Combo1.BackColor = AZemin
    Shape_Contorno.BorderColor = BorderStyle_Over
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent CKeyDown(KeyCode, Shift)
        If KeyCode = vbKeyReturn Then
            SendKeys "^{HOME}"
            SendKeys "{TAB}"
        Else
            HarfKim = KeyCode
            HarfleriAyarla KeyCode
        End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    RaiseEvent CKeyPress(KeyAscii)
        Select Case Style
            Case HepsiKüçük
                If C_BenDe = True Then
                    Select Case KeyAscii
                        Case 73
                            KeyAscii = 253
                        Case 221
                            KeyAscii = 105
                        Case Else
                            KeyAscii = Asc(LCase(Chr(KeyAscii)))
                    End Select
                End If
            Case HepsiBüyük
                If C_BenDe = True Then
                    Select Case KeyAscii
                        Case 105
                            KeyAscii = 221
                        Case 253
                            KeyAscii = 73
                        Case Else
                            KeyAscii = Asc(UCase(Chr(KeyAscii)))
                    End Select
                End If
            Case IlkHarfBüyük
                If C_BenDe = True Then
                    If Len(Combo1.Text) = 0 Then 'Or Len(txtstandart.Text) = 1 Then
                        Select Case KeyAscii
                            Case 105
                                KeyAscii = 221
                            Case 253
                                KeyAscii = 73
                            Case Else
                                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                        End Select
                    Else
                        Select Case KeyAscii
                            Case 73
                                KeyAscii = 253
                            Case 221
                                KeyAscii = 105
                            Case Else
                                KeyAscii = Asc(LCase(Chr(KeyAscii)))
                        End Select
                    End If
                End If
        End Select
End Sub
Sub HarfleriAyarla(TuþNe As Integer)
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Long, PSTR As String
    RaiseEvent CKeyUp(KeyCode, Shift)
    If Style = Normal Then
    C_BenDe = False
        If AutoSize = Evet Then
            If KeyCode <> 8 And (KeyCode < 35 Or KeyCode > 40) And KeyCode <> 46 And KeyCode <> 13 Then
                PSTR = Combo1.Text
                For i = 0 To Combo1.ListCount - 1
                    If StrComp(PSTR, (left(Combo1.List(i), Len(PSTR))), vbTextCompare) = 0 Then
                        Combo1.Text = Combo1.List(i)
                                        'Lv1.ListItems.Item(i).Selected = True
                    Exit For
                    End If
                Next i
                Combo1.SelStart = Len(PSTR)
                Combo1.SelLength = Len(Combo1.Text) - Len(PSTR)
            End If
        End If
        C_BenDe = True
    End If
End Sub

Private Sub Combo1_LostFocus()
    RaiseEvent CLostFocus
    C_BenDe = False
    'Combo1.BackColor = DZemin
    Shape_Contorno.BorderColor = BorderStyle_Normal
End Sub

Private Sub Image_Seta_Click()
    Shape_Contorno.BorderColor = BorderStyle_Over
    Combo1.SetFocus
    SendKeys "{F4}"
End Sub

Private Sub Image_Seta_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbKeyLButton Then
        Image_Seta.Picture = Image2.Picture
    End If
End Sub

Private Sub Image_Seta_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbKeyLButton Then
        Image_Seta.Picture = Image3.Picture
    End If
End Sub

Private Sub UserControl_Initialize()
    UserControl.Width = 1900
    Shape_Contorno.BorderColor = &HC0C0C0    'Cinza
    Combo1.backcolor = vbWhite
    BorderStyle_Over = &HDFB000        'Azul
    BorderStyle_Normal = &HC0C0C0  'Cinza
    Style = Normal
    AZemin = vbWhite
    DZemin = vbWhite
End Sub
 
 Sub UserControl_Resize()
    'Desenhar a combo, ajustando os objectos
    With UserControl
        .Height = Image_Caixa_Texto.Height
    End With

    With Pic_Fundo
        .top = 10
        .Height = UserControl.Height - 30
        .left = 10
        .Width = UserControl.Width - 30
    End With

    With Combo1
        .top = -30
        .Width = UserControl.Width + 30
        .left = -30
    End With

    With Pic_Seta
        .Height = Image_Seta.Height
        .top = 0
        .Width = Image_Seta.Width
        .left = Pic_Fundo.Width - .Width
    End With

    With Image_Seta
        .top = 0
        .left = 0
    End With
    
    With Shape_Contorno
        .top = 0
        .Height = UserControl.Height
        .left = 0
        .Width = UserControl.Width
    End With
End Sub

Public Property Let BorderSyleGotFocus(nColor As OLE_COLOR)
    BorderStyle_Over = nColor
    PropertyChanged "BorderSyleGotFocus"
End Property

Public Property Get BorderSyleGotFocus() As OLE_COLOR
    BorderSyleGotFocus = BorderStyle_Over
End Property

Public Property Let BorderSyleLostFocus(nColor As OLE_COLOR)
    BorderStyle_Normal = nColor
    PropertyChanged "BorderSyleLostFocus"
End Property

Public Property Get BorderSyleLostFocus() As OLE_COLOR
    BorderSyleLostFocus = BorderStyle_Normal
End Property

Public Property Get BorderSyleNormal() As OLE_COLOR
On Error Resume Next
    BorderSyleNormal = Shape_Contorno.BorderColor
End Property

Public Property Let BorderSyleNormal(ByVal Rengi As OLE_COLOR)
On Error Resume Next
    Shape_Contorno.BorderColor = Rengi
    PropertyChanged "BorderSyleNormal"
End Property

Public Property Let BackColorAktif(nColor As OLE_COLOR)
    AZemin = nColor
    PropertyChanged "BackColorAktif"
End Property
Public Property Get BackColorAktif() As OLE_COLOR
    BackColorAktif = AZemin
End Property

Public Property Let BackColorPasif(nColor As OLE_COLOR)
    DZemin = nColor
    PropertyChanged "BackColorPasif"
End Property

Public Property Get BackColorPasif() As OLE_COLOR
    BackColorPasif = DZemin
End Property

Public Property Get backcolor() As OLE_COLOR
On Error Resume Next
    backcolor = Combo1.backcolor
End Property

Public Property Let backcolor(ByVal New_BackColor As OLE_COLOR)
On Error Resume Next
    Combo1.backcolor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Set FontName(nFont As StdFont)
    Set Combo1.Font = nFont
    PropertyChanged "FontName"
    UserControl_Resize
End Property

Public Property Get FontName() As StdFont
    Set FontName = Combo1.Font
End Property

Public Property Let ForeColor(nFontColor As OLE_COLOR)
    Combo1.ForeColor = nFontColor
    PropertyChanged "ForeColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Combo1.ForeColor
End Property

Public Property Let AutoSize(Seçim As TamamlaS)
    TamamlaSeçim = Seçim
    PropertyChanged "AutoSize"
End Property

Public Property Get AutoSize() As TamamlaS
    AutoSize = TamamlaSeçim
End Property

Public Property Get Style() As StyleSNe
On Error Resume Next
    Style = StyleSeçim
End Property

Public Property Let Style(c As StyleSNe)
On Error Resume Next
    Select Case c
        Case 0
            c = IlkHarfBüyük
        Case 1
            c = HepsiKüçük
        Case 2
            c = HepsiBüyük
        Case 3
            c = Normal
    End Select
    StyleSeçim = c
    
    PropertyChanged "Style"
End Property

Public Property Get hwnd() As String
On Error Resume Next
    hwnd = Combo1.hwnd
End Property

Public Property Get Text() As String
On Error Resume Next
    Text = Combo1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
On Error Resume Next
    Combo1.Text() = New_Text
    PropertyChanged "Text"
End Property

Public Property Get Locked() As Boolean
On Error Resume Next
    Locked = Combo1.Locked
End Property

Public Property Let Locked(ByVal New_Kapa As Boolean)
On Error Resume Next
    Combo1.Locked() = New_Kapa
    PropertyChanged "Locked"
    If Locked = True Then
        Combo1.TabStop = False
    Else
        Combo1.TabStop = True
    End If
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
  TamamlaSeçim = PropBag.ReadProperty("AutoSize", "Evet")
  Shape_Contorno.BorderColor = PropBag.ReadProperty("BorderSyleNormal")
  Combo1.backcolor = PropBag.ReadProperty("BackColor")
  BorderStyle_Over = PropBag.ReadProperty("BorderSyleGotFocus")
  BorderStyle_Normal = PropBag.ReadProperty("BorderSyleLostFocus")
  AZemin = PropBag.ReadProperty("BackColorAktif")
  DZemin = PropBag.ReadProperty("BackColorPasif")
  Set Combo1.Font = PropBag.ReadProperty("FontName")
  Combo1.ForeColor = PropBag.ReadProperty("ForeColor")
  Style = PropBag.ReadProperty("Style", StyleSeç)
  Combo1.Text = PropBag.ReadProperty("Text", "Metin")
  Combo1.Locked = PropBag.ReadProperty("Locked", False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
  PropBag.WriteProperty "AutoSize", TamamlaSeçim
  PropBag.WriteProperty "BackColor", Combo1.backcolor
  PropBag.WriteProperty "BackColorAktif", AZemin
  PropBag.WriteProperty "BackColorPasif", DZemin
  PropBag.WriteProperty "BorderSyleGotFocus", BorderStyle_Over
  PropBag.WriteProperty "BorderSyleLostFocus", BorderStyle_Normal
  PropBag.WriteProperty "BorderSyleNormal", Shape_Contorno.BorderColor
  PropBag.WriteProperty "FontName", Combo1.Font, Nothing
  PropBag.WriteProperty "ForeColor", Combo1.ForeColor, Nothing
  PropBag.WriteProperty "Style", Style, StyleSeç
  PropBag.WriteProperty "HWnd", Combo1.hwnd
  PropBag.WriteProperty "Text", Combo1.Text, "Metin"
  PropBag.WriteProperty "Locked", Combo1.Locked, False
End Sub

Public Sub AddItem(Item As Variant)
    Combo1.AddItem CStr(Item)
End Sub

Public Sub Clear()
    Combo1.Clear
End Sub

Public Sub Refresh()
    Combo1.Refresh
End Sub

Public Sub RemoveItem(index As Integer)
    Combo1.RemoveItem index
End Sub

Public Function List(x As Integer) As String
    List = Combo1.List(x)
End Function

Public Function ListCount()
    ListCount = Combo1.ListCount
End Function

