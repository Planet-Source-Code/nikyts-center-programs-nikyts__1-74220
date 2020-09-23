VERSION 5.00
Begin VB.Form Form_Setup 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   0  'None
   Caption         =   "Center programs Nikyts"
   ClientHeight    =   6645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8985
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Setup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   443
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   599
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File_Lingua 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      Height          =   810
      Left            =   7200
      Pattern         =   "*.lng"
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.DirListBox Dir_Lingua 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      Height          =   765
      Left            =   7200
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Frame_Centro 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   0
      ScaleHeight     =   361
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   7500
      Begin VB.PictureBox Barra_Mapa 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00DFB000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3000
         Left            =   0
         ScaleHeight     =   200
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   500
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Width           =   7500
         Begin VB.Shape Shape_Pais 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000080FF&
            Height          =   105
            Left            =   3960
            Shape           =   2  'Oval
            Top             =   1320
            Visible         =   0   'False
            Width           =   105
         End
      End
      Begin VB.PictureBox Lista_Linguas 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   1080
         ScaleHeight     =   63
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   303
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3870
         Visible         =   0   'False
         Width           =   4575
         Begin VB.Label Label_Lingua 
            BackColor       =   &H00EEEEEE&
            BackStyle       =   0  'Transparent
            Caption         =   "Idioma"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   0
            Width           =   960
         End
         Begin VB.Label Shape_Sombra 
            BackColor       =   &H00DFB000&
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   3975
         End
      End
      Begin VB.PictureBox Pic_Atalho 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         Picture         =   "Form_Setup.frx":57E2
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   4440
         Width           =   195
      End
      Begin VB.PictureBox Barra_Text_Lingua 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   120
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3480
         Width           =   5475
         Begin VB.PictureBox Seta_Lingua 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   5040
            Picture         =   "Form_Setup.frx":5A2C
            ScaleHeight     =   19
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   19
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   285
         End
         Begin VB.TextBox Text_Lingua 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   30
            Width           =   2940
         End
         Begin VB.Shape Contorno_Lingua 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00DFB000&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.CheckBox Check_Atalho 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Caption         =   "Criar atalho no ambiente de trabalho"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   1
         Top             =   4440
         Width           =   5400
      End
      Begin VB.Label Label_Idioma 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Idioma"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Width           =   600
      End
   End
   Begin VB.PictureBox Frame_Botoes 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00DFDFDF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6000
      Width           =   6015
      Begin VB.PictureBox Botao_Ok 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   2040
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   2
         Top             =   120
         Width           =   1740
         Begin VB.Shape Contorno_Ok 
            BorderColor     =   &H00DFB000&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label_Ok 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ok"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   750
            TabIndex        =   8
            Top             =   45
            Width           =   240
         End
      End
      Begin VB.PictureBox Botao_Cancelar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   3960
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   3
         Top             =   120
         Width           =   1740
         Begin VB.Shape Contorno_Cancelar 
            BorderColor     =   &H00DFB000&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label_Cancelar 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   465
            TabIndex        =   7
            Top             =   45
            Width           =   780
         End
      End
      Begin VB.Image Fundo_Frame_Botoes 
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox Barra_ControlBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H002B2B2B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   441
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   6615
      Begin VB.TextBox Text_Idioma 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image Botao_Fechar 
         Height          =   195
         Left            =   6120
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   195
      End
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Center programs Nikyts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   75
         TabIndex        =   5
         Top             =   120
         Width           =   2325
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00212121&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form_Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Nikyts Player
'   Copyright © 2011-2012 Nikyts software ™ - Informática e tecnologia
'   www.nikyts.com / nikyts@hotmail.com
'   Desenvolvido por: Nelson do Carmo
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Declaração das variáveis
'Dim bMoveFrom As Boolean, LastPoint As POINTAPI

'Variável para indicar qual a linha que está selecionada da lista linguas
Dim Linha_Selecionada As Integer

Private Sub Barra_ControlBox_Click()
    'Ocultar frame
    Lista_Linguas.Visible = False
End Sub

Private Sub Botao_Cancelar_Click()
    'Atalho para
    Label_Cancelar_Click
End Sub

Private Sub Botao_Cancelar_GotFocus()
    'Colocar o focus no botao
    Contorno_Cancelar.Visible = True
End Sub

Private Sub Botao_Cancelar_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Cancelar_Click
    If KeyCode = vbKeyLeft Then Botao_Cancelar_LostFocus: Botao_Ok_GotFocus: Botao_Ok.SetFocus
End Sub

Private Sub Botao_Cancelar_LostFocus()
    'Remover o focus no botao
    Contorno_Cancelar.Visible = False
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar formulário
    Unload Me
    End
End Sub

Private Sub Botao_Ok_Click()
    'Atalho para
    Label_Ok_Click
End Sub

Private Sub Botao_Ok_GotFocus()
    'Colocar o focus no botao
    Contorno_Ok.Visible = True
End Sub

Private Sub Botao_Ok_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Ok_Click
    If KeyCode = vbKeyRight Then Botao_Ok_LostFocus: Botao_Cancelar_GotFocus: Botao_Cancelar.SetFocus
End Sub

Private Sub Botao_Ok_LostFocus()
    'Ao perder o focus no botao
    Contorno_Ok.Visible = False
End Sub

Private Sub Check_Atalho_Click()
    'Des/Activar a opcção
    If Check_Atalho.Value = 1 Then
        Pic_Atalho.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Pic_Atalho.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Form_Click()
    'Ocultar lista
    Lista_Linguas.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Unload Me: End
End Sub

Private Sub Form_Load()
    'Iniciar o formulário
    Barra_Mapa.Picture = Form_Skin.Mapa_Mundo.Picture
    Text_Idioma.Text = ReadINI("Settings", "Language", Localizacao_Ficheiro_Preferencias)
    Carregar_Idioma
    Desenhar_Formulario
    Carregar_Skin
    Verificar_Pastas
    
    'Variáveis para poder mover o formulário
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    Arredondar_Cantos_do_Form Me, True

    'Carregar idiomas disponiveis
    Dir_Lingua.Path = App.Path & "\Languages\"
    File_Lingua.Path = Dir_Lingua.Path
    File_Lingua.Pattern = "*.lng"
    
    'Criar a lista consoante o nº de idiomas disponiveis
    Label_Lingua(0).Caption = ""
    Label_Lingua(0).Visible = True
    Dim Objecto As Integer
    For Objecto = 1 To File_Lingua.ListCount - 1
        Load Label_Lingua(Objecto)
        Label_Lingua(Objecto).Move Label_Lingua(Objecto - 1).left, Label_Lingua(Objecto - 1).top + Label_Lingua(Objecto - 1).Height
        Label_Lingua(Objecto).Visible = True
        
        Load Shape_Sombra(Objecto)
        Shape_Sombra(Objecto).Move Shape_Sombra(Objecto - 1).left, Shape_Sombra(Objecto - 1).top + Shape_Sombra(Objecto - 1).Height
        Shape_Sombra(Objecto).Visible = False
        Shape_Sombra(Objecto).ZOrder 1
    Next Objecto
    Lista_Linguas.Height = Shape_Sombra.Count * Shape_Sombra(0).Height
        
    'Preencher as label's com as linguas disponiveis
    Dim Z As Integer
    File_Lingua.ListIndex = 0
    For Z = 0 To File_Lingua.ListCount - 1
        Label_Lingua(Z).Caption = left$(File_Lingua.List(Z), InStr(File_Lingua.List(Z), ".") - (1)) 'Retirar a extensão do ficheiro ".lng"
    Next Z
    
    'Carregar preferências do programa
    Text_Lingua.Text = ReadINI("Settings", "Language", Localizacao_Ficheiro_Preferencias)
    
    'Chamar o procedimento
    Carregar_Idioma
    
    'Selecionar a 1ªlinha da lista linguas
    Linha_Selecionada = 0
    Shape_Sombra(0).Visible = True
    Label_Lingua(0).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
End Sub

Public Sub Carregar_Skin()
    'Procedimento para carregar o skin escolhido
    With Form_Skin
        Me.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Shape_Contorno.BorderColor = .Cor_Form_BorderColor.backcolor
        Frame_Centro.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Barra_Mapa.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Fundo_Barra_ControlBox.Picture = .Fundo_Barra_ControlBox.Picture
        Label_Titulo.ForeColor = .Cor_Label_Barra_Titulo.backcolor
        Botao_Fechar.Picture = .Botao_Fechar.Picture
        Label_Idioma.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Pic_Atalho.Picture = .Check_Normal.Picture
        Pic_Atalho.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Atalho.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Check_Atalho.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Fundo_Frame_Botoes.Picture = .Fundo_Frame_Botoes.Picture
        Label_Ok.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Label_Cancelar.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Ok.Picture = .Pic_Button.Picture
        Botao_Cancelar.Picture = .Pic_Button.Picture
        Contorno_Ok.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Cancelar.BorderColor = .Cor_Contorno_Caixas.backcolor
        Seta_Lingua.Picture = .Seta_Combo.Picture
        Barra_Text_Lingua.backcolor = .Cor_Fundo_Textbox.backcolor
        Barra_Text_Lingua.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Lingua.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Lingua.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Lingua.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Lingua.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Contorno_Lingua.BorderColor = .Cor_Contorno_Caixas.backcolor
        Text_Lingua.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Lingua.ForeColor = .Cor_Letra_Textbox.backcolor
        Lista_Linguas.backcolor = .Cor_Fundo_Textbox.backcolor
        Shape_Sombra(0).backcolor = .Cor_Contorno_Caixas.backcolor
        Label_Lingua(0).ForeColor = .Cor_Letra_Textbox.backcolor
    End With
End Sub

Public Sub Desactivar_Objectos()
    'as textboxs
    Text_Mensagem.Enabled = False
End Sub

Public Sub Activar_Objectos()
    'as textboxs
    Text_Mensagem.Enabled = True
End Sub

Private Sub Form_Resize()
    'Iniciar o formulário
    Desenhar_Formulario
End Sub

Private Sub Frame_Botoes_Click()
    'Ocultar frame
    Lista_Linguas.Visible = False
End Sub

Private Sub Frame_Centro_Click()
    'Ocultar frame
    Lista_Linguas.Visible = False
End Sub

Private Sub Label_Cancelar_Click()
    'Atalho para
    Botao_Fechar_Click
End Sub

Private Sub Label_Lingua_Click(Index As Integer)
    'Indicar a lingua selecionada pelo utilizador
    Text_Lingua.Text = Label_Lingua(Index).Caption
    Text_Idioma.Text = Label_Lingua(Index).Caption
    
    'Chamar o procedimento
    Carregar_Idioma
    
    Lista_Linguas.Visible = False
    Text_Lingua.SetFocus
End Sub

Private Sub Label_Lingua_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada = Index Then Exit Sub
    Shape_Sombra(Linha_Selecionada).Visible = False
    Label_Lingua(Linha_Selecionada).ForeColor = Form_Skin.Cor_Letra_Textbox.backcolor
    Shape_Sombra(Index).Visible = True
    Label_Lingua(Index).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
    Linha_Selecionada = Index
End Sub

Private Sub Label_Ok_Click()
    'Guardar nas opções o idioma escolhido
    Call WriteINI("Settings", "Language", Text_Lingua.Text, (Localizacao_Ficheiro_Preferencias))
    Text_Lingua.Text = Text_Lingua.Text
    
    'Verificar se a opção criar shortcut está selecionado
    If Check_Atalho.Value = 1 Then
        Dim lobj_Atalho As IWshRuntimeLibrary.IWshShortcut 'Reference > Windows script host object model
        Dim WshShell As New IWshRuntimeLibrary.WshShell

        Dim desktop As String: desktop = CreateObject("WScript.Shell").SpecialFolders("Desktop")
        Set lobj_Atalho = WshShell.CreateShortcut(desktop & "\" & App.ProductName & ".lnk")

        lobj_Atalho.TargetPath = App.Path & "\" & App.EXEName & ".exe" 'ProductName & ".exe"  '"C:\pasta\programaaserexecutadopeloatalho.exe"
        lobj_Atalho.WindowStyle = 1
        lobj_Atalho.Description = "Descrição do Atalho"
        lobj_Atalho.WorkingDirectory = App.Path '"C:\pasta\"
        lobj_Atalho.IconLocation = lobj_Atalho.TargetPath & ", 0"  '"C:\pasta\programaaserexecutadopeloatalho.exe, 0"
        lobj_Atalho.save
    End If
    
    On Error Resume Next
    Call WriteINI("Settings", "Installed_Program", "True", (Localizacao_Ficheiro_Preferencias))
    Form_Instalar.Show
    Unload Form_Setup
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Setup
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Setup
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Setup
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Setup
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Setup
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Setup
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    
    Barra_Mapa.Width = Form_Skin.Mapa_Mundo.Width
    Barra_Mapa.Height = Form_Skin.Mapa_Mundo.Height
    With Me
        .Width = Screen.TwipsPerPixelX * (Barra_Mapa.ScaleWidth)
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Barra_Mapa.ScaleHeight + 10 _
                    + Label_Idioma.Height + 3 + Form_Skin.Caixa_de_Texto.Height + (2 * Form_Skin.Caixa_de_Texto.Height) + _
                    (3 * Fundo_Frame_Botoes.Height))
    End With
    
    Ajustar_Formulario Form_Setup, False, False, False, True
    
    With Frame_Centro
        .top = Barra_ControlBox.top + Barra_ControlBox.ScaleHeight
        .Height = Me.ScaleHeight - Barra_ControlBox.ScaleHeight - Frame_Botoes.ScaleHeight - 2
        .left = 1
        .Width = Me.ScaleWidth - 3
    End With
    
    Ajustar_Botao Form_Setup, Botao_Cancelar, Label_Cancelar, True, Contorno_Cancelar
    Ajustar_Botao Form_Setup, Botao_Ok, Label_Ok, True, Contorno_Ok
    
    With Botao_Cancelar
        .left = Frame_Botoes.ScaleWidth - .ScaleWidth - .top
    End With
    With Botao_Ok
        .left = Botao_Cancelar.left - .ScaleWidth - .top
    End With

    Ajustar_Caixa_Texto Barra_Text_Lingua, Text_Lingua, Contorno_Lingua, False
    
    With Label_Idioma
        .top = Barra_Mapa.top + Barra_Mapa.ScaleHeight + 20
        .left = 20
    End With
    
    With Barra_Text_Lingua
        .Height = Form_Skin.Caixa_de_Texto.Height
        .top = Label_Idioma.top + Label_Idioma.Height + 3
        .Width = Me.ScaleWidth - (2 * Label_Idioma.left) 'Form_Skin.Caixa_de_Texto.Width
        .left = Label_Idioma.left
    
        Contorno_Lingua.Width = .Width
        Seta_Lingua.left = .ScaleWidth - Seta_Lingua.Width - Seta_Lingua.top
        Lista_Linguas.Width = .ScaleWidth
    End With
    
    With Seta_Lingua
        .Height = Form_Skin.Seta_Combo.Height
        .top = (Barra_Text_Lingua.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Seta_Combo.Width
        .left = Barra_Text_Lingua.ScaleWidth - .ScaleWidth - .top
    End With

    With Lista_Linguas
        .top = Barra_Text_Lingua.top + Barra_Text_Lingua.ScaleHeight
        .Width = Barra_Text_Lingua.ScaleWidth
        .left = Barra_Text_Lingua.left
    End With
    
    With Shape_Sombra(0)
        .Width = Lista_Linguas.ScaleWidth
    End With
    
    With Label_Lingua(0)
        .Width = Lista_Linguas.ScaleWidth
    End With
    
    With Barra_Mapa
        .Height = Form_Skin.Mapa_Mundo.Height
        .top = 0
        .Width = Form_Skin.Mapa_Mundo.Width
        .left = 0
    End With
        
    Ajustar_ChecBox Pic_Atalho, Check_Atalho
    
    With Check_Atalho
        .top = Barra_Text_Lingua.top + (Barra_Text_Lingua.ScaleHeight + 12)
        .Width = Barra_Text_Lingua.ScaleWidth
        .left = Label_Idioma.left
    End With
    
    With Pic_Atalho
        .top = Check_Atalho.top
        .left = Label_Idioma.left
    End With
    
    With Shape_Sombra(0)
        .Width = Lista_Linguas.Width
        .left = 0
    End With
    
    'Ajustar os objectos depois de arredondar os cantos do formulário
    Shape_Contorno.left = 0
    Shape_Contorno.Width = Me.ScaleWidth - 1
    Frame_Botoes.Width = Frame_Botoes.ScaleWidth - 1
End Sub

Private Sub Pic_Atalho_Click()
    'Des/Activar a opcção
    If Check_Atalho.Value = 0 Then
        Check_Atalho.Value = 1
        Pic_Atalho.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Check_Atalho.Value = 0
        Pic_Atalho.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Seta_Lingua_Click()
    'Ver/ocultar lista
    If Lista_Linguas.Visible = True Then
        Lista_Linguas.Visible = False
    Else
        Lista_Linguas.Visible = True
    End If
End Sub

Private Sub Shape_Sombra_Click(Index As Integer)
    'Atalho para
    Label_Lingua_Click (Index)
End Sub

Private Sub Shape_Sombra_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada = Index Then Exit Sub
    Shape_Sombra(Linha_Selecionada).Visible = False
    Label_Lingua(Linha_Selecionada).ForeColor = Form_Skin.Cor_Letra_Textbox.backcolor
    Shape_Sombra(Index).Visible = True
    Label_Lingua(Index).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
    Linha_Selecionada = Index
End Sub

Private Sub Text_Lingua_Click()
    'Ocultar lista
    Lista_Linguas.Visible = False
End Sub

Private Sub Text_Lingua_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Lingua.Visible = True
End Sub

Private Sub Text_Lingua_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Label_Ok_Click
End Sub

Private Sub Text_Lingua_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Lingua.Visible = False
End Sub

Public Sub Verificar_Pastas()
    'Procedimento para verificar se as pastas utilizadas pelo programa existem
    If Not ArquivoExiste(App.Path & "\Components", True) Then
        MkDir App.Path & "\Components\"
    End If
    
    If Not ArquivoExiste(App.Path & "\Languages", True) Then
        MkDir App.Path & "\Languages\"
    End If
    
    If Not ArquivoExiste(App.Path & "\Options", True) Then
        MkDir App.Path & "\Options\"
    End If
    
    If Not ArquivoExiste(App.Path & "\Programs", True) Then
        MkDir App.Path & "\Programs\"
    End If
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    On Error GoTo Corrige_Erro
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Text_Idioma.Text & ".lng"
    
    Botao_Fechar.ToolTipText = ReadINI("Setup", "Button_Close", Localizacao_Ficheiro_Lingua)
'    Label_Defina.Caption = ReadINI("Setup", "Label_Set", Localizacao_Ficheiro_Lingua)
    Label_Idioma.Caption = ReadINI("Setup", "Label_Language", Localizacao_Ficheiro_Lingua)
    Label_Ok.Caption = ReadINI("Setup", "Button_Ok", Localizacao_Ficheiro_Lingua)
    Label_Cancelar.Caption = ReadINI("Setup", "Button_Cancel", Localizacao_Ficheiro_Lingua)
    Check_Atalho.Caption = ReadINI("Setup", "Check_Shortcut", Localizacao_Ficheiro_Lingua)
    
    'Actualizar mapa
    Dim Regiao As String: Regiao = ReadINI("Map", "Region", Localizacao_Ficheiro_Lingua)
    Barra_Mapa.Picture = Form_Skin.Imagem_Regiao(Regiao).Picture
    
    Dim X, Y As String
    X = ReadINI("Map", "X", Localizacao_Ficheiro_Lingua)
    Y = ReadINI("Map", "Y", Localizacao_Ficheiro_Lingua)
    With Shape_Pais
        .left = X
        .top = Y
        .Visible = True
    End With
    
    
Exit Sub
Corrige_Erro:
Barra_Mapa.Picture = Form_Skin.Mapa_Mundo.Picture
Shape_Pais.Visible = False
End Sub


