VERSION 5.00
Begin VB.Form Form_Mensagem 
   Appearance      =   0  'Flat
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   0  'None
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
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
   Icon            =   "Form_Mensagem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Frame_Centro 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   449
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   465
      Width           =   6735
      Begin VB.PictureBox Barra_Text_Servidor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   360
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2160
         Visible         =   0   'False
         Width           =   5475
         Begin VB.TextBox Text_Servidor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   600
            TabIndex        =   5
            Top             =   30
            Width           =   1380
         End
         Begin VB.Shape Contorno_Servidor 
            BorderColor     =   &H00DFB000&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Pic_Remover 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   1440
         Picture         =   "Form_Mensagem.frx":000C
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1560
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CheckBox Check_Remover 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Caption         =   "Remover o ficheiro da biblioteca"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   0
         Top             =   1560
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label Label_Mensagem 
         AutoSize        =   -1  'True
         BackColor       =   &H00313131&
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagem"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1320
         TabIndex        =   13
         Top             =   480
         Width           =   915
      End
      Begin VB.Image Pic_Mensagem 
         Enabled         =   0   'False
         Height          =   720
         Left            =   360
         Top             =   240
         Width           =   720
      End
      Begin VB.Shape Shape_Centro 
         BorderColor     =   &H00212121&
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox Frame_Botoes 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFDFDF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   441
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4320
      Width           =   6615
      Begin VB.PictureBox Botao_Nao 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   4200
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   1740
         Begin VB.Label Label_Nao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Não"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   720
            TabIndex        =   11
            Top             =   45
            Width           =   330
         End
         Begin VB.Shape Contorno_Nao 
            BorderColor     =   &H00DFB000&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Botao_Sim 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   240
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1740
         Begin VB.Label Label_Sim 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sim"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   705
            TabIndex        =   10
            Top             =   45
            Width           =   330
         End
         Begin VB.Shape Contorno_Sim 
            BorderColor     =   &H00DFB000&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox Botao_Ok 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   2280
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
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
            TabIndex        =   9
            Top             =   45
            Width           =   240
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
      BackColor       =   &H002B2B2B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   6255
      Begin VB.Image Botao_Fechar 
         Height          =   195
         Left            =   5880
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   195
      End
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "NPlayer"
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
         TabIndex        =   7
         Top             =   120
         Width           =   765
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
Attribute VB_Name = "Form_Mensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Center programs Nikyts
'   Copyright © 2011-2012 Nikyts software ™ - Informática e tecnologia
'   www.nikyts.com / nikyts@hotmail.com
'   Desenvolvido por: Nelson do Carmo
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Variáveis para poder mover o formulário
'Dim bMoveFrom As Boolean, LastPoint As POINTAPI

Private Sub Botao_Fechar_Click()
    'Fechar o formulário
    Unload Me
End Sub

Private Sub Botao_Nao_Click()
    'Fechar o formulário
    If Check_Remover.Value = 0 Then Remover_da_Biblioteca = False
    Resposta = "No"
    Unload Me
End Sub

Private Sub Botao_Nao_GotFocus()
    'Colocar o focus no botao
    Contorno_Nao.Visible = True
End Sub

Private Sub Botao_Nao_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Nao_Click
    If KeyCode = vbKeyLeft Then Botao_nao_LostFocus: Botao_Sim_GotFocus: Botao_Sim.SetFocus
End Sub

Private Sub Botao_nao_LostFocus()
    'Remover o focus no botao
    Contorno_Nao.Visible = False
End Sub

Private Sub Botao_Ok_Click()
    'Fechar o formulário
    Unload Me
End Sub

Private Sub Botao_Ok_GotFocus()
    'Colocar o focus no botao
    Contorno_Ok.Visible = True
End Sub

Private Sub Botao_Ok_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Ok_Click
End Sub

Private Sub Botao_Ok_LostFocus()
    'Remover o focus no botao
    Contorno_Ok.Visible = False
End Sub

Private Sub Botao_Sim_Click()
    'Fechar o formulário
    If Check_Remover.Value = 1 Then Remover_da_Biblioteca = True
    Resposta = "Yes"
    Unload Me
End Sub

Private Sub Botao_Sim_GotFocus()
    'Colocar o focus no botao
    Contorno_Sim.Visible = True
End Sub

Private Sub Botao_Sim_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Sim_Click
    If KeyCode = vbKeyRight Then Botao_Sim_LostFocus: Botao_Nao_GotFocus: Botao_Nao.SetFocus
End Sub

Private Sub Botao_Sim_LostFocus()
    'Remover o focus no botao
    Contorno_Sim.Visible = False
End Sub

Private Sub Check_Remover_Click()
    'Des/Activar a opcção
    If Check_Remover.Value = 1 Then
        Pic_Remover.Picture = Form_Skin.Check_Over.Picture
    Else
        Pic_Remover.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Form_Activate()
    'Arredondar os cantos do formulário
    Arredondar_Cantos_do_Form Me, True
End Sub

Private Sub Form_Load()
    'Chamar o procedimento
    Carregar_Idioma
    Desenhar_Formulario
    Carregar_Skin
    
    Label_Titulo.Caption = App.ProductName
    
    'Propriedades iniciais do formulário
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Public Sub Carregar_Skin()
    'Procedimento para carregar o skin escolhido
    With Form_Skin
        Me.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Shape_Contorno.BorderColor = .Cor_Form_BorderColor.backcolor
        Frame_Centro.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Shape_Centro.BorderColor = .Cor_Contorno_Frame_Centro.backcolor
        Fundo_Barra_ControlBox.Picture = .Fundo_Barra_ControlBox.Picture
        Label_Titulo.ForeColor = .Cor_Label_Barra_Titulo.backcolor
        Botao_Fechar.Picture = .Botao_Fechar.Picture
        Label_Mensagem.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Fundo_Frame_Botoes.Picture = .Fundo_Frame_Botoes.Picture
        Label_Ok.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Ok.Picture = .Pic_Button.Picture
        Contorno_Ok.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Sim.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Sim.Picture = .Pic_Button.Picture
        Contorno_Sim.BorderColor = .Cor_Contorno_Caixas.backcolor
        Label_Nao.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Nao.Picture = .Pic_Button.Picture
        Contorno_Nao.BorderColor = .Cor_Contorno_Caixas.backcolor
        Check_Remover.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Check_Remover.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        Pic_Remover.Picture = .Check_Normal.Picture
        Pic_Remover.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Barra_Text_Servidor.PaintPicture Form_Skin.Pic_TextBox.Picture, 0, 0, 10, 26, 0, 0, 10, 26
        Barra_Text_Servidor.PaintPicture Form_Skin.Pic_TextBox.Picture, 10, 0, Barra_Text_Servidor.ScaleWidth, 26, 10, 0, 40, 26
        Barra_Text_Servidor.PaintPicture Form_Skin.Pic_TextBox.Picture, (Barra_Text_Servidor.ScaleWidth - 10), 0, 10, 26, 51, 0, 10, 26
        Text_Servidor.backcolor = .Cor_Fundo_Textbox.backcolor
        Text_Servidor.ForeColor = .Cor_Letra_Textbox.backcolor
        Contorno_Servidor.BorderColor = .Cor_Contorno_Caixas.backcolor
    End With
End Sub

Private Sub Form_Resize()
    'Chamar o procedimento
    Desenhar_Formulario
End Sub

Private Sub Label_Nao_Click()
    'Atalho para
    Botao_Nao_Click
End Sub

Private Sub Label_Ok_Click()
    'Atalho para
    Botao_Ok_Click
End Sub

Private Sub Label_Sim_Click()
    'Atalho para
    Botao_Sim_Click
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    With Me
        'Descrição dos 3 ultimos valores. Espaço entre: borda do form -> icon, icon -> label messagem, label_mensagem -> borda do form
        If Barra_Text_Servidor.Visible = False Then
            .Width = Screen.TwipsPerPixelX * (Label_Mensagem.Width + Pic_Mensagem.Width + 40 + 17 + 50)
        Else
            .Width = Screen.TwipsPerPixelX * (Form_Skin.Caixa_de_Texto.Width + Pic_Mensagem.Width + 40 + 17 + 50)
        End If
        .Height = Screen.TwipsPerPixelX * (Barra_ControlBox.ScaleHeight + Frame_Botoes.ScaleHeight + 36 + Label_Mensagem.Height + 30 + 10 + Check_Remover.Height + 20)
    End With
    
    Ajustar_Formulario Form_Mensagem, False, False, True, True
    
    Ajustar_Botao Form_Mensagem, Botao_Sim, Label_Sim, True, Contorno_Sim
    Ajustar_Botao Form_Mensagem, Botao_Ok, Label_Ok, True, Contorno_Ok
    Ajustar_Botao Form_Mensagem, Botao_Nao, Label_Nao, True, Contorno_Nao
    
    With Botao_Sim
        .left = (Frame_Botoes.ScaleWidth / 2) - .ScaleWidth - (.top / 2)
    End With
    With Botao_Ok
        .left = (Frame_Botoes.ScaleWidth - .ScaleWidth) / 2
    End With
    With Botao_Nao
        .left = (Frame_Botoes.ScaleWidth / 2) + (.top / 2)
    End With
    
    Ajustar_ChecBox Pic_Remover, Check_Remover
    
    With Pic_Mensagem
        .top = 20
        .left = 20
    End With
    
    With Label_Mensagem
        .top = Pic_Mensagem.top + 10
        .left = Pic_Mensagem.left + Pic_Mensagem.Width + 20
    End With
    
    With Check_Remover
        .top = Label_Mensagem.top + Label_Mensagem.Height + 20
        .left = Label_Mensagem.left
    End With
    
    With Pic_Remover
        .top = Check_Remover.top
        .left = Check_Remover.left
    End With
    
    Ajustar_Caixa_Texto Barra_Text_Servidor, Text_Servidor, Contorno_Servidor, False
    With Barra_Text_Servidor
        .top = Label_Mensagem.top + Label_Mensagem.Height + 5
        .left = Label_Mensagem.left
    End With
    
    'Ajustar os objectos depois de arredondar os cantos do formulário
    Shape_Contorno.left = 0
    Shape_Contorno.Width = Me.ScaleWidth - 1
    Frame_Botoes.Width = Frame_Botoes.ScaleWidth - 1
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Mensagem
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Mensagem
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Mensagem
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Mensagem
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Mensagem
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Mensagem
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Dim Text_Lingua As String: Text_Lingua = ReadINI("Settings", "Language", Localizacao_Ficheiro_Preferencias)
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Text_Lingua & ".lng"
    
    Botao_Fechar.ToolTipText = ReadINI("Message", "Button_Close", Localizacao_Ficheiro_Lingua)
    Label_Sim.Caption = ReadINI("Message", "Label_Yes", Localizacao_Ficheiro_Lingua)
    Label_Ok.Caption = ReadINI("Message", "Label_Ok", Localizacao_Ficheiro_Lingua)
    Label_Nao.Caption = ReadINI("Message", "Label_No", Localizacao_Ficheiro_Lingua)
    Check_Remover.Caption = ReadINI("Message", "Check_Remove", Localizacao_Ficheiro_Lingua)
End Sub

Private Sub Pic_Remover_Click()
    'Des/Activar a opcção
    If Check_Remover.Value = 0 Then
        Check_Remover.Value = 1
        Pic_Remover.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Remover.Value = 0
        Pic_Remover.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Text_Servidor_GotFocus()
    'Contorno da text box ao receber o focus
    Contorno_Servidor.Visible = True
End Sub

Private Sub Text_Servidor_LostFocus()
    'Contorno da text box ao perder o focus
    Contorno_Servidor.Visible = False
End Sub

