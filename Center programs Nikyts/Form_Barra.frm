VERSION 5.00
Begin VB.Form Form_Barra 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2415
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   161
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox Lista_Pastas 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      Height          =   1395
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Pic_Programa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00DFB000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   0
      Left            =   0
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "5555555555555"
      Top             =   0
      Width           =   960
   End
   Begin VB.Label Label_Programa 
      BackColor       =   &H00FF80FF&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   960
   End
   Begin VB.Image Image_Icon_Normal 
      Height          =   960
      Index           =   0
      Left            =   0
      Top             =   1440
      Width           =   960
   End
   Begin VB.Image Image_Icon_Over 
      Height          =   960
      Index           =   0
      Left            =   0
      Top             =   2520
      Width           =   960
   End
End
Attribute VB_Name = "Form_Barra"
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

'Variáveis do idioma
Dim Idioma_Erro As String
Dim Idioma_Descricao As String
Dim Idioma_Erro_Execucao As String
Dim Idioma_Conectar_Servidor As String
Dim Idioma_Internet_Desligada As String

Private Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos, construir o formulario
    With Me
        .top = 0
        .Height = Screen.TwipsPerPixelY * Form_Skin.Icon_Programa_Normal.Height
        .Width = Screen.TwipsPerPixelX * (Form_Skin.Icon_Programa_Normal.Width * Form_Principal.Image_Icon_Grande.Count)
        .left = (Screen.Width - .Width) / 2
    End With
    
    With Pic_Programa(0)
        .top = 0
        .Height = Form_Skin.Icon_Programa_Normal.Height
        .left = 0
        .Width = Form_Skin.Icon_Programa_Normal.Width
    End With
End Sub

Private Sub Form_Load()
    'Propriedades iniciais do programa
    Carregar_Idioma
    Desenhar_Formulario
    Verificar_Programas_Existentes (App.Path & "\Programs\")
    Carregar_Programas_Existentes
End Sub

Private Sub Form_Resize()
    'Chamar o procedimento
    Desenhar_Formulario
End Sub

Public Sub Verificar_Programas_Existentes(Path)
    'Procedimento para listar os programas existentes
    On Error Resume Next
    Dim Count, d(), I, DirName
    Dim directorio_da_pasta As String
    
    Lista_Pastas.Clear
    DirName = Dir(Path, 16)
    Do While DirName <> ""
        If DirName <> "." And DirName <> ".." Then
            If GetAttr(Path + DirName) = 16 Then
                If (Count Mod 10) = 0 Then
                    ReDim Preserve d(Count + 10)
                End If
                Count = Count + 1
                d(Count) = DirName
            End If
        End If
        DirName = Dir
    Loop
    For I = 1 To Count
        directorio_da_pasta = Path & d(I)
        Lista_Pastas.AddItem Dir(directorio_da_pasta, vbDirectory)
        'ListSubDirs Path & d(i) & "\"
    Next I
    DoEvents
End Sub

Public Sub Carregar_Programas_Existentes()
    'Procedimento para criar os icons dos programas instalados
    On Error Resume Next
    Dim icon_normal, icon_over As String
    
    Botao_Run.Enabled = False: Label_Run.Enabled = False
    Botao_Desinstalar.Enabled = False: Label_Desinstalar.Enabled = False
    
    If Lista_Pastas.ListCount <> 0 Then
        Lista_Pastas.ListIndex = 0
        'Logo normal
        icon_normal = App.Path & "\Programs\" & Lista_Pastas.List(0) & "\Options\Icon_Normal.jpg"
        icon_over = App.Path & "\Programs\" & Lista_Pastas.List(0) & "\Options\Icon_Over.jpg"
        
        If ArquivoExiste(icon_normal, False) And ArquivoExiste(icon_over, False) Then
            Pic_Programa(0).Picture = LoadPicture(icon_normal)
            Image_Icon_Normal(0).Picture = LoadPicture(icon_normal)
            Image_Icon_Over(0).Picture = LoadPicture(icon_over)
        Else
            Pic_Programa(0).Picture = Form_Skin.Icon_Programa_Normal.Picture
            Image_Icon_Normal(0).Picture = Form_Skin.Icon_Programa_Normal.Picture
            Image_Icon_Over(0).Picture = Form_Skin.Icon_Programa_Over.Picture
        End If
        
        Pic_Programa(0).Visible = True
        Label_Programa(0).Caption = Lista_Pastas.List(I)
        Label_Programa(0).Visible = True
        
        'Restantes...
        Dim Objecto As Integer: For Objecto = 1 To Lista_Pastas.ListCount - 1
            Load Pic_Programa(Objecto)
            Pic_Programa(Objecto).Picture = LoadPicture("")
            Pic_Programa(Objecto).Move Pic_Programa(Objecto - 1).left + Pic_Programa(Objecto - 1).Width, Pic_Programa(Objecto - 1).top
            
            icon_normal = App.Path & "\Programs\" & Lista_Pastas.List(Objecto) & "\Options\Icon_Normal.jpg"
            icon_over = App.Path & "\Programs\" & Lista_Pastas.List(Objecto) & "\Options\Icon_Over.jpg"
        
            Load Image_Icon_Normal(Objecto)
            Image_Icon_Normal(Objecto).Move Pic_Programa(Objecto).left, Image_Icon_Normal(0).top
            Image_Icon_Normal(Objecto).Visible = True
            
            Load Image_Icon_Over(Objecto)
            Image_Icon_Over(Objecto).Move Pic_Programa(Objecto).left, Image_Icon_Over(0).top
            Image_Icon_Over(Objecto).Visible = True
            
            If ArquivoExiste(icon_normal, False) And ArquivoExiste(icon_over, False) Then
                Pic_Programa(Objecto).Picture = LoadPicture(icon_normal)
                Image_Icon_Normal(Objecto).Picture = LoadPicture(icon_normal)
                Image_Icon_Over(Objecto).Picture = LoadPicture(icon_over)
            Else
                Pic_Programa(Objecto).Picture = Form_Skin.Icon_Programa_Normal.Picture
                Image_Icon_Normal(Objecto).Picture = Form_Skin.Icon_Programa_Normal.Picture
                Image_Icon_Over(Objecto).Picture = Form_Skin.Icon_Programa_Over.Picture
            End If
            Pic_Programa(Objecto).Visible = True
            
            Load Label_Programa(Objecto)
            Label_Programa(Objecto).Move Pic_Programa(Objecto).left, Label_Programa(Objecto - 1).top
            Label_Programa(Objecto).Caption = Lista_Pastas.List(Objecto)
            Label_Programa(Objecto).Visible = True
        Next
    End If
End Sub

Private Sub Pic_Programa_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Executar o programa automaticamente
    Pic_Programa(Index).Picture = Image_Icon_Normal(Index).Picture
    
    On Error GoTo Corrige_Erro
    Shell App.Path & "\Programs\" & Label_Programa(Index).Caption & "\" & Label_Programa(Index).Caption & ".exe"
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Carregar_Idioma
Select Case err.Number
    Case -2146697211
        Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada
        
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub Pic_Programa_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar o botão
    Dim Ret As Long
    'Aplicar imagem over
    If GetCapture() <> Pic_Programa(Index).hwnd Then
        Ret = SetCapture(Pic_Programa(Index).hwnd)
        Pic_Programa(Index).Picture = Image_Icon_Over(Index).Picture
    End If
    If X > 0 And X < Pic_Programa(Index).Width And Y > 0 And Y < Pic_Programa(Index).Height Then
        CurrentX = X
        CurrentY = Y
    Else
        'Aplicar imagem normal
        If GetCapture() = Pic_Programa(Index).hwnd Then
            Ret = ReleaseCapture()
            Pic_Programa(Index).Picture = Image_Icon_Normal(Index).Picture
        End If
    End If
    
    'Pic_Programa(Index).ToolTipText = Label_Programa(Index).Caption
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Dim Text_Lingua As String: Text_Lingua = ReadINI("Settings", "Language", Localizacao_Ficheiro_Preferencias)
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Text_Lingua & ".lng"
    
    Idioma_Erro = ReadINI("Main", "Label_Error", Localizacao_Ficheiro_Lingua)
    Idioma_Descricao = ReadINI("Main", "Label_Description", Localizacao_Ficheiro_Lingua)
    Idioma_Erro_Execucao = ReadINI("Main", "Error_Execution", Localizacao_Ficheiro_Lingua)
    Idioma_Conectar_Servidor = ReadINI("Main", "Error_Connect", Localizacao_Ficheiro_Lingua)
    Idioma_Internet_Desligada = ReadINI("Main", "Error_Internet", Localizacao_Ficheiro_Lingua)
End Sub
