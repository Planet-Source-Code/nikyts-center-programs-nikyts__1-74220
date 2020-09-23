VERSION 5.00
Begin VB.Form Form_Instalar 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   0  'None
   Caption         =   "Center programs Nikyts"
   ClientHeight    =   5445
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   6720
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
   Icon            =   "Form_Instalar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   448
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Frame_Centro 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   0
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   481
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   465
      Width           =   7215
      Begin VB.PictureBox Frame_Componentes 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   3075
         Left            =   240
         ScaleHeight     =   205
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   840
         Width           =   5475
         Begin VB.PictureBox Lista_Componentes 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1695
            Left            =   120
            ScaleHeight     =   113
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   225
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   120
            Width           =   3375
            Begin VB.Label Label_Componente 
               BackColor       =   &H00EEEEEE&
               BackStyle       =   0  'Transparent
               Caption         =   "Componente"
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   0
               Left            =   15
               TabIndex        =   14
               Top             =   0
               Width           =   2760
            End
            Begin VB.Label Shape_Sombra 
               BackColor       =   &H00DFB000&
               Height          =   240
               Index           =   0
               Left            =   0
               TabIndex        =   15
               Top             =   0
               Width           =   3975
            End
         End
         Begin VB.Shape Shape_Centro 
            BorderColor     =   &H00C0C0C0&
            Height          =   375
            Left            =   4800
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         ForeColor       =   &H00000000&
         Height          =   990
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         ForeColor       =   &H00000000&
         Height          =   1005
         Left            =   1800
         TabIndex        =   12
         Top             =   3120
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.Label Label_Informacao 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "O programa para ser executado correctamente poderá ter que registar algumas componentes."
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   6420
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   60
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
      ScaleWidth      =   417
      TabIndex        =   5
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
         Caption         =   "Verificar componentes"
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
         TabIndex        =   6
         Top             =   120
         Width           =   2205
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Top             =   0
         Width           =   285
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
      ScaleWidth      =   425
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4800
      Width           =   6375
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
         Begin VB.Label Label_Cancelar 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   465
            TabIndex        =   4
            Top             =   45
            Width           =   780
         End
         Begin VB.Shape Contorno_Cancelar 
            BorderColor     =   &H00DFB000&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
      End
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
         TabIndex        =   1
         Top             =   120
         Width           =   1740
         Begin VB.Label Label_Ok 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Registar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   510
            TabIndex        =   2
            Top             =   45
            Width           =   720
         End
         Begin VB.Shape Contorno_Ok 
            BorderColor     =   &H00DFB000&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.Image Fundo_Frame_Botoes 
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   315
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
Attribute VB_Name = "Form_Instalar"
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

'API's para registar as ocx utilizadas pelo programa
Private Const STATUS_WAIT_0 = &H0
Private Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)

Private Declare Sub ExitThread Lib "kernel32.dll" (ByVal lngExitCode As Long)
Private Declare Function LoadLibraryRegister Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibraryRegister Lib "kernel32.dll" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetProcAddressRegister Lib "kernel32.dll" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThreadForRegister Lib "kernel32.dll" Alias "CreateThread" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpparameter As Long, ByVal dwCreationFlags As Long, lngThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32.dll" (ByVal lngThread As Long, lpExitCode As Long) As Long

Public Enum EDllRegister
  DllRegisterServer = 1
  DllUnRegisterServer = 2
End Enum

Public Enum EDllReturnConstants
  edrcNone = 0
  edrcErrorOnLoadInMemory = 1
  edrcInvalidActivex = 2
  edrcFailedRegistration = 3
  edrcRegistered = 4
  edrcUnregistered = 5
End Enum

Private mReturn As EDllReturnConstants

'Variável para identificar as pastas do windows
Dim Pasta_Sistema As String

'Variável para indicar qual a linha que está selecionada da lista linguas
Dim Linha_Selecionada As Integer

'Variáveis do idioma
Dim Idioma_Erro As String
Dim Idioma_Descricao As String
Dim Idioma_Erro_Execucao As String
Dim Idioma_Registar As String
Dim Idioma_Sugestao As String

Public Sub Botao_Cancelar_Click()
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
    'Fechar o formulário
    End
End Sub

Public Sub Botao_Ok_Click()
    'atalho para
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
    'Remover o focus no botao
    Contorno_Ok.Visible = False
End Sub

Public Sub Carregar_Skin()
    'Procedimento para carregar o skin escolhido
    With Form_Skin
        Me.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Shape_Contorno.BorderColor = .Cor_Form_BorderColor.backcolor
        Frame_Centro.backcolor = .Cor_do_Fundo_dos_Formularios.backcolor
        Fundo_Barra_ControlBox.Picture = .Fundo_Barra_ControlBox.Picture
        Label_Titulo.ForeColor = .Cor_Label_Barra_Titulo.backcolor
        Botao_Fechar.Picture = .Botao_Fechar.Picture
        Label_Informacao.ForeColor = .Cor_Letra_Label_Formulario.backcolor
        
        Lista_Componentes.backcolor = .Cor_Fundo_Textbox.backcolor
        Shape_Sombra(0).backcolor = .Cor_Contorno_Caixas.backcolor
        Label_Componente(0).ForeColor = .Cor_Letra_Textbox.backcolor
        Fundo_Frame_Botoes.Picture = .Fundo_Frame_Botoes.Picture
        Label_Ok.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Label_Cancelar.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Botao_Ok.Picture = .Pic_Button.Picture
        Botao_Cancelar.Picture = .Pic_Button.Picture
        Contorno_Ok.BorderColor = .Cor_Contorno_Caixas.backcolor
        Contorno_Cancelar.BorderColor = .Cor_Contorno_Caixas.backcolor
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Unload Me: End
End Sub

Private Sub Form_Load()
    On Error GoTo Corrige_Erro
    Carregar_Idioma
    Desenhar_Formulario
    Carregar_Skin
    
    'Variáveis para poder mover o formulário
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    Arredondar_Cantos_do_Form Me, True
    
    'Pasta dos componentes
    Dir1.Path = App.Path & "\Components\"
    File1.Pattern = "*.ocx;*.dll;*.tlb" 'Filtar os ficheiros
    File1.Path = Dir1.Path

    'Criar a lista consoante o nº de idiomas disponiveis
    If File1.ListCount <> 0 Then
        Label_Componente(0).Caption = ""
        Label_Componente(0).Visible = True
        Dim Objecto As Integer
        For Objecto = 1 To File1.ListCount - 1
            Load Label_Componente(Objecto)
            Label_Componente(Objecto).Move Label_Componente(Objecto - 1).left, Label_Componente(Objecto - 1).top + Label_Componente(Objecto - 1).Height
            Label_Componente(Objecto).Visible = True

            Load Shape_Sombra(Objecto)
            Shape_Sombra(Objecto).Move Shape_Sombra(Objecto - 1).left, Shape_Sombra(Objecto - 1).top + Shape_Sombra(Objecto - 1).Height
            Shape_Sombra(Objecto).Visible = False
            Shape_Sombra(Objecto).ZOrder 1
        Next Objecto

        'Preencher as label's com as linguas disponiveis
        Dim Z As Integer
        File1.ListIndex = 0
        For Z = 0 To File1.ListCount - 1
            Label_Componente(Z).Caption = File1.List(Z)
        Next Z
    End If
    'Selecionar a 1ªlinha da lista linguas
    Linha_Selecionada = 0
    Shape_Sombra(0).Visible = True
    Label_Componente(0).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor

    'Verificar qual é a pasta do system32
    Pasta_Sistema = Environ$("windir") & IIf(Len(Environ$("OS")), "\SYSTEM32", "\SYSTEM")
    Verificar_Pastas
    If File1.ListCount <> 0 Then
        Verificar_Componentes
    Else
        Form_Principal.Show
        Unload Me
    End If

Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case Else
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub Form_Resize()
    'Chamar o procedimento
    Desenhar_Formulario
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Instalar
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Instalar
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Instalar
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Instalar
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Mover_Formulario Form_Instalar
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Instalar
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    Frame_Componentes.Width = Form_Skin.Frame_Componentes.Width
    With Me
        .Width = Screen.TwipsPerPixelX * ((Frame_Componentes.ScaleWidth) + (2 * Label_Informacao.left) + 20)
        '.Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Fundo_Frame_Botoes.Height + Frame_Centro.ScaleHeight)
        .Height = Screen.TwipsPerPixelX * (Fundo_Barra_ControlBox.Height + Label_Informacao.left + Label_Informacao.Height + 3 _
                + Form_Skin.Frame_Componentes.Height + Fundo_Frame_Botoes.Height + Label_Informacao.left)
    End With
    
    Ajustar_Formulario Me, False, False, True, True
    
    Ajustar_Botao Me, Botao_Cancelar, Label_Cancelar, True, Contorno_Cancelar
    Ajustar_Botao Me, Botao_Ok, Label_Ok, True, Contorno_Ok
    
    With Botao_Cancelar
        .left = Frame_Botoes.ScaleWidth - .ScaleWidth - .top
    End With
    With Botao_Ok
        .left = Botao_Cancelar.left - .ScaleWidth - .top
    End With
    
    With Label_Informacao
        .top = .left
        .Width = Form_Skin.Frame_Componentes.Width
    End With
    
    With Frame_Componentes
        .top = Label_Informacao.top + Label_Informacao.Height + 3
        .Height = Form_Skin.Frame_Componentes.Height
        .Width = Form_Skin.Frame_Componentes.Width
        .left = Label_Informacao.left
    End With
    
    With Shape_Centro
        .Height = Frame_Componentes.ScaleHeight
        .top = 0
        .Width = Frame_Componentes.ScaleWidth
        .left = 0
    End With
    
    With Lista_Componentes
        .Height = Frame_Componentes.ScaleHeight - 6
        .top = 3
        .Width = Frame_Componentes.ScaleWidth - 6
        .left = 3
    End With
    
    With Shape_Sombra(0)
        .Width = Lista_Componentes.ScaleWidth
    End With
    
    With Label_Componente(0)
        .Width = Lista_Componentes.ScaleWidth
    End With
    
    With Shape_Sombra(0)
        .Width = Lista_Componentes.Width
        .left = 0
    End With
    
    'Ajustar os objectos depois de arredondar os cantos do formulário
    Shape_Contorno.left = 0
    Shape_Contorno.Width = Me.ScaleWidth - 1
    Frame_Botoes.Width = Frame_Botoes.ScaleWidth - 1
End Sub

Private Sub Label_Cancelar_Click()
    'Atalho para
    Unload Me
    End
End Sub

Private Sub Label_Ok_Click()
    'Percorrer a lista das compoennets do programa
    'On Error GoTo Corrige_Erro
    Me.MousePointer = 11
    Botao_Ok.Enabled = False
    Label_Ok.Enabled = False
    Contorno_Ok.Visible = False
    
    Dim Linha As Integer
    For Linha = 0 To File1.ListCount - 1
        File1.ListIndex = Linha
        
        'Selecionar a linha que está a ser verificada
        Shape_Sombra(Linha_Selecionada).Visible = False
        Label_Componente(Linha_Selecionada).ForeColor = Form_Skin.Cor_Letra_Textbox.backcolor
        Shape_Sombra(File1.ListIndex).Visible = True
        Label_Componente(File1.ListIndex).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
        Linha_Selecionada = File1.ListIndex
        
        'Verificar se a componente já existe na pasta do windows, caso não exista copia a componente para a pasta e regista
        If Dir(Pasta_Sistema & "\" & File1.List(Linha), vbDirectory) = Empty Then
            'Copia a componente para a pasta do windows
            FileCopy App.Path & "\Components\" & File1.List(Linha), Pasta_Sistema & "\" & File1.List(Linha)
            'Regista a componente
            Registar_Ocx_Componente Pasta_Sistema & "\" & File1.List(Linha)
        End If
    Next Linha
    
    'Ao terminar o registo das componentes
    Me.MousePointer = 0
    Form_Principal.Show
    Unload Me
    
Exit Sub
Corrige_Erro:
Mensagem_de_Aviso "Error", Idioma_Registar & vbNewLine & Idioma_Sugestao
End
End Sub

Public Function Registar_Componente(ByVal FileName As String, ByVal RegFunction As EDllRegister) As EDllReturnConstants
    'Função para registar as componentes utilizadas pelo programa
    Dim lngLib As Long
    Dim lngProcAddress As Long
    Dim lngThreadID As Long
    Dim lngSucess As Long
    Dim lngExitCode As Long
    Dim lngThread As Long
    
    'Por padrao passa para nada o retorno...
    Registar_Componente = edrcNone
    
    If FileName = "" Then Exit Function
    
    lngLib = LoadLibraryRegister(FileName)
    If lngLib = 0 Then
       Registar_Componente = edrcErrorOnLoadInMemory
       Exit Function
    End If
    
    Select Case RegFunction
    Case EDllRegister.DllRegisterServer
        lngProcAddress = GetProcAddressRegister(lngLib, "DllRegisterServer")
    Case EDllRegister.DllUnRegisterServer
        lngProcAddress = GetProcAddressRegister(lngLib, "DllUnregisterServer")
    Case Else
    End Select
    
    If lngProcAddress = 0 Then
       Registar_Componente = edrcInvalidActivex
       If lngLib Then Call FreeLibraryRegister(lngLib)
       Exit Function
    Else
       lngThread = CreateThreadForRegister(ByVal 0&, 0&, ByVal lngProcAddress, ByVal 0&, 0&, lngThreadID)
       If lngThread Then
            lngSucess = (WaitForSingleObject(lngThread, 10000) = WAIT_OBJECT_0)
            If Not lngSucess Then
               Call GetExitCodeThread(lngThread, lngExitCode)
               Call ExitThread(lngExitCode)
               Registar_Componente = edrcFailedRegistration
               If lngLib Then Call FreeLibraryRegister(lngLib)
               Exit Function
            Else
                If RegFunction = DllRegisterServer Then
                    Registar_Componente = edrcRegistered
                ElseIf RegFunction = DllUnRegisterServer Then
                    Registar_Componente = edrcUnregistered
                End If
            End If
            Call CloseHandle(lngThread)
            If lngLib Then Call FreeLibraryRegister(lngLib)
       End If
    End If
End Function

Public Sub Registar_Ocx_Componente(Directorio_da_Componente As String)
    'Procedimento para registar as ocx do programa
    On Error Resume Next
    'Proceder ao registro das componentes
    mReturn = Registar_Componente(Trim(Directorio_da_Componente), DllRegisterServer)
    Select Case mReturn
        Case edrcErrorOnLoadInMemory
            'Label_Estado.Caption = "O componente " & File1.List(Linha) & " não foi carregado na memória"
        Case edrcInvalidActivex
            'Label_Estado.Caption = "O componente " & File1.List(Linha) & " é inválido"
        Case edrcFailedRegistration
            'Label_Estado.Caption = "Ocorreu um erro ao registar o componente " & File1.List(Linha)
        Case edrcRegistered
            'Label_Estado.Caption = "O componente " & File1.List(Linha) & " foi registado com sucesso"
    End Select
End Sub

Public Sub Verificar_Componentes()
    'Procedimento para verificar se existe alguma componente que não esteja registada
    'DoEvents
    Dim Linha As Integer
    For Linha = 0 To File1.ListCount - 1
        File1.ListIndex = Linha
        
        'Selecionar a linha que está a ser verificada
        Shape_Sombra(Linha_Selecionada).Visible = False
        Label_Componente(Linha_Selecionada).ForeColor = Form_Skin.Cor_Letra_Textbox.backcolor
        Shape_Sombra(File1.ListIndex).Visible = True
        Label_Componente(File1.ListIndex).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
        Linha_Selecionada = File1.ListIndex
    
        'Verificar se a componente já existe na pasta do windows, caso não exista mostra o formulário install
        If Dir(Pasta_Sistema & "\" & File1.List(Linha), vbDirectory) = Empty Then
            'MsgBox (Pasta_Sistema & "\" & File1.List(Linha) & " Não existe")
            Me.Visible = True
            Label_Titulo.Caption = ReadINI("Install", "Title2", Localizacao_Ficheiro_Lingua)
            Botao_Ok.Enabled = True
            Label_Ok.Enabled = True
            Me.MousePointer = 0
            Botao_Cancelar_LostFocus
            Botao_Ok_GotFocus
            Botao_Ok.SetFocus
            Exit Sub
        
        Else
            'MsgBox (Pasta_Sistema & "\" & File1.List(Linha) & " existe")
            If Linha = File1.ListCount - 1 Then
                'MsgBox ("Concluido")
                Me.Visible = False
                Form_Principal.Show
                Unload Me
            End If
        End If
    Next Linha
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
    Dim Text_Lingua As String: Text_Lingua = ReadINI("Settings", "Language", Localizacao_Ficheiro_Preferencias)
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Text_Lingua & ".lng"
    
    Label_Titulo.Caption = ReadINI("Install", "Title1", Localizacao_Ficheiro_Lingua)
    Botao_Fechar.ToolTipText = ReadINI("Install", "Button_Close", Localizacao_Ficheiro_Lingua)
    Label_Informacao.Caption = ReadINI("Install", "Label_Info", Localizacao_Ficheiro_Lingua)
    Label_Ok.Caption = ReadINI("Install", "Button_Ok", Localizacao_Ficheiro_Lingua)
    Label_Cancelar.Caption = ReadINI("Install", "Button_Cancel", Localizacao_Ficheiro_Lingua)
    
    Idioma_Erro = ReadINI("Message", "Label_Error", Localizacao_Ficheiro_Lingua)
    Idioma_Descricao = ReadINI("Message", "Label_Description", Localizacao_Ficheiro_Lingua)
    Idioma_Erro_Execucao = ReadINI("Message", "Error_Execution", Localizacao_Ficheiro_Lingua)
    Idioma_Registar = ReadINI("Message", "Error_Register", Localizacao_Ficheiro_Lingua)
    Idioma_Sugestao = ReadINI("Message", "Error_Suggestion", Localizacao_Ficheiro_Lingua)
End Sub
