VERSION 5.00
Begin VB.Form Form_Skin 
   Appearance      =   0  'Flat
   BackColor       =   &H0080FF80&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12780
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   17670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   852
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1178
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Pic_TextBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   11880
      Picture         =   "Form_Skin.frx":0000
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   720
      Width           =   1215
   End
   Begin VB.Image Icon_Programa_Over 
      Height          =   960
      Left            =   9480
      Picture         =   "Form_Skin.frx":12F2
      Top             =   600
      Width           =   960
   End
   Begin VB.Image Icon_Programa_Normal 
      Height          =   960
      Left            =   8400
      Picture         =   "Form_Skin.frx":4334
      Top             =   600
      Width           =   960
   End
   Begin VB.Image Icon_Grande_Over 
      Height          =   1800
      Left            =   8400
      Picture         =   "Form_Skin.frx":7376
      Top             =   3600
      Width           =   1800
   End
   Begin VB.Image Imagem_Regiao 
      Height          =   3000
      Index           =   9
      Left            =   3480
      Picture         =   "Form_Skin.frx":11C78
      Top             =   13080
      Width           =   7500
   End
   Begin VB.Image Imagem_Regiao 
      Height          =   3000
      Index           =   8
      Left            =   3120
      Picture         =   "Form_Skin.frx":19C35
      Top             =   12840
      Width           =   7500
   End
   Begin VB.Image Imagem_Regiao 
      Height          =   3000
      Index           =   7
      Left            =   2760
      Picture         =   "Form_Skin.frx":219FE
      Top             =   12480
      Width           =   7500
   End
   Begin VB.Image Imagem_Regiao 
      Height          =   3000
      Index           =   6
      Left            =   2400
      Picture         =   "Form_Skin.frx":29B0D
      Top             =   12240
      Width           =   7500
   End
   Begin VB.Image Imagem_Regiao 
      Height          =   3000
      Index           =   4
      Left            =   1680
      Picture         =   "Form_Skin.frx":31AEC
      Top             =   11760
      Width           =   7500
   End
   Begin VB.Image Imagem_Regiao 
      Height          =   3000
      Index           =   3
      Left            =   1320
      Picture         =   "Form_Skin.frx":39B38
      Top             =   11520
      Width           =   7500
   End
   Begin VB.Image Imagem_Regiao 
      Height          =   3000
      Index           =   2
      Left            =   960
      Picture         =   "Form_Skin.frx":41B49
      Top             =   11280
      Width           =   7500
   End
   Begin VB.Image Imagem_Regiao 
      Height          =   3000
      Index           =   1
      Left            =   600
      Picture         =   "Form_Skin.frx":499D1
      Top             =   11040
      Width           =   7500
   End
   Begin VB.Image Imagem_Regiao 
      Height          =   3000
      Index           =   0
      Left            =   240
      Picture         =   "Form_Skin.frx":51C41
      Top             =   10800
      Width           =   7500
   End
   Begin VB.Image Imagem_Regiao 
      Height          =   3000
      Index           =   5
      Left            =   2040
      Picture         =   "Form_Skin.frx":59EA5
      Top             =   12000
      Width           =   7500
   End
   Begin VB.Image Icon_Pequeno 
      Height          =   600
      Left            =   10320
      Picture         =   "Form_Skin.frx":61DFA
      Top             =   1680
      Width           =   600
   End
   Begin VB.Image Icon_Grande_Normal 
      Height          =   1800
      Left            =   8400
      Picture         =   "Form_Skin.frx":630FC
      Top             =   1680
      Width           =   1800
   End
   Begin VB.Image Icon_Sobre_Down 
      Height          =   1050
      Left            =   7080
      Picture         =   "Form_Skin.frx":6D9FE
      Top             =   6000
      Width           =   1200
   End
   Begin VB.Image Icon_Sobre_Normal 
      Height          =   1050
      Left            =   5760
      Picture         =   "Form_Skin.frx":71BE0
      Top             =   6000
      Width           =   1200
   End
   Begin VB.Image Icon_Menu_Down 
      Height          =   210
      Left            =   14280
      Picture         =   "Form_Skin.frx":75DC2
      Top             =   4440
      Width           =   210
   End
   Begin VB.Image Icon_Menu_Normal 
      Height          =   210
      Left            =   14040
      Picture         =   "Form_Skin.frx":7606C
      Top             =   4440
      Width           =   210
   End
   Begin VB.Image Button_Menu_Normal 
      Height          =   300
      Left            =   14040
      Picture         =   "Form_Skin.frx":76316
      Top             =   4680
      Width           =   2520
   End
   Begin VB.Image Button_Menu_Down 
      Height          =   300
      Left            =   14040
      Picture         =   "Form_Skin.frx":78AB8
      Top             =   5040
      Width           =   2520
   End
   Begin VB.Image Image_Estrelas_5 
      Height          =   465
      Left            =   10920
      Picture         =   "Form_Skin.frx":7B25A
      Top             =   6480
      Width           =   2370
   End
   Begin VB.Image Image_Estrelas_4 
      Height          =   465
      Left            =   10920
      Picture         =   "Form_Skin.frx":7EC40
      Top             =   6000
      Width           =   2370
   End
   Begin VB.Image Image_Estrelas_3 
      Height          =   465
      Left            =   10920
      Picture         =   "Form_Skin.frx":82626
      Top             =   5520
      Width           =   2370
   End
   Begin VB.Image Image_Estrelas_2 
      Height          =   465
      Left            =   10920
      Picture         =   "Form_Skin.frx":8600C
      Top             =   5040
      Width           =   2370
   End
   Begin VB.Image Image_Estrelas_1 
      Height          =   465
      Left            =   10920
      Picture         =   "Form_Skin.frx":899F2
      Top             =   4560
      Width           =   2370
   End
   Begin VB.Image Image_Estrelas_0 
      Height          =   465
      Left            =   10920
      Picture         =   "Form_Skin.frx":8D3D8
      Top             =   4080
      Width           =   2370
   End
   Begin VB.Image Botao_Pesquisar 
      Height          =   195
      Left            =   9720
      Picture         =   "Form_Skin.frx":90DBE
      ToolTipText     =   "Pesquisar"
      Top             =   5520
      Width           =   195
   End
   Begin VB.Image Check_Over 
      Height          =   195
      Left            =   9120
      Picture         =   "Form_Skin.frx":910EE
      Top             =   5580
      Width           =   195
   End
   Begin VB.Image Check_Normal 
      Height          =   195
      Left            =   8880
      Picture         =   "Form_Skin.frx":9143A
      Top             =   5580
      Width           =   195
   End
   Begin VB.Image Seta_Combo 
      Height          =   315
      Left            =   9360
      Picture         =   "Form_Skin.frx":9175E
      Top             =   5520
      Width           =   285
   End
   Begin VB.Image Opcao_Normal 
      Height          =   180
      Left            =   8400
      Picture         =   "Form_Skin.frx":91A7E
      Top             =   5580
      Width           =   180
   End
   Begin VB.Image Opcao_Over 
      Height          =   180
      Left            =   8640
      Picture         =   "Form_Skin.frx":91DA9
      Top             =   5580
      Width           =   180
   End
   Begin VB.Image Icon_Invitation 
      Height          =   720
      Left            =   11040
      Picture         =   "Form_Skin.frx":9210E
      Top             =   7560
      Width           =   720
   End
   Begin VB.Image Icon_Link 
      Height          =   720
      Left            =   10320
      Picture         =   "Form_Skin.frx":92B19
      Top             =   7560
      Width           =   720
   End
   Begin VB.Image Caixa_de_Observacoes 
      Height          =   1905
      Left            =   8160
      Top             =   10200
      Width           =   5475
   End
   Begin VB.Image Caixa_de_Texto_Mini 
      Height          =   390
      Left            =   5400
      Top             =   10800
      Width           =   915
   End
   Begin VB.Image TextBox_Intermediate 
      Height          =   390
      Left            =   5400
      Top             =   10320
      Width           =   2535
   End
   Begin VB.Image Caixa_de_Texto 
      Height          =   390
      Left            =   7920
      Top             =   9600
      Width           =   5475
   End
   Begin VB.Image Icon_Pesquisar_Down 
      Height          =   1050
      Left            =   7080
      Picture         =   "Form_Skin.frx":9465B
      Top             =   600
      Width           =   1200
   End
   Begin VB.Image Icon_Instalados_Down 
      Height          =   1050
      Left            =   7080
      Picture         =   "Form_Skin.frx":9883D
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Image Icon_Partilhar_Down 
      Height          =   1050
      Left            =   7080
      Picture         =   "Form_Skin.frx":9CA1F
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Image Icon_Opcoes_Down 
      Height          =   1050
      Left            =   7080
      Picture         =   "Form_Skin.frx":A0C01
      Top             =   3840
      Width           =   1200
   End
   Begin VB.Image Icon_Suporte_Down 
      Height          =   1050
      Left            =   7080
      Picture         =   "Form_Skin.frx":A4DE3
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Image Icon_Suporte_Normal 
      Height          =   1050
      Left            =   5760
      Picture         =   "Form_Skin.frx":A8FC5
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Image Icon_Opcoes_Normal 
      Height          =   1050
      Left            =   5760
      Picture         =   "Form_Skin.frx":AD1A7
      Top             =   3840
      Width           =   1200
   End
   Begin VB.Image Icon_Partilhar_Normal 
      Height          =   1050
      Left            =   5760
      Picture         =   "Form_Skin.frx":B1389
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Image Icon_Instalados_Normal 
      Height          =   1050
      Left            =   5760
      Picture         =   "Form_Skin.frx":B556B
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Image Icon_Pesquisar_Normal 
      Height          =   1050
      Left            =   5760
      Picture         =   "Form_Skin.frx":B974D
      Top             =   600
      Width           =   1200
   End
   Begin VB.Image Image_Background 
      Height          =   1290
      Left            =   13080
      Picture         =   "Form_Skin.frx":BD92F
      Top             =   2400
      Width           =   1290
   End
   Begin VB.Image Pic_Button 
      Height          =   390
      Left            =   10920
      Picture         =   "Form_Skin.frx":C30C9
      Top             =   3480
      Width           =   1440
   End
   Begin VB.Image Botao_Form 
      Height          =   390
      Left            =   15840
      Top             =   840
      Width           =   1440
   End
   Begin VB.Image Fundo_Frame_Botoes 
      Height          =   615
      Left            =   12600
      Picture         =   "Form_Skin.frx":C3637
      Top             =   2880
      Width           =   405
   End
   Begin VB.Image Fundo_Barra_ControlBox 
      Height          =   360
      Left            =   8400
      Picture         =   "Form_Skin.frx":C393F
      Top             =   120
      Width           =   285
   End
   Begin VB.Image Botao_Tray_Normal 
      Height          =   195
      Left            =   8040
      Picture         =   "Form_Skin.frx":C3BE8
      ToolTipText     =   "Colocar o ícone na bandeja"
      Top             =   240
      Width           =   180
   End
   Begin VB.Image Botao_Maximizar_Normal 
      Height          =   195
      Left            =   7320
      Picture         =   "Form_Skin.frx":C3EEE
      ToolTipText     =   "Maximizar"
      Top             =   240
      Width           =   180
   End
   Begin VB.Image Botao_Restaurar_Normal 
      Height          =   195
      Left            =   7560
      Picture         =   "Form_Skin.frx":C41F4
      ToolTipText     =   "Restaurar"
      Top             =   240
      Width           =   180
   End
   Begin VB.Image Botao_Minimizar_Normal 
      Height          =   195
      Left            =   7800
      Picture         =   "Form_Skin.frx":C44FA
      ToolTipText     =   "Minimizar"
      Top             =   240
      Width           =   180
   End
   Begin VB.Image Botao_Fechar 
      Height          =   195
      Left            =   7080
      Picture         =   "Form_Skin.frx":C4800
      Top             =   240
      Width           =   180
   End
   Begin VB.Image Frame_Componentes 
      Height          =   3060
      Left            =   11280
      Top             =   8760
      Width           =   5475
   End
   Begin VB.Image Mapa_Mundo 
      Height          =   3000
      Left            =   0
      Picture         =   "Form_Skin.frx":C4B06
      Top             =   10560
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Image Icon_Error 
      Enabled         =   0   'False
      Height          =   720
      Left            =   8160
      Picture         =   "Form_Skin.frx":CCD29
      Top             =   7560
      Width           =   720
   End
   Begin VB.Image Icon_Quest 
      Enabled         =   0   'False
      Height          =   720
      Left            =   8880
      Picture         =   "Form_Skin.frx":CDB4D
      Top             =   7560
      Width           =   720
   End
   Begin VB.Image Icon_Info 
      Enabled         =   0   'False
      Height          =   720
      Left            =   9600
      Picture         =   "Form_Skin.frx":CE9A7
      Top             =   7560
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Label_Contador_Popup"
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   92
      Top             =   5430
      Width           =   2400
   End
   Begin VB.Label Cor_Label_Contador_Popup 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   91
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Letra caixas de texto"
      Height          =   195
      Index           =   13
      Left            =   480
      TabIndex        =   90
      Top             =   4710
      Width           =   1800
   End
   Begin VB.Label Cor_Letra_Textbox 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   89
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Cor_Grid_ForeColorFixed 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   88
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_ForeColorFixed"
      Height          =   195
      Index           =   3
      Left            =   3600
      TabIndex        =   87
      Top             =   4350
      Width           =   2175
   End
   Begin VB.Label Cor_Grid_ForeColorSel 
      BackColor       =   &H00E6E6E6&
      Height          =   255
      Left            =   3240
      TabIndex        =   86
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_ForeColorSel"
      Height          =   195
      Left            =   3600
      TabIndex        =   85
      Top             =   4710
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_BackColorSel"
      Height          =   195
      Index           =   1
      Left            =   3600
      TabIndex        =   84
      Top             =   3630
      Width           =   2040
   End
   Begin VB.Label Cor_Grid_BackColorSel 
      BackColor       =   &H00DFB000&
      Height          =   255
      Left            =   3240
      TabIndex        =   83
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_BackColor"
      Height          =   195
      Index           =   1
      Left            =   3600
      TabIndex        =   82
      Top             =   2550
      Width           =   1770
   End
   Begin VB.Label Cor_Grid_BackColor 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   81
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Cor_Grid_ForeColor 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   80
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_ForeColor"
      Height          =   195
      Index           =   1
      Left            =   3600
      TabIndex        =   79
      Top             =   3990
      Width           =   1725
   End
   Begin VB.Label Cor_Grid_Color 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   78
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Cor_do_Fundo_dos_Formularios 
      BackColor       =   &H00313131&
      Height          =   255
      Left            =   120
      TabIndex        =   77
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fundo dos formulários"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   76
      Top             =   2910
      Width           =   1905
   End
   Begin VB.Label Cor_da_Letra_do_Botao 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   75
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Letra do botao"
      Height          =   195
      Index           =   11
      Left            =   480
      TabIndex        =   74
      Top             =   5040
      Width           =   1245
   End
   Begin VB.Label Cor_Letra_Label_Formulario 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   73
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Letra_Label_Formulario"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   72
      Top             =   3630
      Width           =   2430
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_BackColorBkg"
      Height          =   195
      Index           =   2
      Left            =   3600
      TabIndex        =   71
      Top             =   2910
      Width           =   2100
   End
   Begin VB.Label Cor_Grid_BackColorBkg 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   70
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Cor_Grid_BackColorFixed 
      BackColor       =   &H00313131&
      Height          =   255
      Left            =   3240
      TabIndex        =   69
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_BackColorFixed"
      Height          =   195
      Index           =   6
      Left            =   3600
      TabIndex        =   68
      Top             =   3270
      Width           =   2220
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor da progressbar"
      Height          =   195
      Index           =   5
      Left            =   480
      TabIndex        =   67
      Top             =   3990
      Width           =   1680
   End
   Begin VB.Label Cor_Contorno_Caixas 
      BackColor       =   &H00DFB000&
      Height          =   255
      Left            =   120
      TabIndex        =   66
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Cor_Fundo_Textbox 
      BackColor       =   &H00101010&
      Height          =   255
      Left            =   120
      TabIndex        =   65
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fundo caixas de texto"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   64
      Top             =   4350
      Width           =   1875
   End
   Begin VB.Label Cor_Label_Barra_Visor 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   63
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Label_Barra_Visor"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   62
      Top             =   5790
      Width           =   1995
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo da barra de titulo"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   61
      Top             =   3270
      Width           =   2010
   End
   Begin VB.Label Cor_Label_Barra_Titulo 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   60
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_Color"
      Height          =   195
      Left            =   3600
      TabIndex        =   59
      Top             =   5040
      Width           =   1350
   End
   Begin VB.Label Cor_Fundo_Topico_Normal 
      BackColor       =   &H00101010&
      Height          =   255
      Left            =   120
      TabIndex        =   58
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fundo topico normal"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   57
      Top             =   6150
      Width           =   1740
   End
   Begin VB.Label Cor_Letra_Topico_Normal 
      BackColor       =   &H00808080&
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   56
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Letra topico normal"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   55
      Top             =   6510
      Width           =   1665
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fundo topico over"
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   54
      Top             =   6870
      Width           =   1530
   End
   Begin VB.Label Cor_Fundo_Topico_Over 
      BackColor       =   &H00484947&
      Height          =   255
      Left            =   120
      TabIndex        =   53
      Top             =   6840
      Width           =   255
   End
   Begin VB.Label Cor_Letra_Topico_Over 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   52
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Letra topico over"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   51
      Top             =   7230
      Width           =   1455
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fundo scroll bar"
      Height          =   195
      Index           =   1
      Left            =   3600
      TabIndex        =   50
      Top             =   1830
      Width           =   1365
   End
   Begin VB.Label Cor_Scroll_Bar 
      BackColor       =   &H00212121&
      Height          =   255
      Left            =   3240
      TabIndex        =   49
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor_Grid_ColorFixed"
      Height          =   195
      Index           =   0
      Left            =   3600
      TabIndex        =   48
      Top             =   5400
      Width           =   1800
   End
   Begin VB.Label Cor_Grid_ColorFixed 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   47
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Cor_Letra_Tab_Over 
      BackColor       =   &H00DFB000&
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Letra tab over"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   45
      Top             =   2550
      Width           =   1215
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Letra tab normal"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   44
      Top             =   2190
      Width           =   1425
   End
   Begin VB.Label Cor_Letra_Tab_Normal 
      BackColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Cor_Fundo_Task_Bar 
      BackColor       =   &H00101010&
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "fundo task bar"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   41
      Top             =   1830
      Width           =   1245
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Letra barra informação"
      Height          =   195
      Index           =   4
      Left            =   3600
      TabIndex        =   40
      Top             =   2190
      Width           =   1980
   End
   Begin VB.Label Cor_Letra_Bar_Info 
      BackColor       =   &H00101010&
      Height          =   255
      Left            =   3240
      TabIndex        =   39
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   " Cores do programa "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   38
      Top             =   0
      Width           =   2205
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contorno frame centro"
      Height          =   195
      Index           =   2
      Left            =   3600
      TabIndex        =   37
      Top             =   6120
      Width           =   1965
   End
   Begin VB.Label Cor_Contorno_Frame_Centro 
      BackColor       =   &H00404040&
      Height          =   255
      Left            =   3240
      TabIndex        =   36
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Cor_BackColor_Download_Add_Ons 
      BackColor       =   &H00212121&
      Height          =   255
      Left            =   3240
      TabIndex        =   35
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BackColor_Download_Add_Ons"
      Height          =   195
      Index           =   3
      Left            =   3600
      TabIndex        =   34
      Top             =   6480
      Width           =   2700
   End
   Begin VB.Label Cor_Letter_Download_Add_Ons 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   33
      Top             =   6840
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Letter_Download_Add_Ons"
      Height          =   195
      Index           =   4
      Left            =   3600
      TabIndex        =   32
      Top             =   6840
      Width           =   2310
   End
   Begin VB.Label Cor_BackColor_Display 
      BackColor       =   &H00A2AFA7&
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BackColor_Display"
      Height          =   195
      Index           =   5
      Left            =   480
      TabIndex        =   30
      Top             =   7560
      Width           =   1620
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Download_Add_Ons_Buttons"
      Height          =   195
      Index           =   6
      Left            =   3600
      TabIndex        =   29
      Top             =   7200
      Width           =   2460
   End
   Begin VB.Label Cor_Download_Add_Ons_Buttons 
      BackColor       =   &H00222222&
      Height          =   255
      Left            =   3240
      TabIndex        =   28
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Cor_Slider_Music 
      BackColor       =   &H00222222&
      Height          =   255
      Left            =   3240
      TabIndex        =   27
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Slider music"
      Height          =   195
      Index           =   7
      Left            =   3600
      TabIndex        =   26
      Top             =   7560
      Width           =   1050
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fundo form main"
      Height          =   195
      Index           =   5
      Left            =   480
      TabIndex        =   25
      Top             =   1470
      Width           =   1455
   End
   Begin VB.Label Cor_Form_Main 
      BackColor       =   &H00101010&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Cor_Topic_Task_Bar 
      BackColor       =   &H00101010&
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Topic_Task_Bar"
      Height          =   195
      Index           =   6
      Left            =   3600
      TabIndex        =   22
      Top             =   1470
      Width           =   1365
   End
   Begin VB.Label Cor_Label_Button_ForeColor 
      BackColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label_Button_ForeColor"
      Height          =   195
      Index           =   7
      Left            =   480
      TabIndex        =   20
      Top             =   1110
      Width           =   2055
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BackGround_Bar_Label_Buton"
      Height          =   195
      Index           =   8
      Left            =   3600
      TabIndex        =   19
      Top             =   1110
      Width           =   2610
   End
   Begin VB.Label Cor_BackGround_Bar_Label_Button 
      BackColor       =   &H00101010&
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Cor_Form_BorderColor 
      BackColor       =   &H00101010&
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form_BorderColor"
      Height          =   195
      Index           =   1
      Left            =   3600
      TabIndex        =   16
      Top             =   5760
      Width           =   1590
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BackGround_Frame_Cover"
      Height          =   195
      Index           =   9
      Left            =   480
      TabIndex        =   15
      Top             =   750
      Width           =   2325
   End
   Begin VB.Label Cor_BackGround_Frame_Cover 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Cor_Label_Frame_Cover 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label_Frame_Cover"
      Height          =   195
      Index           =   10
      Left            =   480
      TabIndex        =   12
      Top             =   390
      Width           =   1725
   End
   Begin VB.Label Cor_Form_About_Letter 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form_About_Letter"
      Height          =   195
      Index           =   0
      Left            =   3600
      TabIndex        =   10
      Top             =   750
      Width           =   1635
   End
   Begin VB.Label Line_Border_Frames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label_Frame_Cover"
      Height          =   195
      Index           =   11
      Left            =   3600
      TabIndex        =   9
      Top             =   360
      Width           =   1725
   End
   Begin VB.Label Cor_Line_Border_Frames 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu_BackColorSel"
      Height          =   195
      Index           =   11
      Left            =   480
      TabIndex        =   7
      Top             =   7950
      Width           =   1710
   End
   Begin VB.Label Cor_Menu_BackColorSel 
      BackColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label Cor_Menu_ForeColor 
      BackColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   8640
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu_ForeColor"
      Height          =   195
      Index           =   12
      Left            =   480
      TabIndex        =   4
      Top             =   8670
      Width           =   1395
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu_BackColor"
      Height          =   195
      Index           =   8
      Left            =   480
      TabIndex        =   3
      Top             =   8280
      Width           =   1440
   End
   Begin VB.Label Cor_Menu_BackColor 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   8280
      Width           =   255
   End
   Begin VB.Label Cor_Menu_ForeColorSel 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   9000
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu_ForeColorSel"
      Height          =   195
      Index           =   9
      Left            =   480
      TabIndex        =   0
      Top             =   9000
      Width           =   1665
   End
   Begin VB.Image Botao_Linha_2_Over 
      Height          =   375
      Left            =   13320
      Picture         =   "Form_Skin.frx":CF739
      Top             =   1680
      Width           =   1440
   End
   Begin VB.Image Botao_Linha_Over 
      Height          =   375
      Left            =   13320
      Picture         =   "Form_Skin.frx":D139B
      Top             =   720
      Width           =   2160
   End
   Begin VB.Image Botao_Linha_2_Normal 
      Height          =   375
      Left            =   13320
      Picture         =   "Form_Skin.frx":D3E0D
      Top             =   1200
      Width           =   1440
   End
   Begin VB.Image Image_Down_Erro 
      Height          =   240
      Left            =   7440
      Picture         =   "Form_Skin.frx":D5A6F
      Top             =   7560
      Width           =   240
   End
   Begin VB.Image Image_Down_Concluido 
      Height          =   240
      Left            =   7080
      Picture         =   "Form_Skin.frx":D5DB1
      Top             =   7680
      Width           =   240
   End
   Begin VB.Image Image_Down_Processando 
      Height          =   240
      Left            =   7080
      Picture         =   "Form_Skin.frx":D60F3
      Top             =   7440
      Width           =   240
   End
   Begin VB.Image Image_Caixa_Pesquisa 
      Height          =   375
      Left            =   14040
      Picture         =   "Form_Skin.frx":D6435
      Top             =   5520
      Width           =   2370
   End
   Begin VB.Image Botao_Executar 
      Height          =   435
      Left            =   16200
      Picture         =   "Form_Skin.frx":D92F3
      Top             =   1320
      Width           =   1365
   End
   Begin VB.Image Fundo_Barra_Botoes 
      Enabled         =   0   'False
      Height          =   900
      Left            =   15600
      Picture         =   "Form_Skin.frx":D983C
      Top             =   3720
      Width           =   585
   End
   Begin VB.Image Botao_Linha_Normal 
      Height          =   375
      Left            =   13320
      Picture         =   "Form_Skin.frx":D9BA7
      Top             =   240
      Width           =   2160
   End
   Begin VB.Image Botao_Download 
      Height          =   435
      Left            =   11760
      Picture         =   "Form_Skin.frx":DA13C
      Top             =   240
      Width           =   1365
   End
   Begin VB.Image Fundo_Separadores 
      Height          =   405
      Left            =   9720
      Picture         =   "Form_Skin.frx":DA696
      Top             =   0
      Width           =   8085
   End
   Begin VB.Image Imagem_Vazia 
      Height          =   1215
      Left            =   4800
      Top             =   8760
      Width           =   975
   End
   Begin VB.Image Botao_Normal 
      Height          =   495
      Left            =   8760
      Picture         =   "Form_Skin.frx":E51B4
      Top             =   0
      Width           =   915
   End
   Begin VB.Image Botao_Over 
      Height          =   495
      Left            =   12240
      Picture         =   "Form_Skin.frx":E69AE
      Top             =   2280
      Width           =   915
   End
   Begin VB.Image Foto_Programa 
      Height          =   1860
      Left            =   14400
      Picture         =   "Form_Skin.frx":E81A8
      Top             =   2040
      Width           =   3075
   End
   Begin VB.Image Linha_Over 
      Height          =   1290
      Left            =   6240
      Picture         =   "Form_Skin.frx":EA72B
      Top             =   8880
      Width           =   1380
   End
   Begin VB.Image Linha_Normal 
      Height          =   645
      Left            =   6240
      Picture         =   "Form_Skin.frx":F0425
      Top             =   8160
      Width           =   1485
   End
   Begin VB.Image Extermidade_Normal 
      Height          =   405
      Left            =   6360
      Picture         =   "Form_Skin.frx":F36CB
      Top             =   7440
      Width           =   225
   End
   Begin VB.Image Extermidade_Over 
      Height          =   405
      Left            =   6720
      Picture         =   "Form_Skin.frx":F3A0C
      Top             =   7440
      Width           =   225
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   0
      X2              =   424
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Image Image_Balao 
      Height          =   1125
      Left            =   8400
      Picture         =   "Form_Skin.frx":F3D92
      Top             =   5880
      Width           =   5010
   End
   Begin VB.Menu Menu_Tray 
      Caption         =   "Menu_Tray"
      Begin VB.Menu Menu_Fechar 
         Caption         =   "Fechar"
      End
   End
End
Attribute VB_Name = "Form_Skin"
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

Public Sub Carregar_Imagens_do_Skin()
    'Cores
    Cor_Form_Main.backcolor = RGB(249, 249, 249)
    Cor_Form_BorderColor.backcolor = RGB(192, 192, 192)
    Cor_do_Fundo_dos_Formularios.backcolor = RGB(238, 238, 238)
    Cor_Label_Barra_Titulo.backcolor = RGB(255, 255, 255)
    Cor_Letra_Label_Formulario.backcolor = RGB(0, 0, 0)
    Cor_Label_Contador_Popup.backcolor = RGB(153, 153, 153)
    Cor_Contorno_Caixas.backcolor = RGB(0, 176, 223)
    Cor_Fundo_Textbox.backcolor = RGB(255, 255, 255)
    Cor_Letra_Textbox.backcolor = RGB(0, 0, 0)
    Cor_da_Letra_do_Botao.backcolor = RGB(0, 0, 0)
    Cor_Label_Barra_Visor.backcolor = RGB(255, 255, 255)
    Cor_Grid_BackColor.backcolor = RGB(255, 255, 255)
    Cor_Grid_BackColorBkg.backcolor = RGB(255, 255, 255)
    Cor_Grid_BackColorFixed.backcolor = RGB(199, 199, 201)
    Cor_Grid_BackColorSel.backcolor = RGB(0, 176, 223)
    Cor_Grid_ForeColor.backcolor = RGB(0, 0, 0)
    Cor_Grid_ForeColorFixed.backcolor = RGB(46, 53, 69)
    Cor_Grid_ForeColorSel.backcolor = RGB(255, 255, 255)
    Cor_Grid_Color.backcolor = RGB(223, 223, 223)
    Cor_Grid_ColorFixed.backcolor = RGB(230, 230, 230)
    Cor_Fundo_Topico_Normal.backcolor = RGB(218, 222, 231)
    Cor_Letra_Topico_Normal.backcolor = RGB(0, 0, 0)
    Cor_Fundo_Topico_Over.backcolor = RGB(84, 84, 84)
    Cor_Letra_Topico_Over.backcolor = RGB(255, 255, 255)
    Cor_Scroll_Bar.backcolor = RGB(32, 34, 31)
    Cor_Letra_Tab_Normal.backcolor = RGB(255, 255, 255)
    Cor_Letra_Tab_Over.backcolor = RGB(0, 176, 223)
    Cor_Fundo_Task_Bar.backcolor = RGB(218, 222, 231)
    Cor_Letra_Bar_Info.backcolor = RGB(255, 255, 255)
    Cor_Contorno_Frame_Centro.backcolor = RGB(238, 238, 238)
    Cor_BackColor_Download_Add_Ons.backcolor = RGB(33, 33, 33)
    Cor_Letter_Download_Add_Ons.backcolor = RGB(255, 255, 255)
    Cor_BackColor_Display.backcolor = RGB(40, 40, 40)
    Cor_Download_Add_Ons_Buttons.backcolor = RGB(49, 49, 49)
    Cor_Slider_Music.backcolor = RGB(123, 123, 123)
    Cor_Topic_Task_Bar.backcolor = RGB(121, 123, 148)
    Cor_Label_Button_ForeColor.backcolor = RGB(49, 49, 49)
    Cor_BackGround_Bar_Label_Button.backcolor = RGB(255, 255, 255)
    Cor_BackGround_Frame_Cover.backcolor = RGB(255, 255, 255)
    Cor_Label_Frame_Cover.backcolor = RGB(128, 128, 128)
    Cor_Form_About_Letter.backcolor = RGB(255, 255, 255)
    Cor_Line_Border_Frames.backcolor = RGB(192, 192, 192)
End Sub

Private Sub Form_Load()
    'Propriedades iniciais do formulário
    Carregar_Imagens_do_Skin
End Sub

Private Sub Icon_Categoria_Click(Index As Integer)
End Sub

Private Sub Menu_Fechar_Click()
    'Fechar o programa
    Unload Form_Barra
    Unload Form_Principal
    End
End Sub
