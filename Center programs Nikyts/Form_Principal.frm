VERSION 5.00
Begin VB.Form Form_Principal 
   Appearance      =   0  'Flat
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'None
   Caption         =   "Center programs Nikyts"
   ClientHeight    =   11940
   ClientLeft      =   10035
   ClientTop       =   2325
   ClientWidth     =   18405
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
   Icon            =   "Form_Principal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   796
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1227
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Frame_Opcoes 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   12720
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   481
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   7560
      Visible         =   0   'False
      Width           =   7215
      Begin VB.PictureBox Lista_Linguas 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   2640
         ScaleHeight     =   95
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   303
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   4575
         Begin VB.Label Label_Lingua 
            BackColor       =   &H00EEEEEE&
            BackStyle       =   0  'Transparent
            Caption         =   "Idioma"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   60
            Top             =   0
            Width           =   960
         End
         Begin VB.Label Shape_Sombra_Lingua 
            BackColor       =   &H00DFB000&
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   61
            Top             =   0
            Width           =   3975
         End
      End
      Begin VB.PictureBox Pic_Tray 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         Picture         =   "Form_Principal.frx":57E2
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   159
         TabStop         =   0   'False
         Top             =   3840
         Width           =   195
      End
      Begin VB.CheckBox Check_Tray 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
         Caption         =   "Minimizar o programa no systray (ao lado do clock)"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   70
         Top             =   3840
         Value           =   1  'Checked
         Width           =   6495
      End
      Begin VB.PictureBox Barra_Text_Lingua 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   720
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1320
         Width           =   5475
         Begin VB.PictureBox Seta_Lingua 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   5040
            Picture         =   "Form_Principal.frx":5A2C
            ScaleHeight     =   19
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   19
            TabIndex        =   65
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
            TabIndex        =   66
            Top             =   30
            Width           =   1500
         End
         Begin VB.Shape Contorno_Lingua 
            BorderColor     =   &H00C0C0C0&
            Height          =   375
            Left            =   0
            Top             =   0
            Width           =   4575
         End
      End
      Begin VB.DirListBox Dir_Lingua 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         Height          =   540
         Left            =   6240
         TabIndex        =   63
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.FileListBox File_Lingua 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         Height          =   420
         Left            =   6240
         Pattern         =   "*.lng"
         TabIndex        =   62
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox Pic_Actualizar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         Picture         =   "Form_Principal.frx":5F5A
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   2280
         Width           =   195
      End
      Begin VB.CheckBox Check_Actualizar 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
         Caption         =   "Verificar actualizacoes automaticamente"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   68
         Top             =   2280
         Width           =   6495
      End
      Begin VB.PictureBox Pic_Barra 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         Picture         =   "Form_Principal.frx":61A4
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   157
         TabStop         =   0   'False
         Top             =   3360
         Width           =   195
      End
      Begin VB.CheckBox Check_Barra 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
         Caption         =   "Ver a barra dos programas instalados no ambiente de trabalho"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   69
         Top             =   3360
         Value           =   1  'Checked
         Width           =   6495
      End
      Begin VB.Label Label_Opcoes_Actualizadas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "As opções do programa foram actualizadas com sucesso."
         ForeColor       =   &H00DFB000&
         Height          =   195
         Left            =   2160
         TabIndex        =   160
         Top             =   480
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label Label_Idioma_Programa 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Idioma do programa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   67
         Top             =   1080
         Width           =   1770
      End
      Begin VB.Label Label_Frame_Opcoes 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Opções"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   720
         TabIndex        =   46
         Top             =   360
         Width           =   1170
      End
   End
   Begin VB.PictureBox Frame_Instalados 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   12360
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   13320
      Visible         =   0   'False
      Width           =   4335
      Begin VB.ListBox Lista_Pastas 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         Height          =   810
         Left            =   2640
         TabIndex        =   81
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox Frame_Icon_Pequeno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   8160
         ScaleHeight     =   137
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   185
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.PictureBox Frame_Icon_Grande 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   600
         ScaleHeight     =   153
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   145
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   960
         Width           =   2175
         Begin VB.Image Image_Logo_Over 
            Height          =   375
            Index           =   0
            Left            =   480
            Top             =   1800
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Image Image_Logo_Normal 
            Height          =   375
            Index           =   0
            Left            =   0
            Top             =   1800
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label_Icon_Grande 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Programa"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   82
            Top             =   1500
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.Image Image_Icon_Grande 
            Height          =   1800
            Index           =   0
            Left            =   0
            Picture         =   "Form_Principal.frx":63EE
            Top             =   0
            Visible         =   0   'False
            Width           =   1800
         End
      End
      Begin VB.Label Label_Frame_Instalados 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Instalados"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   720
         TabIndex        =   47
         Top             =   480
         Width           =   1635
      End
   End
   Begin VB.PictureBox Frame_Partilhar 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   0
      ScaleHeight     =   577
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   921
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   13815
      Begin VB.PictureBox Conteudo_Frame_Partilhar 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   10485
         Left            =   0
         ScaleHeight     =   699
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   833
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   0
         Width           =   12495
         Begin VB.PictureBox Barra_Txt_Email 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   720
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   365
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   1440
            Width           =   5475
            Begin VB.TextBox Txt_Email 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H00000000&
               Height          =   300
               Left            =   600
               TabIndex        =   0
               Top             =   30
               Width           =   1380
            End
            Begin VB.Shape Contorno_Txt_Email 
               BorderColor     =   &H00C0C0C0&
               Height          =   375
               Left            =   0
               Top             =   0
               Width           =   4935
            End
         End
         Begin VB.PictureBox Barra_Txt_Empresa 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   720
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   365
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   2280
            Width           =   5475
            Begin VB.TextBox Txt_Empresa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H00000000&
               Height          =   300
               Left            =   720
               TabIndex        =   1
               Top             =   30
               Width           =   1380
            End
            Begin VB.Shape Contorno_Txt_Empresa 
               BorderColor     =   &H00C0C0C0&
               Height          =   375
               Left            =   0
               Top             =   0
               Width           =   4935
            End
         End
         Begin VB.PictureBox Barra_Txt_Informacao 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   1920
            Left            =   720
            ScaleHeight     =   128
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   365
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   4800
            Width           =   5475
            Begin VB.TextBox Txt_Informacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H00000000&
               Height          =   1740
               Left            =   600
               MultiLine       =   -1  'True
               TabIndex        =   4
               Top             =   30
               Width           =   1860
            End
            Begin VB.Shape Contorno_Txt_Informacao 
               BorderColor     =   &H00C0C0C0&
               Height          =   1890
               Left            =   0
               Top             =   0
               Width           =   4935
            End
         End
         Begin VB.PictureBox Barra_Txt_Nome 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   720
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   365
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   3120
            Width           =   5475
            Begin VB.TextBox Txt_Nome 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H00000000&
               Height          =   300
               Left            =   720
               TabIndex        =   2
               Top             =   30
               Width           =   1380
            End
            Begin VB.Shape Contorno_Txt_Nome 
               BorderColor     =   &H00C0C0C0&
               Height          =   375
               Left            =   0
               Top             =   0
               Width           =   4935
            End
         End
         Begin VB.PictureBox Barra_Txt_Descricao 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   720
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   365
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   3960
            Width           =   5475
            Begin VB.TextBox Txt_Descricao 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H00000000&
               Height          =   300
               Left            =   720
               TabIndex        =   3
               Top             =   30
               Width           =   1380
            End
            Begin VB.Shape Contorno_Txt_Descricao 
               BorderColor     =   &H00C0C0C0&
               Height          =   375
               Left            =   0
               Top             =   0
               Width           =   4935
            End
         End
         Begin VB.PictureBox Barra_Txt_Site 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   720
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   365
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   7200
            Width           =   5475
            Begin VB.TextBox Txt_Site 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H00000000&
               Height          =   300
               Left            =   720
               TabIndex        =   5
               Top             =   30
               Width           =   1380
            End
            Begin VB.Shape Contorno_Txt_Site 
               BorderColor     =   &H00C0C0C0&
               Height          =   375
               Left            =   0
               Top             =   0
               Width           =   4935
            End
         End
         Begin VB.PictureBox Barra_Txt_Download 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   720
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   365
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   8040
            Width           =   5475
            Begin VB.TextBox Txt_Download 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H00000000&
               Height          =   300
               Left            =   720
               TabIndex        =   6
               Top             =   30
               Width           =   1380
            End
            Begin VB.Shape Contorno_Txt_Download 
               BorderColor     =   &H00C0C0C0&
               Height          =   375
               Left            =   0
               Top             =   0
               Width           =   4935
            End
         End
         Begin VB.PictureBox Frame_Erro_Partilhar 
            Appearance      =   0  'Flat
            BackColor       =   &H00DFB000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1125
            Left            =   6360
            ScaleHeight     =   75
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   334
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   1050
            Visible         =   0   'False
            Width           =   5010
            Begin VB.Label Label_Erro_Partilhar 
               AutoSize        =   -1  'True
               BackColor       =   &H00EEEEEE&
               BackStyle       =   0  'Transparent
               Caption         =   "Indique um email válido."
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   480
               TabIndex        =   104
               Top             =   420
               Width           =   2130
            End
            Begin VB.Label Label_Close_Partilhar 
               AutoSize        =   -1  'True
               BackColor       =   &H00F5F5F5&
               BackStyle       =   0  'Transparent
               Caption         =   " x "
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   4680
               TabIndex        =   103
               Top             =   120
               Width           =   225
            End
         End
         Begin VB.Label Label_Frame_Partilhar 
            AutoSize        =   -1  'True
            BackColor       =   &H00EEEEEE&
            BackStyle       =   0  'Transparent
            Caption         =   "Partilhar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   375
            Left            =   720
            TabIndex        =   121
            Top             =   480
            Width           =   1350
         End
         Begin VB.Label Lb_Info 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(Para possivel contacto caso seja necessário)"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   180
            Left            =   1320
            TabIndex        =   120
            Top             =   1200
            Width           =   3480
         End
         Begin VB.Label Lb_Email 
            AutoSize        =   -1  'True
            BackColor       =   &H00EEEEEE&
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   720
            TabIndex        =   119
            Top             =   1200
            Width           =   465
         End
         Begin VB.Label Lb_Empresa 
            AutoSize        =   -1  'True
            BackColor       =   &H00EEEEEE&
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   720
            TabIndex        =   118
            Top             =   2040
            Width           =   750
         End
         Begin VB.Label Lb_Informacao 
            AutoSize        =   -1  'True
            BackColor       =   &H00EEEEEE&
            BackStyle       =   0  'Transparent
            Caption         =   "Informação sobre o programa"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   720
            TabIndex        =   117
            Top             =   4560
            Width           =   2595
         End
         Begin VB.Label Lb_Nome 
            AutoSize        =   -1  'True
            BackColor       =   &H00EEEEEE&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do programa"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   720
            TabIndex        =   116
            Top             =   2880
            Width           =   1665
         End
         Begin VB.Label Lb_Descricao 
            AutoSize        =   -1  'True
            BackColor       =   &H00EEEEEE&
            BackStyle       =   0  'Transparent
            Caption         =   "Faça uma pequena descrição do programa"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   720
            TabIndex        =   115
            Top             =   3720
            Width           =   3660
         End
         Begin VB.Label Lb_Site 
            AutoSize        =   -1  'True
            BackColor       =   &H00EEEEEE&
            BackStyle       =   0  'Transparent
            Caption         =   "Site oficial"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   720
            TabIndex        =   114
            Top             =   6960
            Width           =   885
         End
         Begin VB.Label Lb_Download 
            AutoSize        =   -1  'True
            BackColor       =   &H00EEEEEE&
            BackStyle       =   0  'Transparent
            Caption         =   "Hiperligação para efectuar o download do programa"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   720
            TabIndex        =   113
            Top             =   7800
            Width           =   4470
         End
         Begin VB.Label Lb_Nota 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nota: Caso se verifique que o programa possua algum virus será removido!"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   180
            Left            =   720
            TabIndex        =   112
            Top             =   8520
            Width           =   5655
         End
      End
      Begin Project_Gadgets.YsVSrcrollBar SrcrollBar2 
         Height          =   2535
         Left            =   13440
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   0
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   4471
      End
   End
   Begin VB.PictureBox Frame_Centro 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFB000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   13455
      Left            =   0
      ScaleHeight     =   897
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1537
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1920
      Width           =   23055
      Begin VB.PictureBox Frame_Conteudo 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   10215
         Left            =   3600
         ScaleHeight     =   681
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   969
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   14535
         Begin VB.PictureBox Barra_Estado 
            Appearance      =   0  'Flat
            BackColor       =   &H00F9F9F9&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   360
            ScaleHeight     =   40
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   513
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   9600
            Visible         =   0   'False
            Width           =   7695
            Begin VB.PictureBox Botao_Cancelar 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   4080
               Picture         =   "Form_Principal.frx":10CF0
               ScaleHeight     =   29
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   91
               TabIndex        =   126
               TabStop         =   0   'False
               Top             =   90
               Width           =   1365
               Begin VB.Label Label_Cancelar 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cancelar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   285
                  TabIndex        =   127
                  Top             =   120
                  Width           =   795
               End
            End
            Begin VB.PictureBox Botao_Executar 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   5640
               Picture         =   "Form_Principal.frx":1123C
               ScaleHeight     =   29
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   91
               TabIndex        =   124
               TabStop         =   0   'False
               Top             =   90
               Width           =   1365
               Begin VB.Label Label_Executar 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Executar"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   300
                  TabIndex        =   125
                  Top             =   120
                  Width           =   765
               End
            End
            Begin VB.Image Image_Download 
               Height          =   240
               Left            =   120
               Top             =   180
               Width           =   240
            End
            Begin VB.Label Label_Estado 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   480
               TabIndex        =   128
               Top             =   180
               Width           =   75
            End
            Begin VB.Shape Shape_Estado 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00C0C0C0&
               Height          =   600
               Left            =   600
               Top             =   0
               Width           =   7095
            End
         End
         Begin Project_Gadgets.YsVSrcrollBar SrcrollBar1 
            Height          =   2535
            Left            =   14160
            TabIndex        =   154
            TabStop         =   0   'False
            Top             =   0
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   4471
         End
         Begin VB.PictureBox Frame_Informacoes 
            Appearance      =   0  'Flat
            BackColor       =   &H00F9F9F9&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   8535
            Left            =   360
            ScaleHeight     =   569
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   912
            TabIndex        =   129
            Top             =   240
            Width           =   13680
            Begin VB.PictureBox Botao_Eu_Gosto 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   405
               Left            =   120
               Picture         =   "Form_Principal.frx":11788
               ScaleHeight     =   27
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   79
               TabIndex        =   144
               TabStop         =   0   'False
               Top             =   7680
               Width           =   1185
               Begin VB.Label Label_Eu_Gosto 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Eu gosto"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   145
                  Top             =   105
                  Width           =   735
               End
               Begin VB.Image Extermidade_Eu_Gosto 
                  Enabled         =   0   'False
                  Height          =   405
                  Left            =   960
                  Picture         =   "Form_Principal.frx":11FC9
                  Top             =   0
                  Width           =   225
               End
            End
            Begin VB.PictureBox Barra_Comentario 
               Appearance      =   0  'Flat
               BackColor       =   &H00F9F9F9&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   120
               ScaleHeight     =   29
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   665
               TabIndex        =   142
               TabStop         =   0   'False
               Top             =   7080
               Width           =   9975
               Begin VB.Line Linha_Barra_Comentario 
                  BorderColor     =   &H00C0C0C0&
                  BorderWidth     =   3
                  X1              =   432
                  X2              =   0
                  Y1              =   0
                  Y2              =   0
               End
               Begin VB.Label Label_Barra_Comentario 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Se gostou do programa contribuia para a sua votação."
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00808080&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   143
                  Top             =   105
                  Width           =   5910
               End
            End
            Begin VB.PictureBox Frame_Avaliacao 
               Appearance      =   0  'Flat
               BackColor       =   &H00F9F9F9&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   885
               Left            =   7680
               Picture         =   "Form_Principal.frx":1230A
               ScaleHeight     =   59
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   165
               TabIndex        =   139
               TabStop         =   0   'False
               Top             =   0
               Width           =   2475
               Begin VB.Label Label_Votos 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF80FF&
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1680
                  TabIndex        =   141
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.Label Label_Total 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H00D3A53A&
                  BackStyle       =   0  'Transparent
                  Caption         =   "0 avaliações"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   720
                  TabIndex        =   140
                  Top             =   600
                  Width           =   1095
               End
            End
            Begin VB.PictureBox Barra_Transferir 
               Appearance      =   0  'Flat
               BackColor       =   &H00F9F9F9&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   600
               Left            =   240
               ScaleHeight     =   40
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   665
               TabIndex        =   132
               TabStop         =   0   'False
               Top             =   2520
               Width           =   9975
               Begin VB.TextBox txtZip 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF80FF&
                  BorderStyle     =   0  'None
                  Height          =   330
                  Left            =   3720
                  TabIndex        =   136
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   2025
               End
               Begin VB.TextBox Text_Servidor 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF80FF&
                  BorderStyle     =   0  'None
                  Height          =   330
                  Left            =   1680
                  TabIndex        =   135
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1935
               End
               Begin VB.PictureBox Botao_Download 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   435
                  Left            =   8040
                  Picture         =   "Form_Principal.frx":15CF0
                  ScaleHeight     =   29
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   91
                  TabIndex        =   133
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   1365
                  Begin VB.Label Label_Download 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Transferir"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   240
                     TabIndex        =   134
                     Top             =   120
                     Width           =   855
                  End
               End
               Begin Project_Gadgets.NProgressBar ProgressBar1 
                  Height          =   375
                  Left            =   6720
                  TabIndex        =   137
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   661
               End
               Begin VB.Label Label_Transferir 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E9CDAD&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ficheiro.zip"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   240
                  Left            =   120
                  TabIndex        =   138
                  Top             =   180
                  Width           =   1110
               End
               Begin VB.Shape Shape_Transferir 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H00C0C0C0&
                  Height          =   600
                  Left            =   0
                  Top             =   0
                  Width           =   9615
               End
            End
            Begin VB.PictureBox Frame_Foto 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1935
               Left            =   8880
               ScaleHeight     =   129
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   209
               TabIndex        =   131
               TabStop         =   0   'False
               Top             =   4320
               Width           =   3135
               Begin VB.Image Image_Tela 
                  Enabled         =   0   'False
                  Height          =   1860
                  Left            =   30
                  Top             =   36
                  Width           =   3072
               End
               Begin VB.Shape Shape_Foto 
                  BorderColor     =   &H00E0E0E0&
                  Height          =   1920
                  Left            =   0
                  Top             =   240
                  Width           =   3135
               End
            End
            Begin VB.TextBox Text_Informacao 
               BackColor       =   &H00F9F9F9&
               BorderStyle     =   0  'None
               Height          =   1335
               Left            =   240
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   130
               TabStop         =   0   'False
               Top             =   4320
               Width           =   7575
            End
            Begin Project_Gadgets.dl dl 
               Left            =   600
               Top             =   600
               _ExtentX        =   1799
               _ExtentY        =   1667
            End
            Begin VB.Label Label_Id_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4560
               TabIndex        =   153
               Top             =   1080
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label_Transferencias 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   6120
               TabIndex        =   152
               Top             =   1080
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label Label_Site_Programa 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF80FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2640
               TabIndex        =   151
               Top             =   6000
               Visible         =   0   'False
               Width           =   3255
            End
            Begin VB.Label Label_Site_Oficial 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Site oficial do programa"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   240
               TabIndex        =   150
               Top             =   6000
               Width           =   2055
            End
            Begin VB.Label Label_Descricao_Programa 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00F1E6D3&
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição do programa"
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   2400
               TabIndex        =   149
               Top             =   1080
               Width           =   2010
            End
            Begin VB.Label Label_Nome_Programa 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00F1E6D3&
               BackStyle       =   0  'Transparent
               Caption         =   "Nome do programa"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   345
               Left            =   2400
               TabIndex        =   148
               Top             =   720
               Width           =   3075
            End
            Begin VB.Image Image_Logo 
               Enabled         =   0   'False
               Height          =   1800
               Left            =   300
               Top             =   300
               Width           =   1800
            End
            Begin VB.Label Label_Total_Downloads 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00F1E6D3&
               BackStyle       =   0  'Transparent
               Caption         =   "(0 Downloads)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   5760
               TabIndex        =   147
               Top             =   810
               Width           =   1590
            End
            Begin VB.Label Label_Enterprise 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00F1E6D3&
               BackStyle       =   0  'Transparent
               Caption         =   "Empresa"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   2400
               TabIndex        =   146
               Top             =   1320
               Width           =   750
            End
         End
      End
      Begin VB.PictureBox Frame_Lista 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   0
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   737
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   2400
         Visible         =   0   'False
         Width           =   11055
         Begin Project_Gadgets.NProgressBar Progresso 
            Height          =   375
            Index           =   0
            Left            =   8880
            TabIndex        =   165
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
         End
         Begin Project_Gadgets.dl Download_Programa 
            Left            =   3360
            Top             =   120
            _ExtentX        =   1799
            _ExtentY        =   1667
         End
         Begin VB.Label Label_Remover_Transferencia 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remover"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   6960
            TabIndex        =   181
            Top             =   720
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Image Botao_Remover_Transferencia 
            Height          =   375
            Index           =   0
            Left            =   6720
            Picture         =   "Form_Principal.frx":1623C
            Top             =   600
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label Label_Executar_Programa 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Executar"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   5520
            TabIndex        =   180
            Top             =   720
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Image Botao_Executar_Programa 
            Height          =   375
            Index           =   0
            Left            =   5160
            Picture         =   "Form_Principal.frx":17E9E
            Top             =   600
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label Label_Mais_Informacoes 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mais informações"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   1080
            TabIndex        =   179
            Top             =   720
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Image Botao_Mais_Informacoes 
            Height          =   375
            Index           =   0
            Left            =   720
            Picture         =   "Form_Principal.frx":19B00
            Top             =   600
            Visible         =   0   'False
            Width           =   2160
         End
         Begin VB.Label Label_Avaliacao 
            BackColor       =   &H00FF8080&
            Height          =   255
            Index           =   0
            Left            =   6600
            TabIndex        =   178
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label_Id 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   176
            Top             =   0
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label_Icon 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   4320
            TabIndex        =   175
            Top             =   0
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label_site 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   5760
            TabIndex        =   174
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Image Tela_Programa 
            Height          =   255
            Index           =   0
            Left            =   9600
            Top             =   0
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Image Logotipo_Programa 
            Height          =   255
            Index           =   0
            Left            =   8760
            Top             =   0
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label_Tela 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   7200
            TabIndex        =   173
            Top             =   0
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label_Logotipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   5760
            TabIndex        =   172
            Top             =   0
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label_Observacoes 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2880
            TabIndex        =   171
            Top             =   0
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Image Icon_Programa 
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   120
            Top             =   120
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label_Nome 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "VbMovieManager"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   720
            TabIndex        =   170
            Top             =   120
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.Label Label_Descricao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Programa para gerenciar os filmes do seu computador."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   720
            TabIndex        =   169
            Top             =   360
            Visible         =   0   'False
            Width           =   4080
         End
         Begin VB.Label Label_Downloads 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   168
            Top             =   0
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label_Programa 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   167
            Top             =   0
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label_Empresa 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   166
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label_Nenum_Resultado 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nenhum programa foi encontrado com esta categoria."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            TabIndex        =   78
            Top             =   0
            Visible         =   0   'False
            Width           =   4650
         End
         Begin VB.Label Pic_Linha 
            Height          =   1095
            Index           =   0
            Left            =   0
            TabIndex        =   177
            Top             =   0
            Visible         =   0   'False
            Width           =   10575
         End
      End
      Begin VB.PictureBox Frame_Pesquisar 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFB000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   240
         ScaleHeight     =   249
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   697
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   0
         Width           =   10455
         Begin VB.Label Label_Pesquisa 
            BackColor       =   &H00FFC0FF&
            Height          =   255
            Left            =   3600
            TabIndex        =   156
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label_Titulo_Frame_Programas 
            Alignment       =   2  'Center
            BackColor       =   &H00CBB534&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   38
            Top             =   2400
            Width           =   2265
         End
         Begin VB.Label Label_Titulo_Frame_Programas 
            AutoSize        =   -1  'True
            BackColor       =   &H00CBB534&
            BackStyle       =   0  'Transparent
            Caption         =   "SELECIONE A CATEGORIA DO PROGRAMA QUE PRETENDE PESQUISAR"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   0
            TabIndex        =   37
            Top             =   600
            Width           =   6480
         End
         Begin VB.Label Label_Titulo_Frame_Programas 
            AutoSize        =   -1  'True
            BackColor       =   &H00C4AD2F&
            BackStyle       =   0  'Transparent
            Caption         =   "PESQUISAR"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   570
            Index           =   0
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   3165
         End
         Begin VB.Image Icon_Pasta_Categoria 
            Height          =   1545
            Index           =   3
            Left            =   6360
            Picture         =   "Form_Principal.frx":1A095
            Top             =   840
            Width           =   1920
         End
         Begin VB.Image Icon_Pasta_Categoria 
            Height          =   1545
            Index           =   4
            Left            =   8520
            Picture         =   "Form_Principal.frx":1B53B
            Top             =   840
            Width           =   1920
         End
         Begin VB.Image Icon_Pasta_Categoria 
            Height          =   1545
            Index           =   0
            Left            =   0
            Picture         =   "Form_Principal.frx":1C6CB
            Top             =   840
            Width           =   1920
         End
         Begin VB.Image Icon_Pasta_Categoria 
            Height          =   1545
            Index           =   1
            Left            =   2160
            Picture         =   "Form_Principal.frx":1DA69
            Top             =   840
            Width           =   1920
         End
         Begin VB.Image Icon_Pasta_Categoria 
            Height          =   1545
            Index           =   2
            Left            =   4200
            Picture         =   "Form_Principal.frx":1EBC6
            Top             =   840
            Width           =   1920
         End
      End
   End
   Begin VB.PictureBox Barra_Ferramentas 
      Appearance      =   0  'Flat
      BackColor       =   &H00212121&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   0
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   465
      Width           =   9600
      Begin VB.PictureBox Menu_Sobre 
         Appearance      =   0  'Flat
         BackColor       =   &H002B2B2B&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   6075
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   80
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   0
         Width           =   1200
         Begin VB.Label Label_Menu_Sobre 
            Alignment       =   2  'Center
            BackColor       =   &H00161616&
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   42
            Top             =   750
            Width           =   1200
         End
      End
      Begin VB.PictureBox Menu_Partilhar 
         Appearance      =   0  'Flat
         BackColor       =   &H002B2B2B&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   2430
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   80
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   0
         Width           =   1200
         Begin VB.Label Label_Menu_Partilhar 
            Alignment       =   2  'Center
            BackColor       =   &H00161616&
            BackStyle       =   0  'Transparent
            Caption         =   "Partilhar"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   34
            Top             =   750
            Width           =   1200
         End
      End
      Begin VB.PictureBox Menu_Opcoes 
         Appearance      =   0  'Flat
         BackColor       =   &H002B2B2B&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   3645
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   80
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   0
         Width           =   1200
         Begin VB.Label Label_Menu_Opcoes 
            Alignment       =   2  'Center
            BackColor       =   &H00161616&
            BackStyle       =   0  'Transparent
            Caption         =   "Opções"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   29
            Top             =   750
            Width           =   1200
         End
      End
      Begin VB.PictureBox Menu_Suporte 
         Appearance      =   0  'Flat
         BackColor       =   &H002B2B2B&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   4860
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   80
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   0
         Width           =   1200
         Begin VB.Label Label_Menu_Suporte 
            Alignment       =   2  'Center
            BackColor       =   &H00161616&
            BackStyle       =   0  'Transparent
            Caption         =   "Suporte"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   27
            Top             =   750
            Width           =   1200
         End
      End
      Begin VB.PictureBox Menu_Instalados 
         Appearance      =   0  'Flat
         BackColor       =   &H002B2B2B&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   1215
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   80
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   0
         Width           =   1200
         Begin VB.Label Label_Menu_Instalados 
            Alignment       =   2  'Center
            BackColor       =   &H00161616&
            BackStyle       =   0  'Transparent
            Caption         =   "Instalados"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   25
            Top             =   750
            Width           =   1200
         End
      End
      Begin VB.PictureBox Menu_Pesquisar 
         Appearance      =   0  'Flat
         BackColor       =   &H002B2B2B&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   0
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   80
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Width           =   1200
         Begin VB.Label Label_Menu_Pesquisar 
            Alignment       =   2  'Center
            BackColor       =   &H00161616&
            BackStyle       =   0  'Transparent
            Caption         =   "Pesquisar"
            ForeColor       =   &H00DFB000&
            Height          =   195
            Left            =   0
            TabIndex        =   23
            Top             =   750
            Width           =   1200
         End
      End
      Begin VB.Image Fundo_Barra_Ferramentas 
         Enabled         =   0   'False
         Height          =   1050
         Left            =   8160
         Picture         =   "Form_Principal.frx":1FF10
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox Barra_Botoes 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   0
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   736
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1440
      Width           =   11040
      Begin VB.PictureBox Botao_Carregar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   5040
         Picture         =   "Form_Principal.frx":20271
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   1065
         Begin VB.Label Label_Carregar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   120
            Width           =   780
         End
         Begin VB.Image Extermidade_Carregar 
            Enabled         =   0   'False
            Height          =   405
            Left            =   840
            Picture         =   "Form_Principal.frx":24AD7
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox Botao_Desinstalar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   3600
         Picture         =   "Form_Principal.frx":24E18
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   87
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   1305
         Begin VB.Label Label_Desinstalar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desinstalar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   74
            Top             =   120
            Width           =   960
         End
         Begin VB.Image Extermidade_Desinstalar 
            Enabled         =   0   'False
            Height          =   405
            Left            =   1080
            Picture         =   "Form_Principal.frx":2967E
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox Botao_Run 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   2520
         Picture         =   "Form_Principal.frx":299BF
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   63
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   945
         Begin VB.Label Label_Run 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Executar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   120
            Width           =   750
         End
         Begin VB.Image Extermidade_Run 
            Enabled         =   0   'False
            Height          =   405
            Left            =   720
            Picture         =   "Form_Principal.frx":2E225
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox Botao_Aplicar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1440
         Picture         =   "Form_Principal.frx":2E566
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   63
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   945
         Begin VB.Label Label_Aplicar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aplicar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   120
            Width           =   585
         End
         Begin VB.Image Extermidade_Aplicar 
            Enabled         =   0   'False
            Height          =   405
            Left            =   720
            Picture         =   "Form_Principal.frx":32DCC
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox Botao_Website 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   0
         Picture         =   "Form_Principal.frx":3310D
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   84
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   1260
         Begin VB.Label Label_Website 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Site oficial"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   120
            Width           =   885
         End
         Begin VB.Image Extermidade_Website 
            Enabled         =   0   'False
            Height          =   405
            Left            =   960
            Picture         =   "Form_Principal.frx":37973
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox Botao_Enviar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   5040
         Picture         =   "Form_Principal.frx":37CB4
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   68
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1020
         Begin VB.Label Label_Enviar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enviar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Width           =   540
         End
         Begin VB.Image Extermidade_Enviar 
            Enabled         =   0   'False
            Height          =   405
            Left            =   720
            Picture         =   "Form_Principal.frx":3C51A
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox Botao_Limpar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   6240
         Picture         =   "Form_Principal.frx":3C85B
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   68
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1020
         Begin VB.Label Label_Limpar 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limpar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Width           =   600
         End
         Begin VB.Image Extermidade_Limpar 
            Enabled         =   0   'False
            Height          =   405
            Left            =   720
            Picture         =   "Form_Principal.frx":410C1
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox Barra_Caixa_Pesquisa 
         Appearance      =   0  'Flat
         BackColor       =   &H00FEFEFE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   7680
         Picture         =   "Form_Principal.frx":41402
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   120
         Width           =   2280
         Begin VB.TextBox Text_Pesquisa 
            Appearance      =   0  'Flat
            BackColor       =   &H00FEFEFE&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   120
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   "Pesquisar"
            Top             =   90
            Width           =   855
         End
         Begin VB.Shape Contorno_Caixa_Pesquisa 
            BorderColor     =   &H00C0C0C0&
            Height          =   375
            Left            =   15
            Shape           =   4  'Rounded Rectangle
            Top             =   15
            Width           =   750
         End
         Begin VB.Image Botao_Pesquisar 
            Height          =   315
            Left            =   1800
            Picture         =   "Form_Principal.frx":436C4
            Top             =   0
            Width           =   285
         End
      End
      Begin VB.PictureBox Separador_Informacoes 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   3240
         Picture         =   "Form_Principal.frx":43BF2
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   111
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1665
         Begin VB.Label Label_Informacoes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Informações"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   105
            Width           =   1080
         End
         Begin VB.Image Extermidade_Informacoes 
            Enabled         =   0   'False
            Height          =   405
            Left            =   1440
            Picture         =   "Form_Principal.frx":4E710
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox Separador_Categorias 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1560
         Picture         =   "Form_Principal.frx":4EA51
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   111
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1665
         Begin VB.Label Label_Categorias 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Categoria"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   105
            Width           =   840
         End
         Begin VB.Image Extermidade_Categorias 
            Enabled         =   0   'False
            Height          =   405
            Left            =   1440
            Picture         =   "Form_Principal.frx":5956F
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox Separador_Programas 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   120
         Picture         =   "Form_Principal.frx":598B0
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   95
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   120
         Width           =   1425
         Begin VB.Label Label_Programas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Programas"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   105
            Width           =   930
         End
         Begin VB.Image Extermidade_Programas 
            Enabled         =   0   'False
            Height          =   405
            Left            =   1200
            Picture         =   "Form_Principal.frx":5A0F1
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.Image Fundo_Barra_Botoes 
         Enabled         =   0   'False
         Height          =   585
         Left            =   0
         Picture         =   "Form_Principal.frx":5A432
         Top             =   0
         Width           =   1845
      End
   End
   Begin VB.PictureBox Barra_ControlBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00313131&
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
      Height          =   465
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   9600
      Begin VB.TextBox Text_Form_Width 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5640
         TabIndex        =   164
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text_Form_Left 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5040
         TabIndex        =   163
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text_Form_Height 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4440
         TabIndex        =   162
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text_Form_Top 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3840
         TabIndex        =   161
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox pichook 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2760
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   158
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox Text_Tela_Cheia 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3240
         TabIndex        =   83
         TabStop         =   0   'False
         Text            =   "True"
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image Botao_Minimizar 
         Height          =   225
         Left            =   8160
         ToolTipText     =   "Minimizar"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Botao_Restaurar 
         Height          =   225
         Left            =   8520
         ToolTipText     =   "Restaurar"
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Botao_Maximizar 
         Height          =   225
         Left            =   8880
         ToolTipText     =   "Maximizar"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Botao_Fechar 
         Height          =   225
         Left            =   9240
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   240
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   2325
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   360
         Left            =   0
         Picture         =   "Form_Principal.frx":5A80C
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.PictureBox Barra_Detalhes 
      Appearance      =   0  'Flat
      BackColor       =   &H002E2E2E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   737
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2640
      Width           =   11055
      Begin VB.PictureBox Botao_Update 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7320
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   168
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   2520
         Begin VB.Image Icon_Update 
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label_Update 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Actualizar programa"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   360
            TabIndex        =   40
            Top             =   30
            Width           =   1740
         End
         Begin VB.Image Image_Update 
            Enabled         =   0   'False
            Height          =   255
            Left            =   2160
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Image Botao_Redimensionar 
         Height          =   150
         Left            =   10800
         Picture         =   "Form_Principal.frx":5AAB5
         Top             =   120
         Width           =   150
      End
      Begin VB.Label Label_Utilizador_Logado 
         AutoSize        =   -1  'True
         BackColor       =   &H002B2B2B&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2011-2012 Nikyts software"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   3420
      End
      Begin VB.Image Fundo_Barra_Detalhes 
         Enabled         =   0   'False
         Height          =   435
         Left            =   0
         Picture         =   "Form_Principal.frx":5AC37
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.PictureBox Frame_Sobre 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   13440
      ScaleHeight     =   185
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox Text_About 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "Form_Principal.frx":5AFCB
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label_Frame_Sobre 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Sobre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   720
         TabIndex        =   48
         Top             =   360
         Width           =   930
      End
   End
   Begin VB.PictureBox Frame_Suporte 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   5415
      Left            =   11640
      ScaleHeight     =   361
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   848
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   12720
      Begin VB.PictureBox Frame_Erro_Suporte 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1125
         Left            =   6840
         Picture         =   "Form_Principal.frx":5AFD1
         ScaleHeight     =   75
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   334
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   5010
         Begin VB.Label Label_Erro_Suporte 
            AutoSize        =   -1  'True
            BackColor       =   &H00EEEEEE&
            BackStyle       =   0  'Transparent
            Caption         =   "Indique um endereço de email válido."
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   480
            TabIndex        =   94
            Top             =   420
            Width           =   3255
         End
         Begin VB.Label Label_Close_Suporte 
            AutoSize        =   -1  'True
            BackColor       =   &H00F5F5F5&
            BackStyle       =   0  'Transparent
            Caption         =   " x "
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   4680
            TabIndex        =   93
            Top             =   120
            Width           =   225
         End
      End
      Begin VB.PictureBox Lista_Assunto 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   1680
         ScaleHeight     =   63
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   303
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   2520
         Visible         =   0   'False
         Width           =   4575
         Begin VB.Label Label_Assunto 
            BackColor       =   &H00EEEEEE&
            BackStyle       =   0  'Transparent
            Caption         =   "Assunto"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   91
            Top             =   0
            Width           =   960
         End
         Begin VB.Label Shape_Sombra 
            BackColor       =   &H00DFB000&
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   90
            Top             =   0
            Width           =   3975
         End
      End
      Begin VB.PictureBox Barra_Text_Email 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   720
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   1260
         Width           =   5475
         Begin VB.TextBox Text_Email 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   600
            TabIndex        =   14
            Top             =   30
            Width           =   1380
         End
         Begin VB.Shape Contorno_Email 
            BorderColor     =   &H00C0C0C0&
            Height          =   375
            Left            =   0
            Top             =   0
            Width           =   4695
         End
      End
      Begin VB.PictureBox Barra_Text_Assunto 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   720
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   2100
         Width           =   5475
         Begin VB.PictureBox Seta_Assunto 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   5040
            Picture         =   "Form_Principal.frx":6D637
            ScaleHeight     =   19
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   19
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   0
            Width           =   285
         End
         Begin VB.TextBox Text_Assunto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   30
            Width           =   1380
         End
         Begin VB.Shape Contorno_Assunto 
            BorderColor     =   &H00C0C0C0&
            Height          =   375
            Left            =   0
            Top             =   0
            Width           =   4695
         End
      End
      Begin VB.PictureBox Barra_Text_Mensagem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   1920
         Left            =   720
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   2940
         Width           =   5475
         Begin VB.TextBox Text_Mensagem 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   1740
            Left            =   600
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   30
            Width           =   1860
         End
         Begin VB.Shape Contorno_Mensagem 
            BorderColor     =   &H00C0C0C0&
            Height          =   1800
            Left            =   0
            Top             =   0
            Width           =   4695
         End
      End
      Begin VB.Label Label_Frame_Suporte 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Suporte técnico"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   720
         TabIndex        =   99
         Top             =   360
         Width           =   2475
      End
      Begin VB.Label Label_Info 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(Para possivel contacto caso seja necessário)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   180
         Left            =   1320
         TabIndex        =   98
         Top             =   1020
         Width           =   3480
      End
      Begin VB.Label Label_De 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   97
         Top             =   1020
         Width           =   465
      End
      Begin VB.Label Label_Texto 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Assunto"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   96
         Top             =   1860
         Width           =   675
      End
      Begin VB.Label Label_Mensagem 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagem"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   95
         Top             =   2700
         Width           =   915
      End
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00808080&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form_Principal"
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

'Declaração das variáveis
'Dim bMoveFrom As Boolean, LastPoint As POINTAPI

'API para abrir web
Private Const SW_NORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Variável para identificar que categoria do programa deve ser procurada
Dim Categoria_a_ser_Pesquisada As String

'Cores utilizadas pelo programa
Const Azul = &HDFB000
Const Cinza = &HC0C0C0
Const Laranja = &H80FF&

'Variável para o progressbar
Private m_cProgress As Collection

'Ajusta o Form para sempre exibir a barra de tarefas do windows, full screen
Private Const SPI_GETWORKAREA = 48
Private Type RECT
  left As Long
  top As Long
  Right As Long
  Bottom As Long
End Type
Private Declare Function SystemParametersInfo Lib "user32" _
Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Variavel para verificar a janela do formulário
Dim Tela_Cheia As Boolean
    
'Api's para actualizar a date e hora dos programas após o seu download
Private Type FILETIME
    LowDateTime As Long
    HighDateTime As Long
End Type
Private Type SYSTEMTIME
    Year As Integer
    Month As Integer
    DayOfWeek As Integer
    Day As Integer
    Hour As Integer
    Minute As Integer
    Second As Integer
    Milliseconds As Integer
End Type
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As Any) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

'Variável para identificar qual foi a linha selecionada da lista de programas
Dim Linha_Selecionada As Integer
Dim Linha_Selecionada_Assunto As Integer
Dim Linha_Selecionada_Lingua As Integer

'Variáveis do idioma
Dim Idioma_Erro As String
Dim Idioma_Descricao As String
Dim Idioma_Erro_Execucao As String
Dim Idioma_Conectar_Servidor As String
Dim Idioma_Internet_Desligada As String

Dim Idioma_Button_Transfer_Program As String
Dim Idioma_Button_Execute_Program As String
Dim Idioma_Button_Remove_Program As String
Dim Idioma_Button_Cancel_Program As String
Dim Idioma_Label_Rate As String
Dim Idioma_Transferring_File As String

'Variável para saber qual é o progressbar activo da lista de programas
Dim progress_activo As Integer

'Variável para identificar qual foi a linha selecionada da lista de programas
Dim Linha_Programa_Selecionado As Integer

'Variável para identificar que separadores estão visiveis durante a pesquisa
Dim Tab_Categoria_Visivel As Boolean
Dim Tab_Informacao_Visivel As Boolean

Dim texto_a_pesquisar As String

'Variável para saber que programa esta selecionado
Dim programa_selecionado As Integer

'tray icon
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim t As NOTIFYICONDATA
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_FINDSTRING = &H18F

'Redimensionar formulário
'Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const WM_NCLBUTTONDOWN = &HA1
Const HTBOTTOMRIGHT = 17
Const HTCAPTION = 2

'Variáveis para definir as dimensões minimas do form
Dim Altura_Standard As Integer
Dim Largura_Standard As Integer

Function FileSetDate(ByVal sFileName As String, ByVal dFileDate As Date, Optional bSetCreationTime As Boolean = False, Optional bSetLastAccessedTime As Boolean = False, Optional bSetLastModified As Boolean = False) As Boolean
    'Função para actualizar a data e hora dos programas após o seu download
    Const GENERIC_WRITE = &H40000000, OPEN_EXISTING = 3
    Const FILE_SHARE_READ = &H1, FILE_SHARE_WRITE = &H2
    
    Dim lhwndFile As Long
    Dim tSystemTime As SYSTEMTIME
    Dim tLocalTime As FILETIME, tFileTime As FILETIME
    
    tSystemTime.Year = Year(dFileDate)
    tSystemTime.Month = Month(dFileDate)
    tSystemTime.Day = Day(dFileDate)
    tSystemTime.DayOfWeek = Weekday(dFileDate) - 1
    tSystemTime.Hour = Hour(dFileDate)
    tSystemTime.Second = Second(dFileDate)
    tSystemTime.Milliseconds = 0

    'Open the file to get the filehandle
    lhwndFile = CreateFile(sFileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    If lhwndFile Then
        'File opened
        'Convert system time to local time
        SystemTimeToFileTime tSystemTime, tLocalTime
        'Convert local time to GMT
        LocalFileTimeToFileTime tLocalTime, tFileTime
'-------Change date/time property of the file
        FileSetDate = True
        If bSetCreationTime Then
            FileSetDate = FileSetDate And CBool(SetFileTime(lhwndFile, tFileTime, 0&, 0&))
        End If
        If bSetLastAccessedTime Then
            FileSetDate = FileSetDate And CBool(SetFileTime(lhwndFile, 0&, tFileTime, 0&))
        End If
        If bSetLastModified Then
            FileSetDate = FileSetDate And CBool(SetFileTime(lhwndFile, 0&, 0&, tFileTime))
        End If
        'Close the file handle
        Call CloseHandle(lhwndFile)
    End If
End Function

Public Function PosFormRelativeTaskBar(F As Form)
    'Função para ao maximizar o form seja visivel a barra do windows iniciar
    'Colocar o WindowsState=0 normal
    Dim WindowRect As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    
    SetWindowPos hwnd, 0, WindowRect.left, WindowRect.top, WindowRect.Right - WindowRect.left, WindowRect.Bottom - WindowRect.top, 0
    F.top = WindowRect.Bottom * Screen.TwipsPerPixelY - F.Height
    F.left = WindowRect.Right * Screen.TwipsPerPixelX - F.Width
End Function

Private Sub Barra_Botoes_Click()
    'Chamar o procedimento
    Ocultar_Listas
End Sub

Private Sub Barra_Botoes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Barra_Detalhes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Barra_Estado_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Barra_Ferramentas_Click()
    'Chamar o procedimento
    Ocultar_Listas
End Sub

Private Sub Barra_Ferramentas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Principal
End Sub

Private Sub Barra_Ferramentas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Repor_Objectos
    If Tela_Cheia = False Then Mover_Formulario Form_Principal
End Sub

Private Sub Barra_Ferramentas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Principal
    Actualizar_Valores
End Sub

Private Sub Barra_Ferramentas_DblClick()
    'Atalho para
    Label_Titulo_DblClick
End Sub

Private Sub Barra_Transferir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Botao_Aplicar_Click()
    'Atalho para
    Label_Aplicar_Click
End Sub

Private Sub Botao_Cancelar_Click()
    'Atalho para
    Label_Cancelar_Click
End Sub

Private Sub Botao_Carregar_Click()
    'Atalho para
    Label_Carregar_Click
End Sub

Private Sub Botao_Desinstalar_Click()
    'Atalho para
    Label_Desinstalar_Click
End Sub

Private Sub Botao_Download_Click()
    'Atalho para
    Label_Download_Click
End Sub

Private Sub Botao_Enviar_Click()
    'Atalho para
    Label_Enviar_Click
End Sub

Private Sub Botao_Eu_Gosto_Click()
    'Atalho para
    Label_Eu_Gosto_Click
End Sub

Private Sub Botao_Executar_Click()
    'Atalho para
    Label_Executar_Click
End Sub

Private Sub Botao_Executar_Programa_Click(Index As Integer)
    'Atalho para
    Label_Executar_Programa_Click (Index)
End Sub

Public Sub Botao_Fechar_Click()
    'Actualizar as opções do programa
    Call WriteINI("Settings", "FullScreen", Text_Tela_Cheia.Text, (Localizacao_Ficheiro_Preferencias))
    Call WriteINI("Settings", "Form_Top", Text_Form_Top, (Localizacao_Ficheiro_Preferencias))
    Call WriteINI("Settings", "Form_Height", Text_Form_Height, (Localizacao_Ficheiro_Preferencias))
    Call WriteINI("Settings", "Form_Left", Text_Form_Left, (Localizacao_Ficheiro_Preferencias))
    Call WriteINI("Settings", "Form_width", Text_Form_Width, (Localizacao_Ficheiro_Preferencias))
    
    'Fechar/ ocultar o formulário
    Unload Form_Barra
    Unload Me
    End
End Sub

Private Sub Botao_Tray_Click()
    'Mensagem no icon do projecto/ coloca-lo ao lado do clock
    t.cbSize = Len(t)
    t.hwnd = pichook.hwnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = Me.Icon
    t.szTip = App.ProductName & Chr$(10)  'Texto a ser exibido no icon
    Shell_NotifyIcon NIM_ADD, t
    App.TaskVisible = False
    
    'Colocar o icon do formulário ao lado do clock do windows
    Me.Hide
End Sub

Private Sub Botao_Limpar_Click()
    'Atalho para
    Label_Limpar_Click
End Sub

Private Sub Botao_Pesquisar_Click()
    'Pesquisar programa pelo nome
    On Error GoTo Corrige_Erro
    If Len(Trim(Text_Pesquisa.Text)) = 0 Then Exit Sub
    If Text_Pesquisa.Text = ReadINI("Main", "Text_Search", Localizacao_Ficheiro_Lingua) Then Exit Sub
    texto_a_pesquisar = Text_Pesquisa.Text
    
    'Ocultar objectos não desejados para a operação actual
    Separador_Informacoes.Visible = False
    Tab_Informacao_Visivel = False
    Barra_Estado.Visible = False
    Ocultar_Objectos
    Formatar_Lista_Programas
    Repor_Altura_das_Linhas
    
    Label_Categorias.Caption = ReadINI("Main", "Tab_Result_Search", Localizacao_Ficheiro_Lingua) & ": " & Text_Pesquisa.Text
    With Separador_Categorias
        .Height = Form_Skin.Fundo_Separadores.Height
        .left = Separador_Programas.left + Separador_Programas.ScaleWidth
        .Width = Label_Categorias.Width + 8 + 15
        Extermidade_Categorias.Picture = Form_Skin.Extermidade_Normal.Picture
        Extermidade_Categorias.left = .ScaleWidth - Extermidade_Categorias.Width
        .Visible = True
    End With
    Extermidade_Programas.Picture = Form_Skin.Extermidade_Over.Picture
    Tab_Categoria_Visivel = True
            
    'Enviar o pedido da pesquisa para o servidor
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.open "GET", "http://www.nikyts.com/gadgets/" & "pesquisarprograma.asp?Recebe_Pesquisa=" & Text_Pesquisa.Text
    servidor.send
    
    'Verificar os dados acesso
    If servidor.responseText = "NaoExiste" Then
        Label_Nenum_Resultado.Visible = True
        'Exit Sub
        
    ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 Then
            Me.MousePointer = 11
        
            Formatar_Lista_Programas
            Ajustar_Linha_Lista_Programas
            Extermidade_Programas.Picture = Form_Skin.Extermidade_Over.Picture
        
            Formatar_Lista_Programas
            Servidor_Carregar_Programas servidor.responseText
            Me.MousePointer = 0
            Frame_Lista.Visible = True
        End If
    End If
    Set servidor = Nothing
    Ocultar_Objectos
    Frame_Lista.Visible = True
    Text_Pesquisa.Text = ReadINI("Main", "Text_Search", Localizacao_Ficheiro_Lingua)
    Text_Pesquisa.ForeColor = &H808080
    Contorno_Caixa_Pesquisa.BorderColor = Cinza
    Label_Pesquisa.Caption = "Programa"
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Ocultar_Objectos
Frame_Lista.Visible = True
Label_Nenum_Resultado.Visible = True
Text_Pesquisa.Text = ReadINI("Main", "Text_Search", Localizacao_Ficheiro_Lingua)
Text_Pesquisa.ForeColor = &H808080
Contorno_Caixa_Pesquisa.BorderColor = Cinza
    
Select Case err.Number
    Case -2146697211
        Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada
        
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub Botao_Redimensionar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Redimensionar o formulário conforme as dimensões pretendidas
    If Button = vbLeftButton Then
        If Tela_Cheia = False Then
            ReleaseCapture
            SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
            
            'Verificar se não exedeu os limites
            If Me.Height < Altura_Standard Then
                Me.Height = Altura_Standard
            End If
        
            If Me.Width < Largura_Standard Then
                Me.Width = Largura_Standard
            End If
            Actualizar_Valores
            Desenhar_Formulario
        End If
    End If
End Sub

Private Sub Botao_Redimensionar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Alterar o mousepointer
    If Tela_Cheia = True Then
        Botao_Redimensionar.MousePointer = vbDefault
    Else
        Botao_Redimensionar.MousePointer = 8
    End If
End Sub

Private Sub Botao_Remover_Transferencia_Click(Index As Integer)
    'Atalho para
    Label_Remover_Transferencia_Click (Index)
End Sub

Private Sub Botao_Mais_Informacoes_Click(Index As Integer)
    'Atalho para
    Label_Mais_Informacoes_Click (Index)
End Sub

Private Sub Botao_Maximizar_Click()
    'Maximixar formulário
    PosFormRelativeTaskBar Me
    Tela_Cheia = True
    Text_Tela_Cheia.Text = "True"
'    form_preferencias.Actualizar_Opcoes
    Botao_Maximizar.Visible = False
    Botao_Restaurar.Visible = True
End Sub

Private Sub Botao_Minimizar_Click()
    'Minimizar o formulário
    If Check_Tray.Value = 1 Then
        Botao_Tray_Click
    Else
        Me.WindowState = 1
    End If
End Sub

Private Sub Botao_Restaurar_Click()
    'Restaurar janela
    With Me
        .top = Text_Form_Top
        .Height = Text_Form_Height
        .left = Text_Form_Left
        .Width = Text_Form_Width
    End With
    Tela_Cheia = False
    Text_Tela_Cheia.Text = "False"
'    form_preferencias.Actualizar_Opcoes
    Botao_Maximizar.Visible = True
    Botao_Restaurar.Visible = False
    
    Actualizar_Valores
End Sub

Private Sub Botao_Run_Click()
    'Atalho para
    Label_Run_Click
End Sub

Private Sub Botao_Update_Click()
    'Atalho para
    Label_Update_Click
End Sub

Private Sub Botao_Update_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mostar a imagem down
    Image_Update.Picture = Form_Skin.Button_Menu_Down.Picture
    Icon_Update.Picture = Form_Skin.Icon_Menu_Down.Picture
End Sub

Private Sub Botao_Update_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Legendas on-line
    Image_Update.Picture = Form_Skin.Button_Menu_Normal.Picture
    Icon_Update.Picture = Form_Skin.Icon_Menu_Normal.Picture
End Sub

Private Sub Botao_Website_Click()
    'Atalho para
    Label_Website_Click
End Sub

Private Sub Check_Barra_Click()
    'Des/Activar a opcção
    If Check_Barra.Value = 1 Then
        Pic_Barra.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Pic_Barra.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Check_Tray_Click()
    'Des/Activar a opcção
    If Check_Tray.Value = 1 Then
        Pic_Tray.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Pic_Tray.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Frame_Centro_Click()
    'Chamar o procedimento
    Ocultar_Listas
End Sub

Private Sub Frame_Conteudo_Click()
    'Chamar o procedimento
    Ocultar_Listas
End Sub

Private Sub Frame_Informacoes_Click()
    'Chamar o procedimento
    Ocultar_Listas
End Sub

Private Sub Frame_Informacoes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Repor_Objectos()
    'Procedimento para repor os objectos ao estado original/ MouseLeave
    If Label_Enterprise.ForeColor <> vblack Then Label_Enterprise.ForeColor = vblack
    If Label_Site_Oficial.ForeColor <> vblack Then Label_Site_Oficial.ForeColor = vblack
End Sub

Private Sub Frame_Instalados_Click()
    'Chamar o procedimento
    Ocultar_Listas
End Sub

Private Sub Frame_Lista_Click()
    'Chamar o procedimento
    Ocultar_Listas
End Sub

Private Sub Frame_Opcoes_Click()
    'Chamar o procedimento
    Ocultar_Listas
End Sub

Private Sub Frame_Partilhar_Click()
    'Chamar o procedimento
    Ocultar_Listas
End Sub

Private Sub Frame_Pesquisar_Click()
    'Chamar o procedimento
    Ocultar_Listas
End Sub

Private Sub Frame_Sobre_Click()
    'Chamar o procedimento
    Ocultar_Listas
End Sub

Private Sub Frame_Suporte_Click()
    'Chamar o procedimento
    Ocultar_Listas
End Sub

Private Sub Image_Icon_Grande_Click(Index As Integer)
    'Selecionar o programa
    If Index = programa_selecionado Then Exit Sub
    If programa_selecionado > -1 Then Image_Icon_Grande(programa_selecionado).Picture = Image_Logo_Normal(programa_selecionado).Picture
    Image_Icon_Grande(Index).Picture = Image_Logo_Over(Index).Picture
    programa_selecionado = Index
End Sub

Private Sub Image_Icon_Grande_DblClick(Index As Integer)
    'Executar o programa automaticamente
    On Error GoTo Corrige_Erro
    Shell App.Path & "\Programs\" & Label_Icon_Grande(Index).Caption & "\" & Label_Icon_Grande(Index).Caption & ".exe"
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case -2146697211
        Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada
        
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub Label_Aplicar_Click()
    'Aplicar as opções do programa
    Call WriteINI("Settings", "Check_Update", Check_Actualizar.Value, (Localizacao_Ficheiro_Preferencias))
    
    Call WriteINI("Settings", "Language", Text_Lingua.Text, (Localizacao_Ficheiro_Preferencias))
    
    Call WriteINI("Settings", "Check_Bar_Desktop", Check_Barra.Value, (Localizacao_Ficheiro_Preferencias))
    If Check_Barra.Value = 0 Then Unload Form_Barra Else: Form_Barra.Show
    
    Call WriteINI("Settings", "Check_Tray", Check_Tray.Value, (Localizacao_Ficheiro_Preferencias))
    
    Label_Opcoes_Actualizadas.Visible = True
    Recarregar_Idioma_do_Programa
    'Mensagem_de_Aviso "Information", "As opções foram actualizadas com sucesso."
End Sub

Private Sub Label_Barra_Comentario_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Recarregar_Idioma_do_Programa()
    'Procedimento para alterar o idioma
    Carregar_Idioma
        
    Select Case Label_Pesquisa.Caption
        Case "Ferramentas"
            Label_Categorias.Caption = ReadINI("Main", "Folder_Tools", Localizacao_Ficheiro_Lingua)
            
        Case "Som e video"
            Label_Categorias.Caption = ReadINI("Main", "Folder_Media", Localizacao_Ficheiro_Lingua)
            
        Case "Acessibilidade"
            Label_Categorias.Caption = ReadINI("Main", "Folder_Accessibility", Localizacao_Ficheiro_Lingua)
            
        Case "Internet"
            Label_Categorias.Caption = ReadINI("Main", "Folder_Internet", Localizacao_Ficheiro_Lingua)
            
        Case "Jogos"
            Label_Categorias.Caption = ReadINI("Main", "Folder_Games", Localizacao_Ficheiro_Lingua)
            
        Case "Programa"
            Label_Categorias.Caption = ReadINI("Main", "Tab_Result_Search", Localizacao_Ficheiro_Lingua) & ": " & texto_a_pesquisar
            
        Case "Empresa"
            Label_Categorias.Caption = ReadINI("Main", "Tab_Result_Enterprise", Localizacao_Ficheiro_Lingua) & ": " & Label_Enterprise.Caption
    End Select
    
    'Ajustar objecos que sejam necessários consoante o idioma
    With Separador_Programas
        .Width = Label_Programas.Width + 8 + 15
        .left = 8
        Extermidade_Programas.left = .ScaleWidth - Extermidade_Programas.Width
    End With
    
    With Separador_Categorias
        .Width = Label_Categorias.Width + 8 + 15
        .left = Separador_Programas.left + Separador_Programas.ScaleWidth
        Extermidade_Categorias.left = .ScaleWidth - Extermidade_Categorias.Width
    End With
    
    With Separador_Informacoes
        .left = Separador_Categorias.left + Separador_Categorias.ScaleWidth
    End With
    
    With Botao_Enviar
        .Height = Form_Skin.Fundo_Separadores.Height
        .top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .Width = Label_Enviar.Width + 8 + 8
        .left = 8
        Extermidade_Enviar.left = .ScaleWidth - Extermidade_Enviar.Width
    End With
    
    With Botao_Limpar
        .Width = Label_Limpar.Width + 8 + 8
        .left = Botao_Enviar.left + Botao_Enviar.ScaleWidth + 8
        Extermidade_Limpar.left = .ScaleWidth - Extermidade_Limpar.Width
    End With
    
    With Botao_Website
        .Width = Label_Website.Width + 8 + 8
        .left = 8
        Extermidade_Website.left = .ScaleWidth - Extermidade_Website.Width
    End With
    
    With Botao_Aplicar
        .Width = Label_Aplicar.Width + 8 + 8
        .left = 8
        Extermidade_Aplicar.left = .ScaleWidth - Extermidade_Aplicar.Width
    End With
    
    With Botao_Run
        .Width = Label_Run.Width + 8 + 8
        .left = 8
        Extermidade_Run.left = .ScaleWidth - Extermidade_Run.Width
    End With
    
    With Botao_Desinstalar
        .Width = Label_Desinstalar.Width + 8 + 8
        .left = Botao_Run.left + Botao_Run.ScaleWidth + 8
        Extermidade_Desinstalar.left = .ScaleWidth - Extermidade_Desinstalar.Width
    End With
    
    With Botao_Carregar
        .Width = Label_Carregar.Width + 8 + 8
        .left = 8
        Extermidade_Carregar.left = .ScaleWidth - Extermidade_Carregar.Width
    End With
    
    With Botao_Eu_Gosto
        .Width = Label_Eu_Gosto.Width + 8 + 10
        .left = Barra_Transferir.left
        Extermidade_Eu_Gosto.left = .ScaleWidth - Extermidade_Eu_Gosto.Width
    End With
    
    'Actualizar a avaliação do programa
    If Label_Votos.Caption = "1" Then
        Label_Total.Caption = Idioma_Label_Rate & ": " & Label_Votos.Caption
    Else
        Label_Total.Caption = Idioma_Label_Rate & ": " & Label_Votos.Caption
    End If
    
    If Tab_Informacao_Visivel = True Then Verificar_Se_Programa_Existe
End Sub

Private Sub Label_Carregar_Click()
    'Enviar pedido de partilhar de programa
    'On Error GoTo Corrige_Erro
    Frame_Erro_Partilhar.top = (Barra_Txt_Email.top + (Barra_Txt_Email.ScaleHeight / 2)) - 40
    Frame_Erro_Partilhar.Visible = False
    
    'Verificar o preencimento das textboxs
    If Len(Trim(Txt_Email.Text)) = 0 Then
        Label_Erro_Partilhar.Caption = ReadINI("Main", "Message_Required_Field", Localizacao_Ficheiro_Lingua)
        Frame_Erro_Partilhar.top = (Barra_Txt_Email.top + (Barra_Txt_Email.ScaleHeight / 2)) - 40
        Frame_Erro_Partilhar.Visible = True
        Txt_Email.SetFocus
        Exit Sub
    End If
    
    'Verifica se o campo email está no formato correcto
    If Not IsEmail(Txt_Email.Text) Then
        Label_Erro_Partilhar.Caption = ReadINI("Main", "Message_Email_Invalid", Localizacao_Ficheiro_Lingua)
        Frame_Erro_Partilhar.top = (Barra_Txt_Email.top + (Barra_Txt_Email.ScaleHeight / 2)) - 40
        Frame_Erro_Partilhar.Visible = True
        Txt_Email.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(Txt_Empresa.Text)) = 0 Then
        Label_Erro_Partilhar.Caption = ReadINI("Main", "Message_Required_Field", Localizacao_Ficheiro_Lingua)
        Frame_Erro_Partilhar.top = (Barra_Txt_Empresa.top + (Barra_Txt_Empresa.ScaleHeight / 2)) - 40
        Frame_Erro_Partilhar.Visible = True
        Txt_Empresa.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(Txt_Nome.Text)) = 0 Then
        Label_Erro_Partilhar.Caption = ReadINI("Main", "Message_Required_Field", Localizacao_Ficheiro_Lingua)
        Frame_Erro_Partilhar.top = (Barra_Txt_Nome.top + (Barra_Txt_Nome.ScaleHeight / 2)) - 40
        Frame_Erro_Partilhar.Visible = True
        Txt_Nome.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(Txt_Descricao.Text)) = 0 Then
        Label_Erro_Partilhar.Caption = ReadINI("Main", "Message_Required_Field", Localizacao_Ficheiro_Lingua)
        Frame_Erro_Partilhar.top = (Barra_Txt_Descricao.top + (Barra_Txt_Descricao.ScaleHeight / 2)) - 40
        Frame_Erro_Partilhar.Visible = True
        Txt_Descricao.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(Txt_Informacao.Text)) = 0 Then
        Label_Erro_Partilhar.Caption = ReadINI("Main", "Message_Required_Field", Localizacao_Ficheiro_Lingua)
        Frame_Erro_Partilhar.top = (Barra_Txt_Informacao.top + (Barra_Txt_Informacao.ScaleHeight / 2)) - 40
        Frame_Erro_Partilhar.Visible = True
        Txt_Informacao.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(Txt_Site.Text)) = 0 Then
        Label_Erro_Partilhar.Caption = ReadINI("Main", "Message_Required_Field", Localizacao_Ficheiro_Lingua)
        Frame_Erro_Partilhar.top = (Barra_Txt_Site.top + (Barra_Txt_Site.ScaleHeight / 2)) - 40
        Frame_Erro_Partilhar.Visible = True
        Txt_Site.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(Txt_Download.Text)) = 0 Then
        Label_Erro_Partilhar.Caption = ReadINI("Main", "Message_Required_Field", Localizacao_Ficheiro_Lingua)
        Frame_Erro_Partilhar.top = (Barra_Txt_Download.top + (Barra_Txt_Download.ScaleHeight / 2)) - 40
        Frame_Erro_Partilhar.Visible = True
        Txt_Download.SetFocus
        Exit Sub
    End If
    
    Me.MousePointer = 11
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.open "GET", "http://www.nikyts.com/gadgets/" & "partilhar.asp?Email=" & Txt_Email.Text & "&Empresa=" & Txt_Empresa & "&Nome=" & Txt_Nome.Text & "&Descricao=" & Txt_Descricao.Text & "&Informacao=" & Txt_Informacao.Text & "&Site=" & Txt_Site.Text & "&Download=" & Txt_Download.Text, False
    servidor.send 'envia o pedido para o servidor

    'Verificar os dados acesso
    If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then ' 4 - deu resposta e 200 validou
            Mensagem_de_Aviso "Information", "O seu pedido foi enviado com sucesso." & vbNewLine _
            & "O seu programa terá de passar por um processo de aprovação antes de ser publicado." & vbNewLine _
            & "Assim que o processo for concluido irá receber um email com a confirmação do mesmo." & vbNewLine _
            & "Obrigado pela sua colaboração!" 'ReadINI("Main", "Info_Posted", Localizacao_Ficheiro_Lingua)
            
            Me.MousePointer = 0
            Frame_Erro_Partilhar.Visible = False
            Limpa_Campos_Partilhar
            Txt_Email.SetFocus
        End If
    End If
    
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case -2146697211
        Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada
        
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub Limpa_Campos_Partilhar()
    'Procedimento para limpar os campos da frame partilhar
    Txt_Email.Text = ""
    Txt_Empresa.Text = ""
    Txt_Nome.Text = ""
    Txt_Descricao.Text = ""
    Txt_Informacao.Text = ""
    Txt_Site.Text = ""
    Txt_Download.Text = ""
End Sub

Private Sub Label_Close_Partilhar_Click()
    'Ocultar frame erro
    Frame_Erro_Partilhar.Visible = False
End Sub

Private Sub Label_Descricao_Programa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Label_Desinstalar_Click()
    'Desinstalar o programa selecionado, remover a pasta, sub-pasta e respectivos ficheiros referentes ao programa
    On Error GoTo Corrige_Erro
    If programa_selecionado = "-1" Then Mensagem_de_Aviso "Information", ReadINI("Main", "Info_None_Selected_Program", Localizacao_Ficheiro_Lingua): Exit Sub
    
    Mensagem_de_Aviso "Question", ReadINI("Main", "Quest_Uninstall_Programa", Localizacao_Ficheiro_Lingua) & vbNewLine & Label_Icon_Grande(programa_selecionado).Caption
    If Resposta = "Yes" Then
        'Remove a pasta
        Me.MousePointer = 11
        DeleteFolderTree App.Path & "\Programs\" & Label_Icon_Grande(programa_selecionado).Caption
        Recarregar_Programas_Instalados
    End If
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
End Sub

Private Sub Recarregar_Programas_Instalados()
    'Faz refresh à frame dos programas instalados
    programa_selecionado = -1
    Image_Icon_Grande(0).Visible = False: Image_Icon_Grande(0).Picture = LoadPicture("")
    Label_Icon_Grande(0).Visible = False: Label_Icon_Grande(0).Caption = ""
    If Image_Icon_Grande.Count > 1 Then
        Dim Objecto As Integer: For Objecto = 1 To Image_Icon_Grande.Count - 1
            Unload Image_Icon_Grande(Objecto)
            Unload Label_Icon_Grande(Objecto)
            Unload Image_Logo_Normal(Objecto)
            Unload Image_Logo_Over(Objecto)
        Next
    End If
    Verificar_Programas_Existentes (App.Path & "\Programs\")
    Carregar_Programas_Existentes
    
    Unload Form_Barra
    Form_Barra.Show
    
    Me.MousePointer = 0
End Sub

Private Sub Label_Enterprise_Click()
    'Pesquisar programa pelo nome
    On Error GoTo Corrige_Erro
    
    'Ocultar objectos não desejados para a operação actual
    Separador_Informacoes.Visible = False
    Tab_Informacao_Visivel = False
    Barra_Estado.Visible = False
    Repor_Altura_das_Linhas
    Ocultar_Objectos
    Formatar_Lista_Programas
    
    Label_Categorias.Caption = ReadINI("Main", "Tab_Result_Enterprise", Localizacao_Ficheiro_Lingua) & ": " & Label_Enterprise.Caption
    With Separador_Categorias
        .Height = Form_Skin.Fundo_Separadores.Height
        .left = Separador_Programas.left + Separador_Programas.ScaleWidth
        .Width = Label_Categorias.Width + 8 + 15
        Extermidade_Categorias.Picture = Form_Skin.Extermidade_Normal.Picture
        Extermidade_Categorias.left = .ScaleWidth - Extermidade_Categorias.Width
        .Visible = True
    End With
    Extermidade_Programas.Picture = Form_Skin.Extermidade_Over.Picture
    Tab_Categoria_Visivel = True
            
    'Enviar o pedido da pesquisa para o servidor
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.open "GET", "http://www.nikyts.com/gadgets/" & "pesquisarempresa.asp?Recebe_Pesquisa=" & Label_Enterprise.Caption
    servidor.send
    
    'Verificar os dados acesso
    If servidor.responseText = "NaoExiste" Then
        Label_Nenum_Resultado.Visible = True
        'Exit Sub
        
    ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 Then
            Me.MousePointer = 11
        
            Formatar_Lista_Programas
            Ajustar_Linha_Lista_Programas
            Extermidade_Programas.Picture = Form_Skin.Extermidade_Over.Picture
        
            Formatar_Lista_Programas
            Servidor_Carregar_Programas servidor.responseText
            Me.MousePointer = 0
            Frame_Lista.Visible = True
        End If
    End If
    Set servidor = Nothing
    Ocultar_Objectos
    Frame_Lista.Visible = True
    Text_Pesquisa.Text = ReadINI("Main", "Text_Search", Localizacao_Ficheiro_Lingua)
    Text_Pesquisa.ForeColor = &H808080
    Contorno_Caixa_Pesquisa.BorderColor = Cinza
    Label_Pesquisa.Caption = "Empresa"
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Ocultar_Objectos
Frame_Lista.Visible = True
Label_Nenum_Resultado.Visible = True
Text_Pesquisa.Text = ReadINI("Main", "Text_Search", Localizacao_Ficheiro_Lingua)
Text_Pesquisa.ForeColor = &H808080
Contorno_Caixa_Pesquisa.BorderColor = Cinza
Label_Pesquisa.Caption = "Empresa"
    
Select Case err.Number
    Case -2146697211
        Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada
        
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub Label_Enterprise_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar a label ao passar o rato
    Label_Enterprise.ForeColor = Laranja
End Sub

Private Sub Label_Icon_Grande_DblClick(Index As Integer)
    'Atalho para
    Image_Icon_Grande_DblClick Index
End Sub

Private Sub Label_Lingua_Click(Index As Integer)
    'Indicar a lingua selecionada pelo utilizador
    Text_Lingua.Text = Label_Lingua(Index).Caption
    Lista_Linguas.Visible = False
    Text_Lingua.SetFocus
End Sub

Private Sub Label_Lingua_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Lingua = Index Then Exit Sub
    Shape_Sombra_Lingua(Linha_Selecionada_Lingua).Visible = False
    Label_Lingua(Linha_Selecionada_Lingua).ForeColor = Form_Skin.Cor_Letra_Textbox.backcolor
    Shape_Sombra_Lingua(Index).Visible = True
    Label_Lingua(Index).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
    Linha_Selecionada_Lingua = Index
End Sub

Private Sub Label_Nome_Programa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Label_Run_Click()
    'Executar o programa selecionado
    On Error GoTo Corrige_Erro
    If programa_selecionado = "-1" Then Mensagem_de_Aviso "Information", ReadINI("Main", "Info_None_Selected_Program", Localizacao_Ficheiro_Lingua): Exit Sub
    
    Shell App.Path & "\Programs\" & Label_Icon_Grande(programa_selecionado).Caption & "\" & Label_Icon_Grande(programa_selecionado).Caption & ".exe"
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case -2146697211
        Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada
        
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub Label_Site_Oficial_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar a label ao passar o rato
    Label_Site_Oficial.ForeColor = Laranja
End Sub

Private Sub Label_Total_Downloads_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar procedimento
    Repor_Objectos
End Sub

Private Sub Pic_Actualizar_Click()
    'Des/Activar a opcção
    If Check_Actualizar.Value = 0 Then
        Check_Actualizar.Value = 1
        Pic_Actualizar.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Check_Actualizar.Value = 0
        Pic_Actualizar.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Check_Actualizar_Click()
    'Des/Activar a opcção
    If Check_Actualizar.Value = 1 Then
        Pic_Actualizar.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Pic_Actualizar.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Public Sub Verificar_Opcoes_Programa()
    'Procedimento para carregar os valores guardados das preferências
    Text_Tela_Cheia.Text = ReadINI("Settings", "FullScreen", Localizacao_Ficheiro_Preferencias)
    Text_Lingua.Text = ReadINI("Settings", "Language", Localizacao_Ficheiro_Preferencias)
       
    Dim actualizar As String
    actualizar = ReadINI("Settings", "Check_Update", Localizacao_Ficheiro_Preferencias)
    If actualizar = "1" Then
        Check_Actualizar.Value = 1
        Pic_Actualizar.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Actualizar.Value = 0
        Pic_Actualizar.Picture = Form_Skin.Check_Normal.Picture
    End If
    
    Dim barra As String
    barra = ReadINI("Settings", "Check_Bar_Desktop", Localizacao_Ficheiro_Preferencias)
    If barra = "1" Then
        Check_Barra.Value = 1
        Pic_Barra.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Barra.Value = 0
        Pic_Barra.Picture = Form_Skin.Check_Normal.Picture
    End If
    
    Dim tray As String
    tray = ReadINI("Settings", "Check_Tray", Localizacao_Ficheiro_Preferencias)
    If tray = "1" Then
        Check_Tray.Value = 1
        Pic_Tray.Picture = Form_Skin.Check_Over.Picture
    Else
        Check_Tray.Value = 0
        Pic_Tray.Picture = Form_Skin.Check_Normal.Picture
    End If
    
    'Dimensão e posicao do form
    Text_Form_Top = ReadINI("Settings", "Form_Top", Localizacao_Ficheiro_Preferencias)
    Text_Form_Height = ReadINI("Settings", "Form_Height", Localizacao_Ficheiro_Preferencias)
    Text_Form_Left = ReadINI("Settings", "Form_Left", Localizacao_Ficheiro_Preferencias)
    Text_Form_Width = ReadINI("Settings", "Form_Width", Localizacao_Ficheiro_Preferencias)
End Sub

Private Sub Form_Load()
    'Activar as imagens/ cores originais dos objectos
    Altura_Standard = 7000
    Largura_Standard = 9000
    
    'Carregar opções do programa
    Verificar_Opcoes_Programa
            
    'Criar a lista consoante o nº de assuntos existentes
    Label_Assunto(0).Visible = True
    Dim Objecto As Integer
    For Objecto = 1 To 3
        Load Label_Assunto(Objecto)
        Label_Assunto(Objecto).Move Label_Assunto(Objecto - 1).left, Label_Assunto(Objecto - 1).top + Label_Assunto(Objecto - 1).Height
        Label_Assunto(Objecto).Visible = True
        
        Load Shape_Sombra(Objecto)
        Shape_Sombra(Objecto).Move Shape_Sombra(Objecto - 1).left, Shape_Sombra(Objecto - 1).top + Shape_Sombra(Objecto - 1).Height
        Shape_Sombra(Objecto).Visible = False
        Shape_Sombra(Objecto).ZOrder 1
    Next Objecto
    
    'Propriedades iniciais do formulário
    Carregar_Idioma
    Carregar_Skin
    Desenhar_Formulario
    Ver_Opcoes
    Label_Enterprise.Caption = ""
    
    'Definir os valores de x e y para poder mover o formulário
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    'Valores iniciais
    Utilizador_Logado = False
    Tab_Categoria_Visivel = False
    Tab_Informacao_Visivel = False
    
    'Propriedades da scrollbar
    Me.SrcrollBar1.Value = 0
    Me.SrcrollBar1.Max = 1450
    Me.SrcrollBar1.LargeChange = 50
    
    Me.SrcrollBar2.Value = 0
    Me.SrcrollBar2.Max = 1450
    Me.SrcrollBar2.LargeChange = 50
    
    'Actualizar a localização da pasta do programa
'    Dim Localizacao_Ficheiro_Preferencias As String
'    Localizacao_Ficheiro_Preferencias = App.Path & "\Options\Properties.ini"
    Call WriteINI("Path", "Location_Of_Program", App.Path & "\", (Localizacao_Ficheiro_Preferencias))
        
    'Alterar cores do progreesbar
    ProgressBar1.backcolor = RGB(255, 127, 0) 'laranja
    Dim xpto As Integer: For xpto = 0 To Pic_Linha.Count - 1
        Progresso(xpto).backcolor = RGB(255, 127, 0)
    Next
    
    'Verificar possiveis actualizações do programa
    If Check_Actualizar.Value = 1 Then Verificar_Actualizacoes
    
    'Selecionar a 1ªlinha da lista assunto
    Linha_Selecionada_Assunto = 0
    Shape_Sombra(0).Visible = True
    Label_Assunto(0).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
    
    'Idiomar disponiveis
    Carregar_Idiomas_Existentes
    
    'Selecionar a 1ªlinha da lista linguas
    Linha_Selecionada_Lingua = 0
    Shape_Sombra_Lingua(0).Visible = True
    Label_Lingua(Index).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
    
    'Programas instalados
    programa_selecionado = -1
    Verificar_Programas_Existentes (App.Path & "\Programs\")
    Carregar_Programas_Existentes
End Sub

Public Sub Carregar_Idiomas_Existentes()
    'Procedimento para carregar os idiomas do programa
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
        
        Load Shape_Sombra_Lingua(Objecto)
        Shape_Sombra_Lingua(Objecto).Move Shape_Sombra_Lingua(Objecto - 1).left, Shape_Sombra_Lingua(Objecto - 1).top + Shape_Sombra_Lingua(Objecto - 1).Height
        Shape_Sombra_Lingua(Objecto).Visible = False
        Shape_Sombra_Lingua(Objecto).ZOrder 1
    Next Objecto
        
    'Preencher as label's com as linguas disponiveis
    Dim Z As Integer
    File_Lingua.ListIndex = 0
    For Z = 0 To File_Lingua.ListCount - 1
        Label_Lingua(Z).Caption = left$(File_Lingua.List(Z), InStr(File_Lingua.List(Z), ".") - (1)) 'Retirar a extensão do ficheiro ".lng"
    Next Z
End Sub

Public Sub Carregar_Idioma()
    'Procedimento para carregar o idioma selecionado
    Dim Text_Lingua As String: Text_Lingua = ReadINI("Settings", "Language", Localizacao_Ficheiro_Preferencias)
    Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Text_Lingua & ".lng"
    
    Botao_Fechar.ToolTipText = ReadINI("Main", "Button_Close", Localizacao_Ficheiro_Lingua)
    Form_Skin.Menu_Fechar.Caption = ReadINI("Main", "Button_Close", Localizacao_Ficheiro_Lingua)
    
    Botao_Restaurar.ToolTipText = ReadINI("Main", "Button_Restore", Localizacao_Ficheiro_Lingua)
    Botao_Minimizar.ToolTipText = ReadINI("Main", "Button_Minimize", Localizacao_Ficheiro_Lingua)
    Botao_Maximizar.ToolTipText = ReadINI("Main", "Button_Maximize", Localizacao_Ficheiro_Lingua)
    Label_Menu_Pesquisar.Caption = ReadINI("Main", "Menu_Search", Localizacao_Ficheiro_Lingua)
    Label_Menu_Instalados.Caption = ReadINI("Main", "Menu_Installed", Localizacao_Ficheiro_Lingua)
    Label_Menu_Partilhar.Caption = ReadINI("Main", "Menu_Share", Localizacao_Ficheiro_Lingua)
    Label_Menu_Opcoes.Caption = ReadINI("Main", "Menu_Options", Localizacao_Ficheiro_Lingua)
    Label_Menu_Suporte.Caption = ReadINI("Main", "Menu_Support", Localizacao_Ficheiro_Lingua)
    Label_Menu_Sobre.Caption = ReadINI("Main", "Menu_About", Localizacao_Ficheiro_Lingua)
    Label_Programas.Caption = ReadINI("Main", "Tab_Select_Category", Localizacao_Ficheiro_Lingua)
    Text_Pesquisa.Text = ReadINI("Main", "Text_Search", Localizacao_Ficheiro_Lingua)
    Label_Titulo_Frame_Programas(0).Caption = ReadINI("Main", "Menu_Search", Localizacao_Ficheiro_Lingua)
    Label_Titulo_Frame_Programas(1).Caption = ReadINI("Main", "Label_Select_Category", Localizacao_Ficheiro_Lingua)
    Label_Update.Caption = ReadINI("Main", "Button_Update_Program", Localizacao_Ficheiro_Lingua)
    Label_Nenum_Resultado.Caption = ReadINI("Main", "Label_No_Programs_Found", Localizacao_Ficheiro_Lingua)
    Dim I As Integer: For I = 0 To Label_Mais_Informacoes.Count - 1
        Label_Mais_Informacoes(I).Caption = ReadINI("Main", "Button_More_Information", Localizacao_Ficheiro_Lingua)
        Label_Executar_Programa(I).Caption = ReadINI("Main", "Button_Execute", Localizacao_Ficheiro_Lingua)
    Next
    Idioma_Button_Transfer_Program = ReadINI("Main", "Button_Transfer_Program", Localizacao_Ficheiro_Lingua)
    Idioma_Button_Execute_Program = ReadINI("Main", "Button_Execute_Program", Localizacao_Ficheiro_Lingua)
    Idioma_Button_Remove_Program = ReadINI("Main", "Button_Remove_Program", Localizacao_Ficheiro_Lingua)
    Idioma_Button_Cancel_Program = ReadINI("Main", "Button_Cancel_Program", Localizacao_Ficheiro_Lingua)
    Label_Barra_Comentario.Caption = ReadINI("Main", "Label_I_Like", Localizacao_Ficheiro_Lingua)
    Label_Eu_Gosto.Caption = ReadINI("Main", "Button_I_Like", Localizacao_Ficheiro_Lingua)
    Label_Download.Caption = Idioma_Button_Transfer_Program
    Label_Cancelar.Caption = Idioma_Button_Cancel_Program
    Label_Executar.Caption = Idioma_Button_Execute_Program
    Idioma_Transferring_File = ReadINI("Main", "Label_Transferring_File", Localizacao_Ficheiro_Lingua)
    Idioma_Label_Rate = ReadINI("Main", "Label_Rate", Localizacao_Ficheiro_Lingua)
    Label_Site_Oficial.Caption = ReadINI("Main", "Label_Site_Official", Localizacao_Ficheiro_Lingua)
    Label_Close_Suporte.ToolTipText = ReadINI("Main", "Close_Frame_Error", Localizacao_Ficheiro_Lingua)
    
    Idioma_Erro = ReadINI("Main", "Label_Error", Localizacao_Ficheiro_Lingua)
    Idioma_Descricao = ReadINI("Main", "Label_Description", Localizacao_Ficheiro_Lingua)
    Idioma_Erro_Execucao = ReadINI("Main", "Error_Execution", Localizacao_Ficheiro_Lingua)
    Idioma_Conectar_Servidor = ReadINI("Main", "Error_Connect", Localizacao_Ficheiro_Lingua)
    Idioma_Internet_Desligada = ReadINI("Main", "Error_Internet", Localizacao_Ficheiro_Lingua)
    
    'Frame instalados------------------------------------------------------------------------------------------
    Label_Frame_Instalados.Caption = ReadINI("Main", "Menu_Installed", Localizacao_Ficheiro_Lingua)
    Label_Run.Caption = ReadINI("Main", "Button_Run_Program", Localizacao_Ficheiro_Lingua)
    Label_Desinstalar.Caption = ReadINI("Main", "Button_Uninstall_Program", Localizacao_Ficheiro_Lingua)
    
    'Frame partilhar-------------------------------------------------------------------------------------------
    Label_Frame_Partilhar.Caption = ReadINI("Main", "Menu_Share", Localizacao_Ficheiro_Lingua)
    Label_Carregar.Caption = ReadINI("Main", "Button_Upload_Program", Localizacao_Ficheiro_Lingua)
    Lb_Email.Caption = ReadINI("Main", "Label_Of", Localizacao_Ficheiro_Lingua)
    Lb_Info.Caption = ReadINI("Main", "Label_Info", Localizacao_Ficheiro_Lingua)
    Lb_Empresa.Caption = ReadINI("Main", "Label_Share_Business", Localizacao_Ficheiro_Lingua)
    Lb_Nome.Caption = ReadINI("Main", "Label_Share_Name", Localizacao_Ficheiro_Lingua)
    Lb_Descricao.Caption = ReadINI("Main", "Label_Share_Description", Localizacao_Ficheiro_Lingua)
    Lb_Informacao.Caption = ReadINI("Main", "Label_Share_Information", Localizacao_Ficheiro_Lingua)
    Lb_Site.Caption = ReadINI("Main", "Label_Share_Web", Localizacao_Ficheiro_Lingua)
    Lb_Download.Caption = ReadINI("Main", "Label_Share_Download", Localizacao_Ficheiro_Lingua)
    Lb_Nota.Caption = ReadINI("Main", "Label_Share_Info", Localizacao_Ficheiro_Lingua)
    Label_Close_Partilhar.ToolTipText = ReadINI("Main", "Close_Frame_Error", Localizacao_Ficheiro_Lingua)
    
    'Frame opções----------------------------------------------------------------------------------------------
    Label_Frame_Opcoes.Caption = ReadINI("Main", "Menu_Options", Localizacao_Ficheiro_Lingua)
    Label_Aplicar.Caption = ReadINI("Main", "Button_Apply", Localizacao_Ficheiro_Lingua)
    Check_Actualizar.Caption = ReadINI("Main", "Ckeck_Update", Localizacao_Ficheiro_Lingua)
    Label_Idioma_Programa.Caption = ReadINI("Main", "Label_Language_Program", Localizacao_Ficheiro_Lingua)
    Check_Barra.Caption = ReadINI("Main", "Check_Bar", Localizacao_Ficheiro_Lingua)
    Check_Tray.Caption = ReadINI("Main", "Check_Tray", Localizacao_Ficheiro_Lingua)
    Label_Opcoes_Actualizadas.Caption = ReadINI("Main", "Label_Options_Successfully_Updated", Localizacao_Ficheiro_Lingua)
    Label_Opcoes_Actualizadas.left = Label_Frame_Opcoes.left + Label_Frame_Opcoes.Width + 10
    
    'Frame suporte---------------------------------------------------------------------------------------------
    Label_Frame_Suporte.Caption = ReadINI("Main", "Menu_Support", Localizacao_Ficheiro_Lingua)
    Label_De.Caption = ReadINI("Main", "Label_Of", Localizacao_Ficheiro_Lingua)
    Label_Info.Caption = ReadINI("Main", "Label_Info", Localizacao_Ficheiro_Lingua)
    Label_Texto.Caption = ReadINI("Main", "Label_Subject", Localizacao_Ficheiro_Lingua)
    Label_Mensagem.Caption = ReadINI("Main", "Label_Message", Localizacao_Ficheiro_Lingua)
    Label_Assunto(0).Caption = ReadINI("Main", "Label_Report", Localizacao_Ficheiro_Lingua)
    Label_Assunto(1).Caption = ReadINI("Main", "Label_Suggestion", Localizacao_Ficheiro_Lingua)
    Label_Assunto(2).Caption = ReadINI("Main", "Label_Question", Localizacao_Ficheiro_Lingua)
    Label_Assunto(3).Caption = ReadINI("Main", "Label_Other", Localizacao_Ficheiro_Lingua)
    Label_Enviar.Caption = ReadINI("Main", "Button_Ok", Localizacao_Ficheiro_Lingua)
    Label_Limpar.Caption = ReadINI("Main", "Button_Clean", Localizacao_Ficheiro_Lingua)
    
    'Frame sobre------------------------------------------------------------------------------------------------
    Label_Frame_Sobre.Caption = ReadINI("Main", "Menu_About", Localizacao_Ficheiro_Lingua)
    Label_Website.Caption = ReadINI("Main", "Label_Website", Localizacao_Ficheiro_Lingua)
    Text_About.Text = App.ProductName & vbNewLine
    Text_About.Text = Text_About.Text + ReadINI("Main", "Label_Version", Localizacao_Ficheiro_Lingua) & " " & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine & vbNewLine
    Text_About.Text = Text_About.Text + ReadINI("Main", "Text_About1", Localizacao_Ficheiro_Lingua) & vbNewLine
    Text_About.Text = Text_About.Text + ReadINI("Main", "Text_About2", Localizacao_Ficheiro_Lingua) & " "
    Text_About.Text = Text_About.Text + ReadINI("Main", "Text_About3", Localizacao_Ficheiro_Lingua) & vbNewLine
    Text_About.Text = Text_About.Text + ReadINI("Main", "Text_About4", Localizacao_Ficheiro_Lingua) & vbNewLine
    Text_About.Text = Text_About.Text + ReadINI("Main", "Text_About5", Localizacao_Ficheiro_Lingua) & " Nelson do Carmo, "
    Text_About.Text = Text_About.Text + ReadINI("Main", "Text_About6", Localizacao_Ficheiro_Lingua) & vbNewLine & vbNewLine
    Text_About.Text = Text_About.Text + App.LegalCopyright & " - " & ReadINI("Main", "Label_Informatic", Localizacao_Ficheiro_Lingua) & vbNewLine
    Text_About.Text = Text_About.Text + ReadINI("Main", "Label_Contact", Localizacao_Ficheiro_Lingua) & ": " & "nikyts@hotmail.com" & vbNewLine
End Sub

Public Sub Carregar_Skin()
    'Procedimento para carregar o skin escolhido
    With Form_Skin
        'Me.BackColor = .Cor_do_Fundo_dos_Formularios.BackColor
        Shape_Contorno.BorderColor = .Cor_Form_BorderColor.backcolor
        Fundo_Barra_ControlBox.Picture = .Fundo_Barra_ControlBox.Picture
        Label_Titulo.ForeColor = .Cor_Label_Barra_Titulo.backcolor
        Botao_Fechar.Picture = .Botao_Fechar.Picture
        Botao_Restaurar.Picture = .Botao_Restaurar_Normal.Picture
        Botao_Minimizar.Picture = .Botao_Minimizar_Normal.Picture
        Botao_Maximizar.Picture = .Botao_Maximizar_Normal.Picture
        Image_Update.Picture = .Button_Menu_Normal.Picture
        Icon_Update.Picture = Form_Skin.Icon_Menu_Normal.Picture
        
        Menu_Pesquisar.Picture = .Icon_Pesquisar_Down.Picture 'Separador activo
        Menu_Instalados.Picture = .Icon_Instalados_Normal.Picture
        Menu_Partilhar.Picture = .Icon_Partilhar_Normal.Picture
        Menu_Opcoes.Picture = .Icon_Opcoes_Normal.Picture
        Menu_Suporte.Picture = .Icon_Suporte_Normal.Picture
        Menu_Sobre.Picture = .Icon_Sobre_Normal.Picture
        Frame_Erro_Partilhar.Picture = .Image_Balao.Picture
        Frame_Conteudo.backcolor = &HF9F9F9 'Cinza claro
        Frame_Partilhar.backcolor = &HF9F9F9
        
        'Frame suporte técnico-------------------------------------------------------------------------------------------
        Label_Enviar.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Label_Limpar.ForeColor = .Cor_da_Letra_do_Botao.backcolor
        Seta_Assunto.Picture = .Seta_Combo.Picture
    End With
End Sub

Private Sub Verificar_Actualizacoes()
    On Error GoTo Corrige_Erro
    DoEvents
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.open "GET", "http://www.nikyts.com/gadgets/" & "verificarversao.asp?", False
    servidor.send 'envia o pedido para o servidor
    
    'Verificar os dados acesso
    If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 Then
            'Ler os dados do xml referente aos dados do perfil do utilizador
            Dim versao_actual, nova_versao As String
            versao_actual = App.Major & App.Minor & App.Revision
            nova_versao = servidor.responseText 'CInt(responseText)
            
            'Verificar se existem versões novas
            If versao_actual < nova_versao Then
                Botao_Update.Visible = True
            End If
        End If
    End If
    Set servidor = Nothing
    
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Botao_Update.Visible = False
'Select Case Err.Number
'    Case -2146697211
'        Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada
'
'    Case Else
'        'Correção de outros erros que poderão surgir
'        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & Err.Number & vbNewLine & Idioma_Descricao & " " & Err.Description
'End Select
End Sub

Private Sub Form_Resize()
    'Chamar o procedimento
    Desenhar_Formulario
End Sub

Private Sub Frame_Centro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Limpar o conteudo da label
    Label_Titulo_Frame_Programas(2).Caption = ""
End Sub

Private Sub Frame_Pesquisar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Limpar o conteudo da label
    Label_Titulo_Frame_Programas(2).Caption = ""
End Sub

Private Sub Icon_Pasta_Categoria_Click(Index As Integer)
    'Ocultar o separador referente ao programa que foi solicitado ver mais informações
    'Selecionar a categoria do programa
    Select Case Icon_Pasta_Categoria(Index).Index
        Case 0
            Me.MousePointer = 11
            Selecionar_Categoria "Ferramentas", ReadINI("Main", "Folder_Tools", Localizacao_Ficheiro_Lingua)
            Label_Pesquisa.Caption = "Ferramentas"
            Me.MousePointer = 0
        Case 1
            Me.MousePointer = 11
            Selecionar_Categoria "Som e video", ReadINI("Main", "Folder_Media", Localizacao_Ficheiro_Lingua)
            Label_Pesquisa.Caption = "Som e video"
            Me.MousePointer = 0
        Case 2
            Me.MousePointer = 11
            Selecionar_Categoria "Acessibilidade", ReadINI("Main", "Folder_Accessibility", Localizacao_Ficheiro_Lingua)
            Label_Pesquisa.Caption = "Acessibilidade"
            Me.MousePointer = 0
        Case 3
            Me.MousePointer = 11
            Selecionar_Categoria "Internet", ReadINI("Main", "Folder_Internet", Localizacao_Ficheiro_Lingua)
            Label_Pesquisa.Caption = "Internet"
            Me.MousePointer = 0
        Case 4
            Me.MousePointer = 11
            Selecionar_Categoria "Jogos", ReadINI("Main", "Folder_Games", Localizacao_Ficheiro_Lingua)
            Label_Pesquisa.Caption = "Jogos"
            Me.MousePointer = 0
    End Select
End Sub

Private Sub Icon_Pasta_Categoria_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Indique que pasta está a ser selecionada
    Select Case Icon_Pasta_Categoria(Index).Index
        Case 0
            Label_Titulo_Frame_Programas(2).Caption = ReadINI("Main", "Folder_Tools", Localizacao_Ficheiro_Lingua)
        Case 1
            Label_Titulo_Frame_Programas(2).Caption = ReadINI("Main", "Folder_Media", Localizacao_Ficheiro_Lingua)
        Case 2
            Label_Titulo_Frame_Programas(2).Caption = ReadINI("Main", "Folder_Accessibility", Localizacao_Ficheiro_Lingua)
        Case 3
            Label_Titulo_Frame_Programas(2).Caption = ReadINI("Main", "Folder_Internet", Localizacao_Ficheiro_Lingua)
        Case 4
            Label_Titulo_Frame_Programas(2).Caption = ReadINI("Main", "Folder_Games", Localizacao_Ficheiro_Lingua)
    End Select
End Sub

Private Sub Label_Assunto_Click(Index As Integer)
    'Indicar a lingua selecionada pelo utilizador
    Text_Assunto.Text = Label_Assunto(Index).Caption
    
    Lista_Assunto.Visible = False
    Text_Assunto.SetFocus
End Sub

Private Sub Label_Assunto_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar linha
    If Linha_Selecionada_Assunto = Index Then Exit Sub
    Shape_Sombra(Linha_Selecionada_Assunto).Visible = False
    Label_Assunto(Linha_Selecionada_Assunto).ForeColor = Form_Skin.Cor_Letra_Textbox.backcolor
    Shape_Sombra(Index).Visible = True
    Label_Assunto(Index).ForeColor = Form_Skin.Cor_Fundo_Textbox.backcolor
    Linha_Selecionada_Assunto = Index
End Sub

Private Sub Label_Close_Suporte_Click()
    'Ocultar frame erro
    Frame_Erro_Suporte.Visible = False
End Sub

Private Sub Label_Enviar_Click()
    'Enviar a mensagem para o suporte técnico
    'On Error GoTo Corrige_Erro
    Frame_Erro_Suporte.top = (Barra_Text_Email.top + (Barra_Text_Email.ScaleHeight / 2)) - 40
    Frame_Erro_Suporte.Visible = False
    
    'Verificar o preencimento das textboxs
    If Len(Trim(Text_Email.Text)) = 0 Then
        Label_Erro_Suporte.Caption = ReadINI("Main", "Message_Required_Field", Localizacao_Ficheiro_Lingua)
        Frame_Erro_Suporte.top = (Barra_Text_Email.top + (Barra_Text_Email.ScaleHeight / 2)) - 40
        Frame_Erro_Suporte.Visible = True
        Text_Email.SetFocus
        Exit Sub
    End If
    
    'Verifica se o campo email está no formato correcto
    If Not IsEmail(Text_Email.Text) Then
        Label_Erro_Suporte.Caption = ReadINI("Main", "Message_Email_Invalid", Localizacao_Ficheiro_Lingua)
        Frame_Erro_Suporte.top = (Barra_Text_Email.top + (Barra_Text_Email.ScaleHeight / 2)) - 40
        Frame_Erro_Suporte.Visible = True
        Text_Email.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(Text_Assunto.Text)) = 0 Then
        Label_Erro_Suporte.Caption = ReadINI("Main", "Message_Required_Field", Localizacao_Ficheiro_Lingua)
        Frame_Erro_Suporte.top = (Barra_Text_Assunto.top + (Barra_Text_Assunto.ScaleHeight / 2)) - 40
        Frame_Erro_Suporte.Visible = True
        Text_Assunto.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(Text_Mensagem.Text)) = 0 Then
        Label_Erro_Suporte.Caption = ReadINI("Main", "Message_Required_Field", Localizacao_Ficheiro_Lingua)
        Frame_Erro_Suporte.top = (Barra_Text_Mensagem.top + (Barra_Text_Mensagem.ScaleHeight / 2)) - 40
        Frame_Erro_Suporte.Visible = True
        Text_Mensagem.SetFocus
        Exit Sub
    End If
    
    Me.MousePointer = 11
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.open "GET", "http://www.nikyts.com/suporte/" & "enviarmensagem.asp?Email=" & Text_Email.Text & "&Assunto=" & App.ProductName & " - " & Text_Assunto.Text & "&Mensagem=" & Text_Mensagem.Text, False
    servidor.send 'envia o pedido para o servidor

    'Verificar os dados acesso
    If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then ' 4 - deu resposta e 200 validou
            Mensagem_de_Aviso "Information", ReadINI("Main", "Info_Posted", Localizacao_Ficheiro_Lingua)
            
            Me.MousePointer = 0
            Frame_Erro_Suporte.Visible = False
            Limpa_Campos
            Text_Email.SetFocus
        End If
    End If
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case -2146697211
        Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada
        
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub Label_Eu_Gosto_Click()
    'Votar no programa
    On Error GoTo Corrige_Erro
    If Label_Id_Programa.Caption = "" Or Label_Votos.Caption = "" Then Exit Sub
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    
    'Adicionar um voto á avaliação do programa
    Label_Votos.Caption = Val(Label_Votos.Caption) + 1
    servidor.open "GET", "http://www.nikyts.com/gadgets/" & "actualizaravaliacao.asp?id_programa=" & Label_Id_Programa.Caption & "&avaliacao=" & Label_Votos.Caption, False
    servidor.send 'envia o pedido para o servidor

    'Actualizar a senha
    If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        With Form_Principal
            If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then ' 4 - deu resposta e 200 validou
                
                'Recebe confirmação de que o voto foi recebido com sucesso
                'Mensagem_de_Aviso "Information", "O seu voto foi atribuido ao programa com sucesso!" & vbNewLine & "Obrigado pela sua contribuição."
                Mensagem_de_Aviso "Information", ReadINI("Main", "Message_Vote_Assigned", Localizacao_Ficheiro_Lingua) & vbNewLine & ReadINI("Main", "Message_Thanks", Localizacao_Ficheiro_Lingua)
                
                'Actualizar a avaliação do programa
                If Label_Votos.Caption = "1" Then
                    Label_Total.Caption = Idioma_Label_Rate & ": " & Label_Votos.Caption
                Else
                    Label_Total.Caption = Idioma_Label_Rate & ": " & Label_Votos.Caption
                End If
    
                'Avaliação do programa, Estrelas
                If Val(Label_Votos.Caption) < 20 Then
                    Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_0.Picture
                
                ElseIf Val(Label_Votos.Caption) >= 20 And Val(Label_Votos.Caption) < 40 Then
                    Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_1.Picture
                
                ElseIf Val(Label_Votos.Caption) >= 40 And Val(Label_Votos.Caption) < 60 Then
                    Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_2.Picture
                
                ElseIf Val(Label_Votos.Caption) >= 60 And Val(Label_Votos.Caption) < 80 Then
                    Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_3.Picture
                
                ElseIf Val(Label_Votos.Caption) >= 80 And Val(Label_Votos.Caption) < 100 Then
                    Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_4.Picture
                
                ElseIf Val(Label_Votos.Caption) > 100 Then
                    Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_5.Picture
                End If
            End If
        End With
    End If
    Set servidor = Nothing
    
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case -2146697211
        Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada
        
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub Label_Executar_Programa_Click(Index As Integer)
    'Executar o programa automaticamente
    On Error GoTo Corrige_Erro
    Shell App.Path & "\Programs\" & Label_Nome(Index).Caption & "\" & Label_Nome(Index).Caption & ".exe"
      
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case -2146697211
        Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada
        
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub Label_Iniciar_Transferencia_Click(Index As Integer)
    'Iniciar a transferência do programa
    Me.MousePointer = 11
    'Label_Titulo.Caption = "Centro de programas Nikyts - Aguarde..."
    Verificar_Pastas
    Botao_Remover_Transferencia(Index).Visible = False
    Label_Remover_Transferencia(Index).Visible = False
    
    'Proceder á transferência dos respectivos programas
    Linha_Selecionada = Label_Remover_Transferencia(Index).Index
    Text_Servidor.Text = "http://www.nikyts.com/gadgets/programas/" & Label_Programa(Index).Caption
    Progresso(Index).Visible = True
    Download_Programa.DownloadFile Text_Servidor.Text, App.Path & "\Programs\" & Label_Programa(Index).Caption
    On Error GoTo 0 'Tratamento de erros
    Me.MousePointer = 0
    
Exit Sub
errHand:
End Sub

Private Sub Label_Limpar_Click()
    'Limpar todos os campos
    Limpa_Campos
End Sub

Public Sub Limpa_Campos()
    'Limpa o conteudo das caixas de texto
    Text_Email.Text = ""
    Text_Assunto.Text = ""
    Text_Mensagem.Text = ""
End Sub

Private Sub Label_Menu_Instalados_Click()
    'Atalho para
    Menu_Instalados_Click
End Sub

Private Sub Label_Menu_Opcoes_Click()
    'Atalho para
    Menu_Opcoes_Click
End Sub

Private Sub Label_Menu_Partilhar_Click()
    'Atalho para
    Menu_Partilhar_Click
End Sub

Private Sub Label_Site_Oficial_Click()
    'Abrir página pessoal
    If Label_Site_Programa.Caption = Empty Then Exit Sub
    Call ShellExecute(0, "open", Label_Site_Programa.Caption, vbNullString, vbNullString, SW_NORMAL)
End Sub

Private Sub Label_Update_Click()
    'Efectuar actualizões do programa
    On Error GoTo Corrige_Erro
    Shell App.Path & "\Options\Update.exe"
    Form_Principal.Botao_Fechar_Click
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub Label_Update_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mostar a imagem down
    Image_Update.Picture = Form_Skin.Button_Menu_Down.Picture
    Icon_Update.Picture = Form_Skin.Icon_Menu_Down.Picture
End Sub

Private Sub Label_Update_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Legendas on-line
    Image_Update.Picture = Form_Skin.Button_Menu_Normal.Picture
    Icon_Update.Picture = Form_Skin.Icon_Menu_Normal.Picture
End Sub

Private Sub Label_Website_Click()
    'Abrir site oficial
    If Label_Website.Caption = Empty Then Exit Sub
    Call ShellExecute(0, "open", "http://www.nikyts.com", vbNullString, vbNullString, SW_NORMAL)
End Sub

Private Sub Menu_Instalados_Click()
    'Ver programas instalados
    If Frame_Instalados.Visible = True Then Exit Sub
    Ocultar_Frames
    Repor_Menus
    Ocultar_Separadores
    Frame_Instalados.Visible = True
    Menu_Instalados.Picture = Form_Skin.Icon_Instalados_Down.Picture
    Label_Menu_Instalados.ForeColor = Azul
    Botao_Run.Visible = True
    Botao_Desinstalar.Visible = True
End Sub

Private Sub Menu_Opcoes_Click()
    'Ver opções do programa
    On Error Resume Next
    If Frame_Opcoes.Visible = True Then Exit Sub
    Verificar_Opcoes_Programa
    
    Ocultar_Frames
    Repor_Menus
    Ocultar_Separadores
    Frame_Opcoes.Visible = True
    Menu_Opcoes.Picture = Form_Skin.Icon_Opcoes_Down.Picture
    Label_Menu_Opcoes.ForeColor = Azul
    Botao_Aplicar.Visible = True
    Text_Lingua.SetFocus
End Sub

Private Sub Menu_Partilhar_Click()
    'Partilhar software
    On Error Resume Next
    If Frame_Partilhar.Visible = True Then Exit Sub
    Ocultar_Frames
    Repor_Menus
    Ocultar_Separadores
    Frame_Partilhar.Visible = True
    Menu_Partilhar.Picture = Form_Skin.Icon_Partilhar_Down.Picture
    Label_Menu_Partilhar.ForeColor = Azul
    Botao_Carregar.Visible = True
    Txt_Email.SetFocus
End Sub

Private Sub Pic_Barra_Click()
    'Des/Activar a opcção
    If Check_Barra.Value = 0 Then
        Check_Barra.Value = 1
        Pic_Barra.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Check_Barra.Value = 0
        Pic_Barra.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Pic_Tray_Click()
    'Des/Activar a opcção
    If Check_Tray.Value = 0 Then
        Check_Tray.Value = 1
        Pic_Tray.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Check_Tray.Value = 0
        Pic_Tray.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'pichook é uma picture box, utilizada pelo Windows para reconhecer o ícone na barra de tarefas.
    Static Rec As Boolean, Msg As Long
    Msg = X / Screen.TwipsPerPixelX
    If Rec = False Then
        Rec = True
        Select Case Msg
            Case WM_LBUTTONDBLCLK:
                'Remover do sistema o icon do programa
                Remover_Tray_Icon
    
            Case WM_LBUTTONDOWN:
            Case WM_LBUTTONUP:
            Case WM_RBUTTONDBLCLK:
            Case WM_RBUTTONDOWN:
            Case WM_RBUTTONUP:
                'Ver o menu icon se for pressionado o botão direito
                Form_Skin.PopupMenu Form_Skin.Menu_Tray
        End Select
        Rec = False
    End If
End Sub

Public Sub Remover_Tray_Icon()
    'Remover do sistema o icon do programa
    Me.Show
    Modo_Tray = False
    t.cbSize = Len(t)
    t.hwnd = pichook.hwnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub


Private Sub Seta_Assunto_Click()
    'Ver/ocultar lista
    If Lista_Assunto.Visible = True Then
        Lista_Assunto.Visible = False
    Else
        Lista_Assunto.Visible = True
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

Public Sub Ver_Opcoes()
    'Procedimento ver o estado da janela
    If Text_Tela_Cheia.Text = "True" Then
        Botao_Maximizar_Click
        Tela_Cheia = True
    Else
        Botao_Restaurar_Click
        Tela_Cheia = False
    End If
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    On Error GoTo Corrige_Erro
    If Me.WindowState = 1 Then Exit Sub

    Ajustar_Formulario Me, False, True, False, False
    
    With Barra_Ferramentas
        .Height = Fundo_Barra_Ferramentas.Height
        .top = Barra_ControlBox.top + Barra_ControlBox.ScaleHeight
        .Width = Barra_ControlBox.ScaleWidth
        .left = 0
    End With
    
    With Fundo_Barra_Ferramentas
        .Stretch = True
        .top = 0
        .Width = Barra_Ferramentas.ScaleWidth
        .left = 0
    End With
    
    With Menu_Pesquisar
        .top = 0
        .Height = Form_Skin.Icon_Pesquisar_Normal.Height
        .left = 1
        .Width = Form_Skin.Icon_Pesquisar_Normal.Width
    End With
    
    With Menu_Instalados
        .top = Menu_Pesquisar.top
        .Height = Menu_Pesquisar.ScaleHeight
        .left = Menu_Pesquisar.left + Menu_Pesquisar.ScaleWidth
        .Width = Menu_Pesquisar.ScaleWidth
    End With
    
    With Menu_Partilhar
        .top = Menu_Pesquisar.top
        .Height = Menu_Pesquisar.ScaleHeight
        .left = Menu_Instalados.left + Menu_Instalados.ScaleWidth
        .Width = Menu_Pesquisar.ScaleWidth
    End With
    
    With Menu_Opcoes
        .top = Menu_Pesquisar.top
        .Height = Menu_Pesquisar.ScaleHeight
        .left = Menu_Partilhar.left + Menu_Partilhar.ScaleWidth
        .Width = Menu_Pesquisar.ScaleWidth
    End With
    
    With Menu_Suporte
        .top = Menu_Pesquisar.top
        .Height = Menu_Pesquisar.ScaleHeight
        .left = Menu_Opcoes.left + Menu_Opcoes.ScaleWidth
        .Width = Menu_Pesquisar.ScaleWidth
    End With
    
    With Menu_Sobre
        .top = Menu_Pesquisar.top
        .Height = Menu_Pesquisar.ScaleHeight
        .left = Menu_Suporte.left + Menu_Suporte.ScaleWidth
        .Width = Menu_Pesquisar.ScaleWidth
    End With
    
    With Barra_Botoes
        .Height = Fundo_Barra_Botoes.Height
        .top = Barra_Ferramentas.top + Barra_Ferramentas.ScaleHeight
        .Width = Barra_ControlBox.ScaleWidth - 2
        .left = 1
    End With
    
    With Fundo_Barra_Botoes
        .Stretch = True
        .top = 0
        .Width = Barra_Botoes.ScaleWidth
        .left = 0
    End With
    
    With Barra_Caixa_Pesquisa
        .Height = Form_Skin.Image_Caixa_Pesquisa.Height
        .top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Image_Caixa_Pesquisa.Width
        .left = Barra_Botoes.ScaleWidth - .ScaleWidth - 8
    End With
    
    With Contorno_Caixa_Pesquisa
        .top = 0
        .Height = Barra_Caixa_Pesquisa.ScaleHeight
        .left = 0
        .Width = Barra_Caixa_Pesquisa.ScaleWidth
    End With
    
    With Text_Pesquisa
        .top = 6
        .Height = 17 'Barra_Caixa_Pesquisa.ScaleHeight - 2
        .left = 3
        .Width = Barra_Caixa_Pesquisa.ScaleWidth - Botao_Pesquisar.Width - 8
    End With
    
    With Botao_Pesquisar
        .top = (Barra_Caixa_Pesquisa.ScaleHeight - .Height) / 2
        .left = Barra_Caixa_Pesquisa.ScaleWidth - .Width - 2
    End With
    
    With Barra_Detalhes
        .Height = Fundo_Barra_Detalhes.Height
        .top = Me.ScaleHeight - .Height
        .Width = Barra_ControlBox.ScaleWidth
        .left = Barra_ControlBox.left
    End With
    
    With Fundo_Barra_Detalhes
        .Stretch = True
        .top = 0
        .Width = Barra_Detalhes.ScaleWidth
        .left = 0
    End With
    
    With Label_Utilizador_Logado
        .top = (Barra_Detalhes.ScaleHeight - .Height) / 2
        .left = 10
    End With
    
    With Botao_Redimensionar
        .top = Barra_Detalhes.ScaleHeight - .Height
        .left = Barra_Detalhes.ScaleWidth - .Width
    End With
    
    With Frame_Centro
        .top = Barra_Botoes.top + Barra_Botoes.ScaleHeight
        .Height = Me.ScaleHeight - Barra_ControlBox.ScaleHeight - Barra_Ferramentas.ScaleHeight - Barra_Botoes.ScaleHeight - Barra_Detalhes.ScaleHeight
        .Width = Barra_Botoes.ScaleWidth
        .left = Barra_Botoes.left
    End With
    
    With Frame_Pesquisar
        If Frame_Centro.Height >= .Height Then
            .top = (Frame_Centro.ScaleHeight - .Height) / 2
        Else
            .top = 0
        End If
        .Width = ((Icon_Pasta_Categoria.Count) * Icon_Pasta_Categoria(0).Width) + (40 * (Icon_Pasta_Categoria.Count + 1))
        If Frame_Centro.Width >= .Width Then
            .left = (Frame_Centro.ScaleWidth - .ScaleWidth) / 2
        Else
            .left = 0
        End If
    End With
    
    With Label_Titulo_Frame_Programas(0)
        .top = 0
        .left = (Frame_Pesquisar.ScaleWidth - .Width) / 2
    End With
    
    With Label_Titulo_Frame_Programas(1)
        .top = Label_Titulo_Frame_Programas(0).top + Label_Titulo_Frame_Programas(0).Height + 5
        .left = (Frame_Pesquisar.ScaleWidth - .Width) / 2
    End With
    
    Icon_Pasta_Categoria(0).top = Label_Titulo_Frame_Programas(1).top + Label_Titulo_Frame_Programas(1).Height + 40
    Icon_Pasta_Categoria(0).left = 40
    Dim pastas As Integer: For pastas = 1 To Icon_Pasta_Categoria.Count - 1
        Icon_Pasta_Categoria(pastas).top = Icon_Pasta_Categoria(0).top
        Icon_Pasta_Categoria(pastas).left = Icon_Pasta_Categoria(pastas - 1).left + Icon_Pasta_Categoria(pastas - 1).Width + 40
    Next
    
    With Label_Titulo_Frame_Programas(2)
        .top = Icon_Pasta_Categoria(0).top + Icon_Pasta_Categoria(0).Height + 10
        .left = 0
        .Width = Frame_Pesquisar.ScaleWidth
    End With
    
    With Separador_Programas
        .Height = Form_Skin.Fundo_Separadores.Height
        .top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .Width = Label_Programas.Width + 8 + 15
        .left = 8
        Extermidade_Programas.left = .ScaleWidth - Extermidade_Programas.Width
    End With
    
    With Separador_Categorias
        .Height = Form_Skin.Fundo_Separadores.Height
        .top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .left = Separador_Programas.left + Separador_Programas.ScaleWidth
    End With
    
    With Separador_Informacoes
        .Height = Form_Skin.Fundo_Separadores.Height
        .top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .left = Separador_Categorias.left + Separador_Categorias.ScaleWidth
    End With
    
    With Label_Programas
        .top = (Separador_Programas.ScaleHeight - .Height) / 2
    End With
    
    With Label_Categorias
        .top = (Separador_Categorias.ScaleHeight - .Height) / 2
    End With
    
    With Label_Informacoes
        .top = (Separador_Informacoes.ScaleHeight - .Height) / 2
    End With
    
    With Frame_Lista
        .top = 0
        .Height = Frame_Centro.Height
        .Width = Frame_Centro.Width
        .left = 0
    End With
    
    With Frame_Conteudo
        .top = 0
        .Height = Frame_Centro.Height
        .Width = Frame_Centro.Width
        .left = 0
    End With
    
    With SrcrollBar1
        .top = 0
        .Height = Frame_Conteudo.ScaleHeight
        .left = Frame_Conteudo.ScaleWidth - .Width
    End With
    
    With Frame_Informacoes
        .top = 0
        '.Height = Frame_Conteudo.ScaleHeight
        .left = 0
        .Width = Frame_Conteudo.ScaleWidth - SrcrollBar1.Width
    End With
    
    With Frame_Avaliacao
        .Height = Form_Skin.Image_Estrelas_0.Height + 40
        .top = Label_Nome_Programa.top
        .Width = Form_Skin.Image_Estrelas_0.Width
        .left = Frame_Informacoes.ScaleWidth - .ScaleWidth - 20
    End With
    
    'Connstruir a lista de programas-----------------------------------------------------------------------------------------------
    Dim Linha, Altura As Integer
    Linha = 0: Altura = 0
    For Linha = 0 To Pic_Linha.Count - 1
        With Pic_Linha(Linha)
            .Height = Form_Skin.Linha_Normal.Height
            .top = Altura
            .left = 0
            .Width = Frame_Lista.ScaleWidth
        End With
    
        Botao_Remover_Transferencia(Linha).left = Botao_Remover_Transferencia(0).left
           
        Altura = Altura + Form_Skin.Linha_Normal.Height
    Next Linha
    
    'Barra transferir-------------------------------------------------------------------------------------
    With Barra_Transferir
        .Height = Shape_Transferir.Height
        .Width = Frame_Informacoes.ScaleWidth - 24 - 24
        .left = 24
    End With
    
    With Shape_Transferir
        .Height = Barra_Transferir.ScaleHeight
        .top = 0
        .Width = Barra_Transferir.ScaleWidth
        .left = 0
    End With
    
    With Botao_Download
        .Height = Form_Skin.Botao_Download.Height
        .top = (Barra_Transferir.ScaleHeight - .Height) / 2
        .Width = Form_Skin.Botao_Download.Width
        .left = Barra_Transferir.ScaleWidth - .Width - 8
    End With
    
    With Label_Download
        .top = (Botao_Download.ScaleHeight - .Height) / 2
        .Width = Botao_Download.ScaleWidth
        .left = 0
    End With
    
    With ProgressBar1
        .Height = Botao_Download.ScaleHeight
        .top = Botao_Download.top
        .Width = Botao_Download.ScaleWidth
        .left = Botao_Download.left
    End With
    
    With Frame_Foto
        .Height = Form_Skin.Foto_Programa.Height
        .Width = Form_Skin.Foto_Programa.Width
        .left = Barra_Transferir.left + Barra_Transferir.ScaleWidth - .ScaleWidth
    End With
    
    With Shape_Foto
        .Height = Frame_Foto.ScaleHeight
        .top = 0
        .Width = Frame_Foto.ScaleWidth
        .left = 0
    End With
    
    With Image_Tela
        .Stretch = True
        .Height = Frame_Foto.ScaleHeight - 2
        .top = 1
        .Width = Frame_Foto.ScaleWidth - 2
        .left = 1
    End With
    
    With Label_Nenum_Resultado
        .top = 16
        .left = 16
    End With
    
    With Label_Transferir
        .top = (Barra_Transferir.ScaleHeight - .Height) / 2
        .left = 10
    End With
    
    Dim J As Integer
    For J = 0 To Pic_Linha.Count - 1
        With Botao_Mais_Informacoes(J)
            .Height = Form_Skin.Botao_Linha_Normal.Height
            .Width = Form_Skin.Botao_Linha_Normal.Width
            .Visible = False
        End With
        
        With Label_Mais_Informacoes(J)
            .top = Botao_Mais_Informacoes(J).top + ((Botao_Mais_Informacoes(J).Height - .Height) / 2)
            .Width = Botao_Mais_Informacoes(Index).Width
            .left = Botao_Mais_Informacoes(J).left
        End With
        
        With Botao_Remover_Transferencia(J)
            .top = Botao_Mais_Informacoes(J).top
            .Height = Form_Skin.Botao_Linha_2_Normal.Height
            .Width = Form_Skin.Botao_Linha_2_Normal.Width
            .left = Pic_Linha(J).Width - Botao_Remover_Transferencia(J).Width - 8
            .Visible = False
        End With
        
        With Label_Remover_Transferencia(J)
            .top = Botao_Remover_Transferencia(J).top + ((Botao_Remover_Transferencia(J).Height - .Height) / 2)
            .Width = Botao_Remover_Transferencia(Index).Width
            .left = Botao_Remover_Transferencia(J).left
        End With
        
        With Botao_Executar_Programa(J)
            .top = Botao_Mais_Informacoes(J).top
            .Height = Form_Skin.Botao_Linha_2_Normal.Height
            .Width = Form_Skin.Botao_Linha_2_Normal.Width
            .left = Botao_Remover_Transferencia(J).left - .Width - 8
            .Visible = False
        End With
        
        With Label_Executar_Programa(J)
            .top = Botao_Executar_Programa(J).top + ((Botao_Executar_Programa(J).Height - .Height) / 2)
            .Width = Botao_Executar_Programa(Index).Width
            .left = Botao_Executar_Programa(J).left
        End With
    Next J
    
    With Text_Informacao
        .Width = Frame_Informacoes.ScaleWidth - Frame_Foto.ScaleWidth - 40 - 40 - 24
        .left = Image_Logo.left
    End With
    
    Repor_Altura_das_Linhas
    
    'Barra de estado-------------------------------------------------------------------------------
    With Barra_Estado
        .Height = 40 + 24
        .top = Frame_Conteudo.ScaleHeight - .Height
        .Width = Barra_Transferir.ScaleWidth 'Frame_Informacoes.ScaleWidth - SrcrollBar1.Width
        .left = Barra_Transferir.left
    End With
    
    With Shape_Estado
        '.Height = Barra_Estado.ScaleHeight
        .top = 0
        .Width = Barra_Estado.ScaleWidth
        .left = 0
    End With
    
    With Botao_Executar
        .Height = Form_Skin.Botao_Executar.Height
        .top = (Shape_Estado.Height - .Height) / 2
        .Width = Form_Skin.Botao_Executar.Width
        .left = Barra_Estado.ScaleWidth - .Width - 8
    End With
    
    With Label_Executar
        .top = (Botao_Executar.ScaleHeight - .Height) / 2
        .Width = Botao_Executar.ScaleWidth
        .left = 0
    End With
    
    With Botao_Cancelar
        .Height = Botao_Executar.ScaleHeight
        .top = Botao_Executar.top
        .Width = Botao_Executar.ScaleWidth
        .left = Botao_Executar.left - .ScaleWidth - 8
    End With
    
    With Label_Cancelar
        .top = (Botao_Cancelar.ScaleHeight - .Height) / 2
        .Width = Botao_Cancelar.ScaleWidth
        .left = 0
    End With
    
    With Label_Estado
        .top = (Shape_Estado.Height - .Height) / 2
        .left = 32
    End With
    
    'I like -------------------------------------------------------------------------------------------------------
    With Barra_Comentario
        .Height = Form_Skin.Fundo_Separadores.Height
        .Width = Frame_Informacoes.ScaleWidth - 24 - 24
        .left = 24
    End With
    
    With Linha_Barra_Comentario
        .Y1 = 0
        .Y2 = 0
        .X1 = Barra_Comentario.ScaleWidth
        .X2 = 0
    End With

    With Botao_Eu_Gosto
        .Height = Form_Skin.Fundo_Separadores.Height
        .top = Barra_Comentario.top + Barra_Comentario.Height + 10
        .Width = Label_Eu_Gosto.Width + 8 + 10
        .left = Barra_Transferir.left
        Extermidade_Eu_Gosto.left = .ScaleWidth - Extermidade_Eu_Gosto.Width
    End With
        
    With Image_Logo
        .top = 20
        .left = 30
    End With
    
    With Label_Nome_Programa
        .left = Image_Logo.left + Image_Logo.Width + 20
    End With
    
    With Label_Descricao_Programa
        .left = Label_Nome_Programa.left
    End With
    
    With Label_Enterprise
        .left = Label_Nome_Programa.left
    End With
    
    With Label_Total_Downloads
        .left = Label_Nome_Programa.left + Label_Nome_Programa.Width + 10
    End With
    
    'Linhas da lista de programas
    Dim I As Integer
    For I = 0 To Pic_Linha.Count - 1
        Pic_Linha(I).Height = Form_Skin.Linha_Normal.Height
        Label_Nome(I).ForeColor = vbBlack
        Label_Descricao(I).ForeColor = &H808080
    Next I
    
    With Label_Site_Oficial
        .left = Image_Logo.left
    End With
        
    'Ajustar as progressbars da lista de programas
    Dim h As Integer: For h = 0 To Progresso.Count - 1
        Progresso(h).top = Botao_Remover_Transferencia(h).top
        Progresso(h).Height = Form_Skin.Botao_Linha_2_Normal.Height
        Progresso(h).left = Botao_Remover_Transferencia(h).left
        Progresso(h).Width = Form_Skin.Botao_Linha_2_Normal.Width
    Next
    
    With Label_Eu_Gosto
        .top = (Botao_Eu_Gosto.ScaleHeight - .Height) / 2
    End With
    
    With Botao_Update
        .Height = Form_Skin.Button_Menu_Normal.Height
        .top = (Barra_Detalhes.ScaleHeight - .ScaleHeight) / 2
        .Width = Label_Update.Width + 25 + 10 'Form_Skin.Button_Menu_Normal.Width
        .left = Barra_Detalhes.ScaleWidth - .ScaleWidth - 10
    End With
    
    With Image_Update
        .Stretch = True
        .top = 0
        .Height = Botao_Update.ScaleHeight
        .left = 0
        .Width = Botao_Update.ScaleWidth
    End With
    
    With Icon_Update
        .top = (Botao_Update.ScaleHeight - .Height) / 2
        .left = 6
    End With
    
    With Label_Update
        .top = (Botao_Update.ScaleHeight - .Height) / 2
        .left = 25
    End With
    
    'Frame instalados------------------------------------------------------------------------------------------
    With Frame_Instalados
        .top = Frame_Centro.top
        .Height = Frame_Centro.Height
        .left = Frame_Centro.left
        .Width = Frame_Centro.Width
    End With
    
    With Label_Frame_Instalados
        .top = 30
        .left = 30
    End With
    
    With Frame_Icon_Grande
        .top = Label_Frame_Instalados.top + Label_Frame_Instalados.Height + 20
        .Height = Frame_Instalados.ScaleHeight - .top - Label_Frame_Instalados.Height
        .left = Label_Frame_Instalados.left
        .Width = Frame_Instalados.ScaleWidth - (2 * .left)
    End With
    
    With Frame_Icon_Pequeno
        .top = Frame_Icon_Grande.top
        .Height = Frame_Icon_Grande.Height
        .left = Frame_Icon_Grande.left
        .Width = Frame_Icon_Grande.Width
    End With
    
    With Botao_Run
        .Height = Form_Skin.Fundo_Separadores.Height
        .top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .Width = Label_Run.Width + 8 + 8
        .left = 8
        Extermidade_Run.left = .ScaleWidth - Extermidade_Run.Width
    End With
    
    With Label_Run
        .top = (Botao_Run.ScaleHeight - .Height) / 2
        .left = 8
    End With
    
    With Botao_Desinstalar
        .Height = Form_Skin.Fundo_Separadores.Height
        .top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .Width = Label_Desinstalar.Width + 8 + 8
        .left = Botao_Run.left + Botao_Run.ScaleWidth + 8
        Extermidade_Desinstalar.left = .ScaleWidth - Extermidade_Desinstalar.Width
    End With
    
    With Label_Desinstalar
        .top = (Botao_Desinstalar.ScaleHeight - .Height) / 2
        .left = 8
    End With
    '----------------------------------------------------------------------------------------------------------
    
    'Frame partilhar-------------------------------------------------------------------------------------------
    With Frame_Partilhar
        .top = Frame_Centro.top
        .Height = Frame_Centro.Height
        .left = Frame_Centro.left
        .Width = Frame_Centro.Width
    End With
    
    With SrcrollBar2
        .top = 0
        .Height = Frame_Partilhar.ScaleHeight
        .left = Frame_Partilhar.ScaleWidth - .Width
    End With
    
    With Conteudo_Frame_Partilhar
        .top = 0
        '.Height = Frame_Centro.Height
        .left = 0
        '.Width = Frame_Centro.Width
    End With
    
    With Label_Frame_Partilhar
        .top = 30
        .left = 30
    End With
    
    With Botao_Carregar
        .Height = Form_Skin.Fundo_Separadores.Height
        .top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .Width = Label_Carregar.Width + 8 + 8
        .left = 8
        Extermidade_Carregar.left = .ScaleWidth - Extermidade_Carregar.Width
    End With
    
    With Label_Carregar
        .top = (Botao_Carregar.ScaleHeight - .Height) / 2
        .left = 8
    End With
    
    Ajustar_Caixa_Texto Barra_Txt_Email, Txt_Email, Contorno_Txt_Email, False
    Ajustar_Caixa_Texto Barra_Txt_Empresa, Txt_Empresa, Contorno_Txt_Empresa, False
    Ajustar_Caixa_Texto Barra_Txt_Nome, Txt_Nome, Contorno_Txt_Nome, False
    Ajustar_Caixa_Texto Barra_Txt_Descricao, Txt_Descricao, Contorno_Txt_Descricao, False
    Ajustar_Caixa_Texto Barra_Txt_Informacao, Txt_Informacao, Contorno_Txt_Informacao, True
    Ajustar_Caixa_Texto Barra_Txt_Site, Txt_Site, Contorno_Txt_Site, False
    Ajustar_Caixa_Texto Barra_Txt_Download, Txt_Download, Contorno_Txt_Download, False
    
    With Lb_Email
        .top = Label_Frame_Partilhar.top + Label_Frame_Partilhar.Height + 20
        .left = Label_Frame_Partilhar.left
    End With
    
    With Lb_Info
        .top = Lb_Email.top
        .left = Lb_Email.left + Lb_Email.Width + 6
    End With
    
    With Barra_Txt_Email
        .top = Lb_Email.top + Lb_Email.Height + 3
        .left = Label_Frame_Partilhar.left
    End With
    
    With Lb_Empresa
        .top = Barra_Txt_Email.top + Barra_Txt_Email.ScaleHeight + 10
        .left = Label_Frame_Partilhar.left
    End With
    
    With Barra_Txt_Empresa
        .top = Lb_Empresa.top + Lb_Empresa.Height + 3
        .left = Label_Frame_Partilhar.left
    End With
    
    With Lb_Nome
        .top = Barra_Txt_Empresa.top + Barra_Txt_Empresa.ScaleHeight + 10
        .left = Label_Frame_Partilhar.left
    End With
    
    With Barra_Txt_Nome
        .top = Lb_Nome.top + Lb_Nome.Height + 3
        .left = Label_Frame_Partilhar.left
    End With
    
    With Lb_Descricao
        .top = Barra_Txt_Nome.top + Barra_Txt_Nome.ScaleHeight + 10
        .left = Label_Frame_Partilhar.left
    End With

    With Barra_Txt_Descricao
        .top = Lb_Descricao.top + Lb_Descricao.Height + 3
        .left = Label_Frame_Partilhar.left
    End With

    With Lb_Informacao
        .top = Barra_Txt_Descricao.top + Barra_Txt_Descricao.ScaleHeight + 10
        .left = Label_Frame_Partilhar.left
    End With
    
    With Barra_Txt_Informacao
        .top = Lb_Informacao.top + Lb_Informacao.Height + 3
        .left = Label_Frame_Partilhar.left
    End With
    
    With Lb_Site
        .top = Barra_Txt_Informacao.top + Barra_Txt_Informacao.ScaleHeight + 10
        .left = Label_Frame_Partilhar.left
    End With
    
    With Barra_Txt_Site
        .top = Lb_Site.top + Lb_Site.Height + 3
        .left = Label_Frame_Partilhar.left
    End With
    
    With Lb_Download
        .top = Barra_Txt_Site.top + Barra_Txt_Site.ScaleHeight + 10
        .left = Label_Frame_Partilhar.left
    End With
    
    With Barra_Txt_Download
        .top = Lb_Download.top + Lb_Download.Height + 3
        .left = Label_Frame_Partilhar.left
    End With
    
    With Lb_Nota
        .top = Barra_Txt_Download.top + Barra_Txt_Download.ScaleHeight + 3
        .left = Label_Frame_Partilhar.left
    End With
    
    With Frame_Erro_Partilhar
        .Width = Form_Skin.Image_Balao.Width
        .left = Barra_Txt_Email.left + Barra_Txt_Email.ScaleWidth + 30
        .Height = Form_Skin.Image_Balao.Height
        .top = (Barra_Txt_Email.top + (Barra_Txt_Email.ScaleHeight / 2)) - 40
    End With
    '----------------------------------------------------------------------------------------------------------
    
    'Frame opções---------------------------------------------------------------------------------------------
    With Frame_Opcoes
        .top = Frame_Centro.top
        .Height = Frame_Centro.Height
        .left = Frame_Centro.left
        .Width = Frame_Centro.Width
    End With
    
    With Label_Frame_Opcoes
        .top = 30
        .left = 30
    End With
    
    With Label_Opcoes_Actualizadas
        .top = Label_Frame_Opcoes.top + ((Label_Frame_Opcoes.Height - .Height) / 2)
        .left = Label_Frame_Opcoes.left + Label_Frame_Opcoes.Width + 10
    End With
    
    With Botao_Aplicar
        .Height = Form_Skin.Fundo_Separadores.Height
        .top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .Width = Label_Aplicar.Width + 8 + 8
        .left = 8
        Extermidade_Aplicar.left = .ScaleWidth - Extermidade_Aplicar.Width
    End With
    
    With Label_Aplicar
        .top = (Botao_Aplicar.ScaleHeight - .Height) / 2
        .left = 8
    End With
    
    Ajustar_Caixa_Texto Barra_Text_Lingua, Text_Lingua, Contorno_Lingua, False
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
    
    With Shape_Sombra_Lingua(0)
        .Width = Lista_Linguas.ScaleWidth
    End With
    
    With Label_Lingua(0)
        .Width = Lista_Linguas.ScaleWidth
    End With
    
    With Label_Idioma_Programa
        .top = Label_Frame_Opcoes.top + Label_Frame_Opcoes.Height + 20
        .left = Label_Frame_Opcoes.left
    End With
    
    With Barra_Text_Lingua
        .top = Label_Idioma_Programa.top + Label_Idioma_Programa.Height + 3
        .left = Label_Idioma_Programa.left
    End With
    
    Ajustar_ChecBox Me.Pic_Actualizar, Me.Check_Actualizar
    With Check_Actualizar
        .top = Barra_Text_Lingua.top + Barra_Text_Lingua.ScaleHeight + 60
        .left = Label_Frame_Opcoes.left
    End With
    
    With Pic_Actualizar
        .top = Check_Actualizar.top
        .left = Label_Frame_Opcoes.left
    End With
    
    Ajustar_ChecBox Me.Pic_Barra, Me.Check_Barra
    With Check_Barra
        .top = Check_Actualizar.top + Check_Actualizar.Height + 10
        .left = Label_Frame_Opcoes.left
    End With
    
    With Pic_Barra
        .top = Check_Barra.top
        .left = Label_Frame_Opcoes.left
    End With
    
    Ajustar_ChecBox Me.Pic_Tray, Me.Check_Tray
    With Check_Tray
        .top = Check_Barra.top + Check_Barra.Height + 10
        .left = Label_Frame_Opcoes.left
    End With
    
    With Pic_Tray
        .top = Check_Tray.top
        .left = Label_Frame_Opcoes.left
    End With
    '----------------------------------------------------------------------------------------------------------
    
    'Frame suporte técnico-------------------------------------------------------------------------------------
    With Frame_Suporte
        .top = Frame_Centro.top
        .Height = Frame_Centro.Height
        .left = Frame_Centro.left
        .Width = Frame_Centro.Width
    End With
    
    With Label_Frame_Suporte
        .top = 30
        .left = 30
    End With

    Ajustar_Caixa_Texto Barra_Text_Email, Text_Email, Contorno_Email, False
    Ajustar_Caixa_Texto Barra_Text_Assunto, Text_Assunto, Contorno_Assunto, False
    Ajustar_Caixa_Texto Barra_Text_Mensagem, Text_Mensagem, Contorno_Mensagem, True
    
    Dim xpto As Integer: For xpto = 0 To 3
        Shape_Sombra(xpto).Width = Lista_Assunto.ScaleWidth
        Label_Assunto(xpto).Width = Lista_Assunto.ScaleWidth
    Next
    
    With Label_De
        .top = Label_Frame_Suporte.top + Label_Frame_Suporte.Height + 20
        .left = Label_Frame_Suporte.left
    End With
    
    With Label_Info
        .top = Label_De.top
        .left = Label_De.left + Label_De.Width + 3
    End With
    
    With Barra_Text_Email
        .top = Label_De.top + Label_De.Height + 3
        .left = Label_Frame_Suporte.left
    End With
    
    With Label_Texto
        .top = Barra_Text_Email.top + Barra_Text_Email.Height + 10
        .left = Label_Frame_Suporte.left
    End With
    
    With Barra_Text_Assunto
        .top = Label_Texto.top + Label_Texto.Height + 3
        .left = Label_Frame_Suporte.left
    End With
    
    With Label_Mensagem
        .top = Barra_Text_Assunto.top + Barra_Text_Assunto.Height + 10
        .left = Label_Frame_Suporte.left
    End With
    
    With Barra_Text_Mensagem
        .Height = Form_Skin.Caixa_de_Observacoes.Height
        .top = Label_Mensagem.top + Label_Mensagem.Height + 3
        .left = Label_Frame_Suporte.left
        .Width = Form_Skin.Caixa_de_Observacoes.Width
    End With
    
    With Seta_Assunto
        .Height = Form_Skin.Seta_Combo.Height
        .top = (Barra_Text_Assunto.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Seta_Combo.Width
        .left = Barra_Text_Assunto.ScaleWidth - .ScaleWidth - .top
    End With

    With Lista_Assunto
        .top = Barra_Text_Assunto.top + Barra_Text_Assunto.ScaleHeight - 1
        .Width = Barra_Text_Assunto.ScaleWidth
        .left = Barra_Text_Assunto.left
    End With
    
    With Shape_Sombra(0)
        .Width = Lista_Assunto.Width
        .left = 0
    End With
    
    With Botao_Enviar
        .Height = Form_Skin.Fundo_Separadores.Height
        .top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .Width = Label_Enviar.Width + 8 + 8
        .left = 8
        Extermidade_Enviar.left = .ScaleWidth - Extermidade_Enviar.Width
    End With
    
    With Label_Enviar
        .top = (Botao_Enviar.ScaleHeight - .Height) / 2
        .left = 8
    End With
    
    With Botao_Limpar
        .Height = Form_Skin.Fundo_Separadores.Height
        .top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .Width = Label_Limpar.Width + 8 + 8
        .left = Botao_Enviar.left + Botao_Enviar.ScaleWidth + 8
        Extermidade_Limpar.left = .ScaleWidth - Extermidade_Limpar.Width
    End With
    
    With Label_Limpar
        .top = (Botao_Limpar.ScaleHeight - .Height) / 2
        .left = 8
    End With
    
    With Frame_Erro_Suporte
        .Width = Form_Skin.Image_Balao.Width
        .left = Barra_Text_Email.left + Barra_Text_Email.ScaleWidth + 30
        .Height = Form_Skin.Image_Balao.Height
        .top = (Barra_Text_Email.top + (Barra_Text_Email.ScaleHeight / 2)) - 40
    End With
    '----------------------------------------------------------------------------------------------------------
    
    'Frame sobre ----------------------------------------------------------------------------------------------
    With Frame_Sobre
        .top = Frame_Centro.top
        .Height = Frame_Centro.Height
        .left = Frame_Centro.left
        .Width = Frame_Centro.Width
    End With
    
    With Label_Frame_Sobre
        .top = 30
        .left = 30
    End With
    
    With Text_About
        .Height = Frame_Sobre.ScaleHeight - 40
        .top = Label_Frame_Sobre.top + Label_Frame_Sobre.Height + 20
        .Width = Frame_Sobre.ScaleWidth - 60
        .left = 30
    End With
    
    With Botao_Website
        .Height = Form_Skin.Fundo_Separadores.Height
        .top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .Width = Label_Website.Width + 8 + 8
        .left = 8
        Extermidade_Website.left = .ScaleWidth - Extermidade_Website.Width
    End With
    
    With Label_Website
        .top = (Botao_Website.ScaleHeight - .Height) / 2
        .left = 8
    End With
    '----------------------------------------------------------------------------------------------------------
    
Exit Sub
Corrige_Erro:
With Me
    .Height = Altura_Standard
    .top = (Screen.Height - .Height) / 2
    .Width = Largura_Standard
    .left = (Screen.Width - .Width) / 2
End With
End Sub

Private Sub Label_Cancelar_Click()
    'Cancelar a transferência
    Barra_Estado.Visible = False
    dl.cancel
    ProgressBar1.Visible = False
    ProgressBar1.Value = 0
    Label_Download.Caption = Idioma_Button_Transfer_Program
    
    Botao_Download.Visible = True
    Barra_Estado.Visible = False
    'Label_Titulo.Caption = "Center programs Nikyts"
    Label_Estado.Caption = ReadINI("Main", "Operation_Canceled", Localizacao_Ficheiro_Lingua)
    Me.MousePointer = 0
End Sub

Private Sub Label_Categorias_Click()
    'Ver a frame lista de programas
    If Frame_Lista.Visible = True Then Exit Sub
    'If Label_Categorias.Caption = "Sugestão" Then Label_Programas_Click: Exit Sub
    
    Extermidade_Categorias.Picture = Form_Skin.Extermidade_Normal.Picture
    Ocultar_Objectos
    Frame_Lista.Visible = True
    
    Separador_Informacoes.Visible = False
    Barra_Estado.Visible = False
    
    Tab_Informacao_Visivel = False
End Sub

Private Sub Label_Descricao_Click(Index As Integer)
    'Atalho para
    Pic_Linha_Click (Index)
End Sub

Private Sub Label_Download_Click()
    'Transferir o programa selecionado
    Select Case Label_Download.Caption
        Case Idioma_Button_Transfer_Program
            ProgressBar1.Visible = True
            Botao_Executar.Enabled = False
            Label_Executar.Enabled = False
            Botao_Cancelar.Enabled = True
            Label_Cancelar.Enabled = True
        
            Image_Download.Picture = Form_Skin.Image_Down_Processando.Picture
            Label_Estado.Caption = Idioma_Transferring_File
            Barra_Estado.Visible = True
            Me.MousePointer = 11
            'Label_Titulo.Caption = "Centro de programas Nikyts - Aguarde..."
            Verificar_Pastas
            Botao_Download.Visible = False
            Text_Servidor.Text = "http://www.nikyts.com/gadgets/programas/" & Label_Transferir.Caption
            dl.DownloadFile Text_Servidor.Text, App.Path & "\Programs\" & Label_Transferir.Caption '& GetFileName(Label_Transferir.Caption)
            On Error GoTo 0 'Tratamento de erros
            
        '------------------------------------------------------------------------------------------------------
        Case Idioma_Button_Remove_Program
            Me.MousePointer = 11
            'Remover a pasta, sub-pasta e respectivos ficheiros referentes ao programa
            DeleteFolderTree App.Path & "\Programs\" & Label_Nome_Programa.Caption
            Label_Transferir.Caption = Label_Nome_Programa.Caption & ".zip"
            Label_Download.Caption = Idioma_Button_Transfer_Program
            Barra_Estado.Visible = False
            Recarregar_Programas_Instalados
    End Select
    Me.MousePointer = 0
    
Exit Sub
errHand:
Me.MousePointer = 0
End Sub

Private Sub dl_DowloadComplete()
    'Transferência concluida
    On Error GoTo Corrige_Erro
    GetFileName (Text_Servidor.Text)
    ProgressBar1.Value = 0
    GetFileName (Text_Servidor.Text)
    
    Botao_Download.Visible = True
    ProgressBar1.Visible = False
    Me.MousePointer = 0
    Label_Estado.Caption = ReadINI("Main", "Label_Download_Complete", Localizacao_Ficheiro_Lingua)
    Botao_Executar.Enabled = True
    Label_Executar.Enabled = True
    Botao_Cancelar.Enabled = False
    Label_Cancelar.Enabled = False
    
    'Actualiza no servidor nº de downloads do programa
    Verificar_Downloads
    
    'Iniciar a decompactação do programa zipado
    DesCompacta App.Path & "\Programs\" & Label_Transferir.Caption, "*.*", App.Path & "\Programs\", True
    Kill App.Path & "\Programs\" & Label_Transferir.Caption
    
    'Ao terminar a transferência do ficheiro a Idioma_Button_Transfer_Program passa a ser Idioma_Button_Remove_Program
    Label_Download.Caption = Idioma_Button_Remove_Program
    Label_Transferir.Caption = ReadINI("Main", "Label_Installed_In", Localizacao_Ficheiro_Lingua) & ": " & Date & " " & Time
    Barra_Estado.Visible = True
    
    'Actualizar a data e hora de criação do programa
    Dim Ficheiro_Para_Actualizar As String
    Ficheiro_Para_Actualizar = App.Path & "\Programs\" & Label_Nome_Programa.Caption & "\" & Label_Nome_Programa.Caption & ".exe"
    
    'Set the creation time
    FileSetDate Ficheiro_Para_Actualizar, Now, True
    'Set the last accessed time
    FileSetDate Ficheiro_Para_Actualizar, Now, , True
    'Set the last write time
    FileSetDate Ficheiro_Para_Actualizar, Now, , , True
    
    Image_Download.Picture = Form_Skin.Image_Down_Concluido.Picture
    Recarregar_Programas_Instalados
    Me.MousePointer = 0
    
Exit Sub
Corrige_Erro:
End Sub

Private Sub dl_DownloadErrors(strError As String)
    'Caso ocorra um erro durante o download
    Label_Estado.Caption = ReadINI("Main", "Error_Transfer_Program", Localizacao_Ficheiro_Lingua)
    Label_Download.Caption = Idioma_Button_Transfer_Program
    Botao_Download.Visible = True
    ProgressBar1.Visible = False
    Image_Download.Picture = Form_Skin.Image_Down_Erro.Picture
    Me.MousePointer = 0
End Sub

Private Sub dl_DownloadProgress(intPercent As String)
    'Mostrar o progresso do download
    ProgressBar1.Value = intPercent
    GetFileName (Text_Servidor.Text)
    Text_Servidor.Text = ""
    Label_Estado.Caption = ReadINI("Main", "Label_Transferring_File", Localizacao_Ficheiro_Lingua)
    
    Image_Download.Picture = Form_Skin.Image_Down_Processando.Picture
End Sub

Private Sub Label_Executar_Click()
    'Executar o programa automaticamente
    On Error GoTo Corrige_Erro
    Shell App.Path & "\Programs\" & Label_Nome_Programa.Caption & "\" & Label_Nome_Programa.Caption & ".exe"
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case -2146697211
        Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada
        
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub


Private Sub Label_Informacoes_Click()
    'Ver a frame informações do programa
    
End Sub

Private Sub Label_Remover_Transferencia_Click(Index As Integer)
    'Remover a pasta, sub-pasta e respectivos ficheiros referentes ao programa
    Me.MousePointer = 11
    Select Case Label_Remover_Transferencia(Index).Caption
        Case Idioma_Button_Transfer_Program
            Me.MousePointer = 11
            Verificar_Pastas
            Botao_Remover_Transferencia(Index).Visible = False
            Label_Remover_Transferencia(Index).Visible = False
            
            'Proceder á transferência dos respectivos programas
            Linha_Programa_Selecionado = Label_Remover_Transferencia(Index).Index
            Text_Servidor.Text = "http://www.nikyts.com/gadgets/programas/" & Label_Programa(Index).Caption
            progress_activo = Index
            Progresso(progress_activo).Visible = True
            Download_Programa.DownloadFile Text_Servidor.Text, App.Path & "\Programs\" & Label_Programa(Index).Caption
            On Error GoTo 0 'Tratamento de erros
        
        
        Case Idioma_Button_Remove_Program
            DeleteFolderTree App.Path & "\Programs\" & Label_Nome(Index).Caption
            Label_Remover_Transferencia(Index).Caption = Idioma_Button_Transfer_Program
            Botao_Executar_Programa(Index).Enabled = False
            Label_Executar_Programa(Index).Enabled = False
            
            Recarregar_Programas_Instalados
    End Select
    Me.MousePointer = 0
    
Exit Sub
errHand:
Me.MousePointer = 0
End Sub

Private Sub Label_Mais_Informacoes_Click(Index As Integer)
    'Ver informações do programa
    Selecionar_Programa Label_Nome(Index), Label_Descricao(Index), Label_Programa(Index), Label_Downloads(Index), Label_Observacoes(Index), _
        Label_Icon(Index), Label_Logotipo(Index), Label_Tela(Index), Label_Avaliacao(Index), Label_Id(Index), Label_site(Index), Label_Empresa(Index)
        
    'Carregar o logotipo e respectiva tela do programa
    Image_Logo.Picture = Logotipo_Programa(Index).Picture
    Image_Tela.Picture = Tela_Programa(Index).Picture
    
    With Label_Total_Downloads
        .left = Label_Nome_Programa.left + Label_Nome_Programa.Width + 10
    End With
    
    Verificar_Se_Programa_Existe
End Sub

Private Sub Label_Menu_Pesquisar_Click()
    'Atalho para
    Menu_Pesquisar_Click
End Sub

Private Sub Label_Menu_Sobre_Click()
    'Atalho para
    Menu_Sobre_Click
End Sub

Private Sub Label_Menu_Suporte_Click()
    'Atalho para
    Menu_Suporte_Click
End Sub

Private Sub Label_Nome_Click(Index As Integer)
    'Atalho para
    Pic_Linha_Click (Index)
End Sub

Private Sub Label_Programas_Click()
    'Atalho para
    Frame_Centro.backcolor = Azul
    Ocultar_Frames
    Repor_Menus
    Ocultar_Separadores
    Frame_Centro.Visible = True
    Menu_Pesquisar.Picture = Form_Skin.Icon_Pesquisar_Down.Picture
    Label_Menu_Pesquisar.ForeColor = Azul
    
    Me.MousePointer = 0
    Extermidade_Programas.Picture = Form_Skin.Extermidade_Normal.Picture
    Separador_Programas.Visible = True
    Barra_Caixa_Pesquisa.Visible = True
    
    Ocultar_Objectos
    Frame_Pesquisar.Visible = True
    
    Tab_Categoria_Visivel = False
    Tab_Informacao_Visivel = False
End Sub

Private Sub Label_Titulo_DblClick()
    'Maximixar/ Restaurar Formulários
    If Tela_Cheia = True Then
        Botao_Restaurar_Click
    Else
        Botao_Maximizar_Click
    End If
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Principal
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    If Tela_Cheia = False Then Mover_Formulario Form_Principal
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Principal
    Actualizar_Valores
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Capturar_Posicao_Formulario Form_Principal
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    If Tela_Cheia = False Then Mover_Formulario Form_Principal
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    Largar_Formulario Form_Principal
    Actualizar_Valores
End Sub

Private Sub Barra_ControlBox_DblClick()
    'Atalho para
    Label_Titulo_DblClick
End Sub

Public Sub Selecionar_Categoria(Categoria_Selecionada As String, texto_categoria As String)
    'Procedimento para escolher a categoria do programa
    Barra_Estado.Visible = False
    Repor_Altura_das_Linhas
    Ocultar_Objectos
    Frame_Centro.backcolor = vbWhite
    Formatar_Lista_Programas
    
    Label_Categorias.Caption = texto_categoria
    With Separador_Categorias
        .Height = Form_Skin.Fundo_Separadores.Height
        .left = Separador_Programas.left + Separador_Programas.ScaleWidth
        .Width = Label_Categorias.Width + 8 + 15
        Extermidade_Categorias.Picture = Form_Skin.Extermidade_Normal.Picture
        Extermidade_Categorias.left = .ScaleWidth - Extermidade_Categorias.Width
        .Visible = True
    End With
    Extermidade_Programas.Picture = Form_Skin.Extermidade_Over.Picture
    Tab_Categoria_Visivel = True
            
    Categoria_a_ser_Pesquisada = Categoria_Selecionada
    Carregar_Programas Categoria_Selecionada
End Sub

Public Sub Verificar_Pastas()
    'Procedimento para verificar se as pastas utilizadas pelo programa existem
    If Not ArquivoExiste(App.Path & "\Programs", True) Then
        MkDir App.Path & "\Programs\"
    End If
End Sub

Public Sub Selecionar_Programa(Label_Nome As Label, Label_Descricao As Label, Label_Programa As Label, Label_Downloads As Label, _
                                Label_Observacoes As Label, Label_Icon As Label, Label_Logotipo As Label, Label_Tela As Label, _
                                Label_Avaliacao As Label, Label_Id As Label, Label_site As Label, Label_Empresa As Label)
    'Procedimento para escolher a categoria do programa
    Ocultar_Objectos
    
    Me.MousePointer = 11
    Extermidade_Categorias.Picture = Form_Skin.Extermidade_Over.Picture
        
    Label_Informacoes.Caption = Label_Nome.Caption
    With Separador_Informacoes
        .Height = Form_Skin.Fundo_Separadores.Height
        .left = Separador_Categorias.left + Separador_Categorias.ScaleWidth
        .Width = Label_Informacoes.Width + 8 + 15
        Extermidade_Categorias.Picture = Form_Skin.Extermidade_Over.Picture
        Extermidade_Informacoes.left = .ScaleWidth - Extermidade_Informacoes.Width
        .Visible = True
    End With
    Tab_Informacao_Visivel = True
    
    Label_Nome_Programa.Caption = Label_Nome.Caption
    Label_Descricao_Programa.Caption = Label_Descricao.Caption
    
    Label_Transferencias.Caption = Label_Downloads.Caption
    If Val(Label_Downloads.Caption) = 1 Then
        Label_Total_Downloads.Caption = "(" & Label_Downloads.Caption & " download)"
    Else
        Label_Total_Downloads.Caption = "(" & Label_Downloads.Caption & " downloads)"
    End If
    
    Label_Transferir.Caption = Label_Programa.Caption
    Text_Informacao.Text = Label_Observacoes.Caption
    
    'Avaliação do programa, Estrelas
    If Val(Label_Avaliacao.Caption) < 20 Then
        Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_0.Picture
    
    ElseIf Val(Label_Avaliacao.Caption) >= 20 And Val(Label_Avaliacao.Caption) < 40 Then
        Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_1.Picture
    
    ElseIf Val(Label_Avaliacao.Caption) >= 40 And Val(Label_Avaliacao.Caption) < 60 Then
        Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_2.Picture
    
    ElseIf Val(Label_Avaliacao.Caption) >= 60 And Val(Label_Avaliacao.Caption) < 80 Then
        Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_3.Picture
    
    ElseIf Val(Label_Avaliacao.Caption) >= 80 And Val(Label_Avaliacao.Caption) < 100 Then
        Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_4.Picture
    
    ElseIf Val(Label_Avaliacao.Caption) > 100 Then
        Frame_Avaliacao.Picture = Form_Skin.Image_Estrelas_5.Picture
    End If
    
    'Total de avaliações
    Label_Votos.Caption = Label_Avaliacao.Caption
    If Label_Avaliacao.Caption = "1" Then
        Label_Total.Caption = Idioma_Label_Rate & ": " & Label_Votos.Caption
    Else
        Label_Total.Caption = Idioma_Label_Rate & ": " & Label_Votos.Caption
    End If
    
    'Receber o id do programa para depois obter os comentários sobre o mesmo
    Label_Id_Programa.Caption = Label_Id
    
    Label_Site_Programa.Caption = Label_site
    If Label_Site_Programa.Caption = Empty Then
        Label_Site_Oficial.Visible = False
    Else
        Label_Site_Oficial.Visible = True
    End If
    
    'Nome da empresa/programador
    Label_Enterprise.Caption = Label_Empresa.Caption
    
    'Carregar_Comentarios Label_Id_Programa.Caption
    Verificar_Se_Programa_Existe
    
    'Visualizar a frame da informação detalhada do programa
    Frame_Conteudo.Visible = True
    Me.MousePointer = 0
End Sub

Private Sub Menu_Pesquisar_Click()
    'Ver a frame das categorias
    If Frame_Centro.Visible = True Then Exit Sub
    Ocultar_Frames
    Repor_Menus
    Ocultar_Separadores
    Frame_Centro.Visible = True
    Menu_Pesquisar.Picture = Form_Skin.Icon_Pesquisar_Down.Picture
    Label_Menu_Pesquisar.ForeColor = Azul
    
    Me.MousePointer = 0
    If Tab_Categoria_Visivel = True Then
        Extermidade_Programas.Picture = Form_Skin.Extermidade_Over.Picture
    Else
        Extermidade_Programas.Picture = Form_Skin.Extermidade_Normal.Picture
    End If
    Separador_Programas.Visible = True
    Barra_Caixa_Pesquisa.Visible = True
    
    If Tab_Categoria_Visivel = True Then Separador_Categorias.Visible = True
    If Tab_Informacao_Visivel = True Then Separador_Informacoes.Visible = True
End Sub

Private Sub Menu_Sobre_Click()
    'Ver formulário sobre
    If Frame_Sobre.Visible = True Then Exit Sub
    Ocultar_Frames
    Repor_Menus
    Ocultar_Separadores
    Frame_Sobre.Visible = True
    Menu_Sobre.Picture = Form_Skin.Icon_Sobre_Down.Picture
    Label_Menu_Sobre.ForeColor = Azul
    Botao_Website.Visible = True
End Sub

Private Sub Menu_Suporte_Click()
    'Ver form suporte
    On Error Resume Next
    If Frame_Suporte.Visible = True Then Exit Sub
    Ocultar_Frames
    Repor_Menus
    Ocultar_Separadores
    Frame_Suporte.Visible = True
    Menu_Suporte.Picture = Form_Skin.Icon_Suporte_Down.Picture
    Label_Menu_Suporte.ForeColor = Azul
    Botao_Enviar.Visible = True
    Botao_Limpar.Visible = True
    Text_Email.SetFocus
End Sub

Private Sub Ocultar_Separadores()
    'Procedimento para des/activar separadores
    Separador_Programas.Visible = False
    Separador_Categorias.Visible = False
    Separador_Informacoes.Visible = False
    Botao_Enviar.Visible = False
    Botao_Limpar.Visible = False
    Barra_Caixa_Pesquisa.Visible = False
    Botao_Website.Visible = False
    Botao_Aplicar.Visible = False
    Botao_Run.Visible = False
    Botao_Desinstalar.Visible = False
    Botao_Carregar.Visible = False
End Sub

Private Sub Ocultar_Frames()
    'Procedimento para ocultar as frames
    Frame_Centro.Visible = False
    Frame_Instalados.Visible = False
    Frame_Partilhar.Visible = False
    Frame_Opcoes.Visible = False
    Frame_Suporte.Visible = False
    Frame_Sobre.Visible = False
    
    Label_Opcoes_Actualizadas.Visible = False
    Ocultar_Listas
End Sub

Private Sub Repor_Menus()
    'Procedimento para repor as imagens originais dos menus
    With Form_Skin
        Menu_Pesquisar.Picture = .Icon_Pesquisar_Normal.Picture
        Menu_Instalados.Picture = .Icon_Instalados_Normal.Picture
        Menu_Partilhar.Picture = .Icon_Partilhar_Normal.Picture
        Menu_Opcoes.Picture = .Icon_Opcoes_Normal.Picture
        Menu_Suporte.Picture = .Icon_Suporte_Normal.Picture
        Menu_Sobre.Picture = .Icon_Sobre_Normal.Picture
        
        Label_Menu_Pesquisar.ForeColor = vbWhite
        Label_Menu_Instalados.ForeColor = vbWhite
        Label_Menu_Partilhar.ForeColor = vbWhite
        Label_Menu_Opcoes.ForeColor = vbWhite
        Label_Menu_Suporte.ForeColor = vbWhite
        Label_Menu_Sobre.ForeColor = vbWhite
    End With
End Sub

Private Sub Pic_Linha_Click(Index As Integer)
    'Selecionar a linha
    If Pic_Linha(Index).Height = Form_Skin.Linha_Normal.Height Then
        Repor_Altura_das_Linhas
        Pic_Linha(Index).backcolor = Azul 'Azul
        Botao_Mais_Informacoes(Index).Picture = Form_Skin.Botao_Linha_Over.Picture
        Botao_Remover_Transferencia(Index).Picture = Form_Skin.Botao_Linha_2_Over.Picture
        Botao_Executar_Programa(Index).Picture = Form_Skin.Botao_Linha_2_Over.Picture
        Pic_Linha(Index).Height = Form_Skin.Linha_Over.Height
        Botao_Mais_Informacoes(Index).Visible = True
        Label_Mais_Informacoes(Index).Visible = True
        Botao_Remover_Transferencia(Index).Visible = True
        Label_Remover_Transferencia(Index).Visible = True
        Botao_Executar_Programa(Index).Visible = True
        Label_Executar_Programa(Index).Visible = True
        Label_Nome(Index).ForeColor = vbWhite
        Label_Descricao(Index).ForeColor = vbWhite
        
    Else
        Pic_Linha(Index).Height = Form_Skin.Linha_Normal.Height
        Pic_Linha(Index).backcolor = &HF9F9F9    'Branco
        Botao_Mais_Informacoes(Index).Picture = Form_Skin.Botao_Linha_Normal.Picture
        Botao_Remover_Transferencia(Index).Picture = Form_Skin.Botao_Linha_2_Normal.Picture
        Botao_Executar_Programa(Index).Picture = Form_Skin.Botao_Linha_2_Normal.Picture
        Botao_Mais_Informacoes(Index).Visible = False
        Label_Mais_Informacoes(Index).Visible = False
        Botao_Remover_Transferencia(Index).Visible = False
        Label_Remover_Transferencia(Index).Visible = False
        Botao_Executar_Programa(Index).Visible = False
        Label_Executar_Programa(Index).Visible = False
        Label_Nome(Index).ForeColor = vbBlack
        Label_Descricao(Index).ForeColor = &H808080
    End If
    
    Ajustar_Linha_Lista_Programas
End Sub

Private Sub Separador_Categorias_Click()
    'Atalho para
    Label_Categorias_Click
End Sub

Private Sub Separador_Informacoes_Click()
    'Atalho para
    Label_Informacoes_Click
End Sub

Private Sub Separador_Programas_Click()
    'Atalho para
    Label_Programas_Click
End Sub

Public Sub Ocultar_Objectos()
    'Procedimento para ocultar objectos não pretendidos
    Frame_Pesquisar.Visible = False
    Frame_Lista.Visible = False
    Frame_Conteudo.Visible = False
End Sub

Public Sub Carregar_Programas(Categoria_Selecionada As String)
    'Efectuar pesquisa na base de dados consuante os dados introduzidos
    'On Error GoTo Corrige_Erro
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.open "GET", "http://www.nikyts.com/gadgets/" & "carregarprogramas.asp?Recebe_Categoria=" & Categoria_a_ser_Pesquisada
    servidor.send 'envia o pedido para o servidor
    
    'Verificar os dados acesso
    If servidor.responseText = "NaoExiste" Then
        Label_Nenum_Resultado.Visible = True
        'Exit Sub
        
    ElseIf Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 Then
            Me.MousePointer = 11
        
            Formatar_Lista_Programas
            Ajustar_Linha_Lista_Programas
            Extermidade_Programas.Picture = Form_Skin.Extermidade_Over.Picture
        
            Formatar_Lista_Programas
            Servidor_Carregar_Programas servidor.responseText
            Me.MousePointer = 0
            Frame_Lista.Visible = True
        End If
    End If
    Set servidor = Nothing
    Ocultar_Objectos
    Frame_Lista.Visible = True
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Ocultar_Objectos
Frame_Lista.Visible = True
Label_Nenum_Resultado.Visible = True
Select Case err.Number
    Case -2146697211
        Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada

    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub Servidor_Carregar_Programas(responseText As String)
    'Procedimento para ler os dados do xml referente aos dados da galeria do utilizador
    Dim X As Integer: X = 0
    Dim verticalGap As Integer: verticalGap = 15
    Dim xml As MSXML2.DOMDocument: Set xml = New MSXML2.DOMDocument
    
    If xml.loadXML(responseText) Then
        Dim node As IXMLDOMNode
        Dim nodeList As IXMLDOMNodeList
        Set nodeList = xml.selectNodes("/meusprogramas/resultado")
        Dim I As Integer: I = 0
        
        For Each node In nodeList
            DoEvents
            
            'Caso exista + do que 1 programa para listar
            If I > 0 Then 'nodeList.length
                Load Pic_Linha(I)
                Pic_Linha(I).Move Pic_Linha(I - 1).left, Pic_Linha(I - 1).top + Pic_Linha(I - 1).Height
                Pic_Linha(I).Visible = True
                
                Load Icon_Programa(I)
                Icon_Programa(I).Move Icon_Programa(0).left, Pic_Linha(I).top + Icon_Programa(0).top
                Icon_Programa(I).Visible = True
                
                Load Label_Nome(I)
                Label_Nome(I).Move Label_Nome(0).left, Pic_Linha(I).top + Label_Nome(0).top
                Label_Nome(I).Visible = True
                
                Load Label_Descricao(I)
                Label_Descricao(I).Move Label_Descricao(0).left, Pic_Linha(I).top + Label_Descricao(0).top
                Label_Descricao(I).Visible = True
                
                Load Progresso(I)
                Progresso(I).Move Label_Descricao(0).left, Pic_Linha(I).top + Progresso(0).top
                Progresso(I).Visible = False
                
                Load Botao_Mais_Informacoes(I)
                Botao_Mais_Informacoes(I).Move Botao_Mais_Informacoes(0).left, Pic_Linha(I).top + Botao_Mais_Informacoes(0).top
                Botao_Mais_Informacoes(I).Visible = False
                
                Load Label_Mais_Informacoes(I)
                Label_Mais_Informacoes(I).Move Label_Mais_Informacoes(0).left, Pic_Linha(I).top + Label_Mais_Informacoes(0).top
                Label_Mais_Informacoes(I).Visible = False
                Label_Mais_Informacoes(I).ZOrder 0
                
                Load Botao_Executar_Programa(I)
                Botao_Executar_Programa(I).Move Botao_Executar_Programa(0).left, Pic_Linha(I).top + Botao_Executar_Programa(0).top
                Botao_Executar_Programa(I).Visible = False
                
                Load Label_Executar_Programa(I)
                Label_Executar_Programa(I).Move Label_Executar_Programa(0).left, Pic_Linha(I).top + Label_Executar_Programa(0).top
                Label_Executar_Programa(I).Visible = False
                Label_Executar_Programa(I).ZOrder 0
                
                Load Botao_Remover_Transferencia(I)
                Botao_Remover_Transferencia(I).Move Botao_Remover_Transferencia(0).left, Pic_Linha(I).top + Botao_Remover_Transferencia(0).top
                Botao_Remover_Transferencia(I).Visible = False
                
                Load Label_Remover_Transferencia(I)
                Label_Remover_Transferencia(I).Move Label_Remover_Transferencia(0).left, Pic_Linha(I).top + Label_Remover_Transferencia(0).top
                Label_Remover_Transferencia(I).Visible = False
                Label_Remover_Transferencia(I).ZOrder 0
                
                Pic_Linha(I).ZOrder 1
                
                'Objectos ocultos
                Load Label_Programa(I)
                'Load Label_Nome(i)
                'Load Label_Descricao(i)
                Load Label_Downloads(I)
                Load Label_Observacoes(I)
                Load Label_Icon(I)
                Load Label_Logotipo(I)
                Load Label_Tela(I)
                Load Label_Avaliacao(I)
                Load Label_Id(I)
                Load Label_site(I)
                Load Label_Empresa(I)
                
                'Load Icon_Programa(i)
                Load Logotipo_Programa(I)
                Load Tela_Programa(I)
                
            End If
            
            'Carregar a informação fornecida pelo servidor
            If Not IsEmpty(node.selectSingleNode("programa")) Then Label_Programa(I).Caption = node.selectSingleNode("programa").Text
            If Not IsEmpty(node.selectSingleNode("titulo")) Then Label_Nome(I).Caption = node.selectSingleNode("titulo").Text
            If Not IsEmpty(node.selectSingleNode("descricao")) Then Label_Descricao(I).Caption = node.selectSingleNode("descricao").Text
            If Not IsEmpty(node.selectSingleNode("downloads")) Then Label_Downloads(I).Caption = node.selectSingleNode("downloads").Text
            If Not IsEmpty(node.selectSingleNode("observacoes")) Then Label_Observacoes(I).Caption = node.selectSingleNode("observacoes").Text
            If Not IsEmpty(node.selectSingleNode("icon")) Then Label_Icon(I).Caption = node.selectSingleNode("icon").Text
            If Not IsEmpty(node.selectSingleNode("logotipo")) Then Label_Logotipo(I).Caption = node.selectSingleNode("logotipo").Text
            If Not IsEmpty(node.selectSingleNode("tela")) Then Label_Tela(I).Caption = node.selectSingleNode("tela").Text
            If Not IsEmpty(node.selectSingleNode("avaliacao")) Then Label_Avaliacao(I).Caption = node.selectSingleNode("avaliacao").Text
            If Not IsEmpty(node.selectSingleNode("id")) Then Label_Id(I).Caption = node.selectSingleNode("id").Text
            If Not IsEmpty(node.selectSingleNode("site")) Then Label_site(I).Caption = node.selectSingleNode("site").Text
            If Not IsEmpty(node.selectSingleNode("empresa")) Then Label_Empresa(I).Caption = node.selectSingleNode("empresa").Text
            
            'Carregar o icon, logotipo e tela do programa
            If Label_Icon(I).Caption <> Empty Then Set Icon_Programa(I).Picture = LoadPicture(Label_Icon(I).Caption) '"http://www.nikyts.com/gadgets/imagens/" &
            If Label_Logotipo(I).Caption <> Empty Then Set Logotipo_Programa(I).Picture = LoadPicture(Label_Logotipo(I).Caption)
            If Label_Tela(I).Caption <> Empty Then Set Tela_Programa(I).Picture = LoadPicture(Label_Tela(I).Caption)
            Pic_Linha(I).Visible = True
            Icon_Programa(I).Visible = True
            Label_Nome(I).Visible = True
            Label_Descricao(I).Visible = True
            
            '------------------------------------------------------------------------------------------------------------------------
            'Verificar se as pastas utilizadas pelo programa existem
            ficheiro = App.Path & "\Programs\" & Label_Nome(I).Caption & "\" '& Label_Nome_Programa.Caption & ".exe"
            If ArquivoExiste(ficheiro, True) Then
'                Label_Remover_Transferencia(i).Caption = Idioma_Button_Remove_Program
                sVar = DataArq(ficheiro & Label_Nome(I).Caption & ".exe")
                If sVar <> "ERRO" Then
                    'Label_Transferir.Caption = ReadINI("Main", "Label_Installed_In", Localizacao_Ficheiro_Lingua) & ": " & sVar
                    Botao_Executar_Programa(I).Enabled = True
                    Label_Executar_Programa(I).Enabled = True
                    'Botao_Remover_Transferencia(i).Enabled = True
                    'Label_Remover_Transferencia(i).Enabled = True
                    Label_Remover_Transferencia(I).Caption = Idioma_Button_Remove_Program
                Else
                    'Label_Transferir.Caption = Label_Nome(i).Caption & ".zip"
                    Botao_Executar_Programa(I).Enabled = False
                    Label_Executar_Programa(I).Enabled = False
                    'Botao_Remover_Transferencia(i).Enabled = False
                    'Label_Remover_Transferencia(i).Enabled = False
                    Label_Remover_Transferencia(I).Caption = Idioma_Button_Transfer_Program
                End If
            Else
                Botao_Executar_Programa(I).Enabled = False
                Label_Executar_Programa(I).Enabled = False
                'Botao_Remover_Transferencia(i).Enabled = False
                'Label_Remover_Transferencia(i).Enabled = False
                Label_Remover_Transferencia(I).Caption = Idioma_Button_Transfer_Program
            End If
            '------------------------------------------------------------------------------------------------------------------------
            
            I = I + 1
        Next
    End If
End Sub

Public Sub Formatar_Lista_Programas()
    'Procedimento para limpar todos os campos da lista de programas
    Dim I As Integer: For I = 0 To Pic_Linha.Count - 1
        Pic_Linha(I).Visible = False
        Pic_Linha(I).Height = Form_Skin.Linha_Normal.Height
        
        Botao_Mais_Informacoes(I).Visible = False
        Label_Mais_Informacoes(I).Visible = False
        Botao_Executar_Programa(I).Visible = False
        Label_Executar_Programa(I).Visible = False
        Botao_Remover_Transferencia(I).Visible = False
        Label_Remover_Transferencia(I).Visible = False
        
        Icon_Programa(I).Visible = False
        Icon_Programa(I).Picture = Form_Skin.Imagem_Vazia.Picture
        Logotipo_Programa(I).Picture = Form_Skin.Imagem_Vazia.Picture
        Tela_Programa(I).Picture = Form_Skin.Imagem_Vazia.Picture
        
        Label_Programa(I).Caption = Empty
        Label_Nome(I).Visible = False
        Label_Nome(I).Caption = Empty
        Label_Descricao(I).Visible = False
        Label_Descricao(I).Caption = Empty
        Label_Downloads(I).Caption = Empty
        Label_Observacoes(I).Caption = Empty
        Label_Icon(I).Caption = Empty
        Label_Logotipo(I).Caption = Empty
        Label_Tela(I).Caption = Empty
        Label_Avaliacao(I).Caption = Empty
        Label_Id(I).Caption = Empty
        Label_site(I).Caption = Empty
        
        Label_Nome(I).ForeColor = vbBlack
        Label_Descricao(I).ForeColor = &H808080
    
        'Ocultar as progressbars
        Progresso(I).Visible = False
    Next I
    
    Label_Nenum_Resultado.Visible = False
    
    'Apagar as restantes linhas e respectivos objectos
    If Pic_Linha.Count > 1 Then
        Dim J As Integer: For J = 1 To Pic_Linha.Count - 1
            Unload Pic_Linha(J)
            Unload Icon_Programa(J)
            Unload Label_Nome(J)
            Unload Label_Descricao(J)
            Unload Progresso(J)
            Unload Botao_Mais_Informacoes(J)
            Unload Label_Mais_Informacoes(J)
            Unload Botao_Executar_Programa(J)
            Unload Label_Executar_Programa(J)
            Unload Botao_Remover_Transferencia(J)
            Unload Label_Remover_Transferencia(J)
            Unload Label_Programa(J)
            Unload Label_Downloads(J)
            Unload Label_Observacoes(J)
            Unload Label_Icon(J)
            Unload Label_Logotipo(J)
            Unload Label_Tela(J)
            Unload Label_Avaliacao(J)
            Unload Label_Id(J)
            Unload Label_site(J)
            Unload Label_Empresa(J)
            Unload Logotipo_Programa(J)
            Unload Tela_Programa(J)
        Next
    End If
End Sub

Private Sub SrcrollBar1_Change()
    'Posicionar as linhas da lista conforme a posição da scroll
    Me.Frame_Informacoes.top = 0 - Me.SrcrollBar1.Value
End Sub

Private Sub SrcrollBar1_Scroll(Value As Long)
    'Posicionar as linhas da lista conforme a posição da scroll
    Me.Frame_Informacoes.top = -Value
End Sub

Private Sub SrcrollBar2_Change()
    'Posicionar as linhas da lista conforme a posição da scroll
    Me.Conteudo_Frame_Partilhar.top = 0 - Me.SrcrollBar2.Value
End Sub

Private Sub SrcrollBar2_Scroll(Value As Long)
    'Posicionar as linhas da lista conforme a posição da scroll
    Me.Conteudo_Frame_Partilhar.top = -Value
End Sub

Private Sub Text_Assunto_Click()
    'Ocultar lista
    Lista_Assunto.Visible = False
End Sub

Private Sub Text_Assunto_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Assunto.BorderColor = Azul
End Sub

Private Sub Text_Assunto_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Assunto.BorderColor = Cinza
End Sub

Private Sub Text_Email_Click()
    'Ocultar lista
    Lista_Assunto.Visible = False
End Sub

Private Sub Text_Email_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Email.BorderColor = Azul
End Sub

Private Sub Text_Email_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Email.BorderColor = Cinza
End Sub

Private Sub Text_Lingua_Click()
    'Ocultar lista
    Lista_Linguas.Visible = False
End Sub

Private Sub Text_Lingua_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Lingua.BorderColor = Azul
End Sub

Private Sub Text_Lingua_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Lingua.BorderColor = Cinza
End Sub

Private Sub Text_Mensagem_Click()
    'Ocultar lista
    Lista_Assunto.Visible = False
End Sub

Private Sub Text_Mensagem_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Mensagem.BorderColor = Azul
End Sub

Private Sub Text_Mensagem_LostFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Mensagem.BorderColor = Cinza
End Sub

Private Sub Text_Pesquisa_GotFocus()
    'Ao receber o focus
    Text_Pesquisa.Text = Empty
    Text_Pesquisa.ForeColor = vbBlack
    Contorno_Caixa_Pesquisa.BorderColor = Azul
End Sub

Private Sub Text_Pesquisa_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Pesquisar_Click
End Sub

Private Sub Text_Pesquisa_LostFocus()
    'Ao perder o focus
    Text_Pesquisa.Text = ReadINI("Main", "Text_Search", Localizacao_Ficheiro_Lingua)
    Text_Pesquisa.ForeColor = &H808080
    Contorno_Caixa_Pesquisa.BorderColor = Cinza
End Sub

Public Sub Ajustar_Linha_Lista_Programas()
    'Procedimento para ajustar as linhas da lista dos programas
    Dim Linha, Altura As Integer
    Linha = 0: Altura = 0
    For Linha = 0 To Pic_Linha.Count - 1
        Pic_Linha(Linha).top = Altura
        Altura = Altura + Pic_Linha(Linha).Height
        
        Icon_Programa(Linha).top = Pic_Linha(Linha).top + Icon_Programa(0).top
        Label_Nome(Linha).top = Pic_Linha(Linha).top + Label_Nome(0).top
        Label_Descricao(Linha).top = Pic_Linha(Linha).top + Label_Descricao(0).top
        Botao_Mais_Informacoes(Linha).top = Pic_Linha(Linha).top + Botao_Mais_Informacoes(0).top
        Label_Mais_Informacoes(Linha).top = Pic_Linha(Linha).top + Label_Mais_Informacoes(0).top
        Botao_Executar_Programa(Linha).top = Pic_Linha(Linha).top + Botao_Executar_Programa(0).top
        Label_Executar_Programa(Linha).top = Pic_Linha(Linha).top + Label_Executar_Programa(0).top
        Botao_Remover_Transferencia(Linha).top = Pic_Linha(Linha).top + Botao_Remover_Transferencia(0).top
        Label_Remover_Transferencia(Linha).top = Pic_Linha(Linha).top + Label_Remover_Transferencia(0).top
    Next Linha
End Sub

Public Sub Repor_Altura_das_Linhas()
    'Procedimento para repor a altura de todas as linhas da lista de programas
    Dim Linha As Integer
    Linha = 0
    For Linha = 0 To Pic_Linha.Count - 1
        With Pic_Linha(Linha)
            .Height = Form_Skin.Linha_Normal.Height
            .backcolor = &HF9F9F9 'Branco
            Botao_Mais_Informacoes(Linha).Visible = False
            Label_Mais_Informacoes(Linha).Visible = False
            Botao_Remover_Transferencia(Linha).Visible = False
            Label_Remover_Transferencia(Linha).Visible = False
            Botao_Executar_Programa(Linha).Visible = False
            Label_Executar_Programa(Linha).Visible = False
            Label_Nome(Linha).ForeColor = vbBlack
            Label_Descricao(Linha).ForeColor = &H808080
        End With
    Next Linha
End Sub

Private Sub Txt_Email_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Txt_Email.BorderColor = Azul
End Sub

Private Sub Txt_Email_LostFocus()
    'Ao perder o focus da caixa de texto
    Contorno_Txt_Email.BorderColor = Cinza
End Sub

Private Sub Txt_Empresa_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Txt_Empresa.BorderColor = Azul
End Sub

Private Sub Txt_Empresa_LostFocus()
    'Ao perder o focus da caixa de texto
    Contorno_Txt_Empresa.BorderColor = Cinza
End Sub

Private Sub Txt_Nome_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Txt_Nome.BorderColor = Azul
End Sub

Private Sub Txt_Nome_LostFocus()
    'Ao perder o focus da caixa de texto
    Contorno_Txt_Nome.BorderColor = Cinza
End Sub

Private Sub Txt_Descricao_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Txt_Descricao.BorderColor = Azul
End Sub

Private Sub Txt_Descricao_LostFocus()
    'Ao perder o focus da caixa de texto
    Contorno_Txt_Descricao.BorderColor = Cinza
End Sub

Private Sub Txt_Informacao_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Txt_Informacao.BorderColor = Azul
End Sub

Private Sub Txt_Informacao_LostFocus()
    'Ao perder o focus da caixa de texto
    Contorno_Txt_Informacao.BorderColor = Cinza
End Sub

Private Sub Txt_Site_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Txt_Site.BorderColor = Azul
End Sub

Private Sub Txt_Site_LostFocus()
    'Ao perder o focus da caixa de texto
    Contorno_Txt_Site.BorderColor = Cinza
End Sub

Private Sub Txt_Download_GotFocus()
    'Ao receber o focus na caixa de texto
    Contorno_Txt_Download.BorderColor = Azul
End Sub

Private Sub Txt_Download_LostFocus()
    'Ao perder o focus da caixa de texto
    Contorno_Txt_Download.BorderColor = Cinza
End Sub

Private Sub txtZip_Change()
    'Indicação de progresso da compactação/descompactação por arquivo
    '----------------------------------------------------------------
    'Tipo de ação que esta sendo feita no momento
    lblProgresso = TipoAção(Val(GetAction(txtZip.Text))) & " "
    'Nome do arquivo que esta sendo compactado
    lblProgresso = lblProgresso & GetFileName(txtZip.Text) & " -> "
    'Porcentagem de compactação do arquivo
    lblProgresso = lblProgresso & GetPercentComplete(txtZip.Text) & "%"
    'Força a atualização da tela
    DoEvents
End Sub

Public Sub Verificar_Downloads()
    'Procedimento para verificar o total de downloads de cada programa
    'On Error GoTo Corrige_Erro
    If Label_Transferencias.Caption = "" Then Exit Sub
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    
    'Adicionar um voto á avaliação do programa
    Label_Transferencias.Caption = Val(Label_Transferencias.Caption) + 1
    servidor.open "GET", "http://www.nikyts.com/gadgets/" & "actualizardownloads.asp?id_programa=" & Label_Id_Programa.Caption & "&downloads=" & Label_Transferencias.Caption, False
    servidor.send 'envia o pedido para o servidor

    '"http://www.nikyts.com/gadgets/actualizardownloads.asp?id_programa=" + idPrograma
    
'    'Actualizar a senha
'    If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
'        With Form_Principal
'            If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then ' 4 - deu resposta e 200 validou
'                'Adicionar + 1 download ao total de dpwnloads do programa
'                If Val(Label_Transferencias.Caption) = 1 Then
'                    Label_Total_Downloads.Caption = "(" & Label_Transferencias.Caption & " download)"
'                Else
'                    Label_Total_Downloads.Caption = "(" & Label_Transferencias.Caption & " downloads)"
'                End If
'            End If
'        End With
'    End If
    Set servidor = Nothing
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case -2146697211
        Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada
        
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Public Sub Verificar_Se_Programa_Existe()
    On Error Resume Next
    'Procedimento para verificar se o programa em questão já foi instalado
    Dim ficheiro, sVar As String
    ficheiro = App.Path & "\Programs\" & Label_Nome_Programa.Caption & "\" '& Label_Nome_Programa.Caption & ".exe"
    
    'Procedimento para verificar se as pastas utilizadas pelo programa existem
    If ArquivoExiste(ficheiro, True) Then
        Label_Download.Caption = Idioma_Button_Remove_Program
            
        sVar = DataArq(ficheiro & Label_Nome_Programa.Caption & ".exe")
        If sVar <> "ERRO" Then
            Label_Transferir.Caption = ReadINI("Main", "Label_Installed_In", Localizacao_Ficheiro_Lingua) & ": " & sVar
            Barra_Estado.Visible = True
            Botao_Cancelar.Enabled = False
            Label_Cancelar.Enabled = False
            
        Else
            Label_Transferir.Caption = Label_Nome_Programa.Caption & ".zip"
            Label_Download.Caption = Idioma_Button_Transfer_Program
            Barra_Estado.Visible = False
        End If
    
    Else
        Label_Download.Caption = Idioma_Button_Transfer_Program
        Barra_Estado.Visible = False
    End If
End Sub

Public Sub DeleteFolderTree(ByVal vFolder As String)
    'Procedimento para eliminar a pasta, sub-pastas e respectivos ficheiros referentes ao programa
    Dim FSO As FileSystemObject
    Dim FoldersObj As Folders
    Dim FolderObj As Folder
    Set FSO = New FileSystemObject
    
    If Not FSO.FolderExists(vFolder) Then
    Set FSO = Nothing
    Exit Sub
    End If
    
    Set FolderObj = FSO.GetFolder(vFolder)
    Set FoldersObj = FolderObj.SubFolders
    For Each FolderObj In FoldersObj
    DeleteFolderTree FolderObj.Path
    Next FolderObj
    On Error Resume Next
    
    Kill vFolder & "\*.*"
    RmDir vFolder
    
    err.Clear
    On Error GoTo 0
    
    Set FolderObj = Nothing
    Set FoldersObj = Nothing
    Set FSO = Nothing
End Sub

Private Sub Download_Programa_DowloadComplete()
    'Transferência concluida
    On Error GoTo Corrige_Erro
    GetFileName (Text_Servidor.Text)
    Progresso(Linha_Programa_Selecionado).Value = 0
    GetFileName (Text_Servidor.Text)
    
    Botao_Remover_Transferencia(Linha_Programa_Selecionado).Visible = True
    Botao_Executar_Programa(Linha_Programa_Selecionado).Enabled = True
    Label_Executar_Programa(Linha_Programa_Selecionado).Enabled = True
    Progresso(progress).Visible = False
    
    'Actualiza no servidor nº de downloads do programa
    Label_Transferencias.Caption = Val(Label_Downloads(Linha_Programa_Selecionado).Caption) + 1
    Label_Id_Programa.Caption = Label_Downloads(Linha_Programa_Selecionado).Caption
    Verificar_Downloads
    
    'Iniciar a decompactação do programa zipado
    DesCompacta App.Path & "\Programs\" & Label_Programa(Linha_Programa_Selecionado).Caption, "*.*", App.Path & "\Programs\", True
    Kill App.Path & "\Programs\" & Label_Programa(Linha_Programa_Selecionado).Caption
    
    'Ao terminar a transferência do ficheiro a Idioma_Button_Transfer_Program passa a ser Idioma_Button_Remove_Program
    Label_Remover_Transferencia(Linha_Programa_Selecionado).Caption = Idioma_Button_Remove_Program
    
    'Actualizar a data e hora de criação do programa
    Dim Ficheiro_Para_Actualizar As String
    Ficheiro_Para_Actualizar = App.Path & "\Programs\" & Label_Nome(Linha_Programa_Selecionado).Caption & "\" & Label_Nome(Linha_Programa_Selecionado).Caption & ".exe"
    
    'Set the creation time
    FileSetDate Ficheiro_Para_Actualizar, Now, True
    'Set the last accessed time
    FileSetDate Ficheiro_Para_Actualizar, Now, , True
    'Set the last write time
    FileSetDate Ficheiro_Para_Actualizar, Now, , , True
    
    Recarregar_Programas_Instalados
    Me.MousePointer = 0
    
Exit Sub
Corrige_Erro:
Me.MousePointer = 0
Select Case err.Number
    Case -2146697211
        Mensagem_de_Aviso "Error", Idioma_Conectar_Servidor & vbNewLine & Idioma_Internet_Desligada

    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Error", Idioma_Erro_Execucao & vbNewLine & Idioma_Erro & " " & err.Number & vbNewLine & Idioma_Descricao & " " & err.Description
End Select
End Sub

Private Sub Download_Programa_DownloadErrors(strError As String)
    'Caso ocorra um erro durante o download
    Label_Remover_Transferencia(Linha_Programa_Selecionado).Caption = Idioma_Button_Transfer_Program
    Botao_Remover_Transferencia(Linha_Programa_Selecionado).Visible = True
    Botao_Executar_Programa(Linha_Programa_Selecionado).Enabled = False
    Label_Executar_Programa(Linha_Programa_Selecionado).Enabled = False
    Progresso(progress_activo).Visible = False
    
    Mensagem_de_Aviso "Error", ReadINI("Main", "Error_Transfer_Program", Localizacao_Ficheiro_Lingua)
    Me.MousePointer = 0
End Sub

Private Sub Download_Programa_DownloadProgress(intPercent As String)
    'Mostrar o progresso do download
    Progresso(Linha_Programa_Selecionado).Value = intPercent
    GetFileName (Text_Servidor.Text)
    Text_Servidor.Text = ""
End Sub

Private Sub Ocultar_Listas()
    'Procedimento para ocultar objectos (ex. as listas das comboboxs)
    Lista_Linguas.Visible = False
    Lista_Assunto.Visible = False
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
    Dim logotipo_normal, logotipo_over As String
    
    Botao_Run.Enabled = False: Label_Run.Enabled = False
    Botao_Desinstalar.Enabled = False: Label_Desinstalar.Enabled = False
    
    If Lista_Pastas.ListCount <> 0 Then
        Botao_Run.Enabled = True: Label_Run.Enabled = True
        Botao_Desinstalar.Enabled = True: Label_Desinstalar.Enabled = True

        Lista_Pastas.ListIndex = 0
        'Logo normal
        logotipo_normal = App.Path & "\Programs\" & Lista_Pastas.List(0) & "\Options\Logo_Normal.jpg"
        logotipo_over = App.Path & "\Programs\" & Lista_Pastas.List(0) & "\Options\Logo_Over.jpg"
        
        If ArquivoExiste(logotipo_normal, False) And ArquivoExiste(logotipo_over, False) Then
            Image_Icon_Grande(0).Picture = LoadPicture(logotipo_normal)
            Image_Logo_Normal(0).Picture = LoadPicture(logotipo_normal)
            Image_Logo_Over(0).Picture = LoadPicture(logotipo_over)
        Else
            Image_Icon_Grande(0).Picture = Form_Skin.Icon_Grande_Normal.Picture
            Image_Logo_Normal(0).Picture = Form_Skin.Icon_Grande_Normal.Picture
            Image_Logo_Over(0).Picture = Form_Skin.Icon_Grande_Over.Picture
        End If
        
        Image_Icon_Grande(0).Visible = True
        Label_Icon_Grande(0).Caption = Lista_Pastas.List(I)
        Label_Icon_Grande(0).Visible = True
        
        'Restantes...
        Dim Objecto As Integer: For Objecto = 1 To Lista_Pastas.ListCount - 1
            Load Image_Icon_Grande(Objecto)
            Image_Icon_Grande(Objecto).Move Image_Icon_Grande(Objecto - 1).left + Image_Icon_Grande(Objecto - 1).Width + 20, Image_Icon_Grande(Objecto - 1).top
            
            logotipo_normal = App.Path & "\Programs\" & Lista_Pastas.List(Objecto) & "\Options\Logo_Normal.jpg"
            logotipo_over = App.Path & "\Programs\" & Lista_Pastas.List(Objecto) & "\Options\Logo_Over.jpg"
        
            Load Image_Logo_Normal(Objecto)
            Load Image_Logo_Over(Objecto)
            
            If ArquivoExiste(logotipo_normal, False) And ArquivoExiste(logotipo_over, False) Then
                Image_Icon_Grande(Objecto).Picture = LoadPicture(logotipo_normal)
                Image_Logo_Normal(Objecto).Picture = LoadPicture(logotipo_normal)
                Image_Logo_Over(Objecto).Picture = LoadPicture(logotipo_over)
            Else
                Image_Icon_Grande(Objecto).Picture = Form_Skin.Icon_Grande_Normal.Picture
                Image_Logo_Normal(Objecto).Picture = Form_Skin.Icon_Grande_Normal.Picture
                Image_Logo_Over(Objecto).Picture = Form_Skin.Icon_Grande_Over.Picture
            End If
            Image_Icon_Grande(Objecto).Visible = True
            
            Load Label_Icon_Grande(Objecto)
            Label_Icon_Grande(Objecto).Move Image_Icon_Grande(Objecto).left, Label_Icon_Grande(Objecto - 1).top
            Label_Icon_Grande(Objecto).Caption = Lista_Pastas.List(Objecto)
            Label_Icon_Grande(Objecto).Visible = True
            Label_Icon_Grande(Objecto).ZOrder 0
        Next
        
        If Check_Barra.Value = 1 Then Form_Barra.Show
    End If
End Sub

Public Sub Actualizar_Valores()
    'Procedimento para actualizar as dimensões e posições do formulário
    If Tela_Cheia = False Then
        With Me
            Text_Form_Top = .top
            Text_Form_Height = .Height
            Text_Form_Left = .left
            Text_Form_Width = .Width
        End With
    End If
End Sub
