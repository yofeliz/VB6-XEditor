VERSION 5.00
Begin VB.Form frmAcerca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de..."
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcerca.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmProgramador 
      Height          =   2535
      Left            =   3840
      TabIndex        =   6
      Top             =   120
      Width           =   3615
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Mayo 2005"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1260
         MouseIcon       =   "frmAcerca.frx":0442
         TabIndex        =   15
         Top             =   1800
         Width           =   1050
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "GALICIA (ESPAÑA)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   855
         MouseIcon       =   "frmAcerca.frx":074C
         TabIndex        =   14
         Top             =   1560
         Width           =   1830
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "As Pontes (A Coruña)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         MouseIcon       =   "frmAcerca.frx":0A56
         TabIndex        =   13
         Top             =   1320
         Width           =   2100
      End
      Begin VB.Label lblProgramacion 
         AutoSize        =   -1  'True
         Caption         =   "Programación:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label lblProgramador 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "David Díaz"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1260
         MouseIcon       =   "frmAcerca.frx":0D60
         TabIndex        =   8
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label lblLocalización 
         AutoSize        =   -1  'True
         Caption         =   "Localización:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Frame frmPrograma 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmAcerca.frx":106A
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblWindowsNT2000XP 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Windows NT/2000/XP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   750
         TabIndex        =   12
         Top             =   2040
         Width           =   2130
      End
      Begin VB.Label lblWindows9x 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Windows 95/98/ME"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   855
         TabIndex        =   11
         Top             =   1800
         Width           =   1920
      End
      Begin VB.Label lblCompatible 
         AutoSize        =   -1  'True
         Caption         =   "Compatible con:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   1395
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         Caption         =   "X-Editor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   840
         TabIndex        =   5
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "v.0.0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2160
         TabIndex        =   4
         Top             =   360
         Width           =   510
      End
      Begin VB.Label lblProgramado 
         AutoSize        =   -1  'True
         Caption         =   "Programado en:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label lblVB6SP6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Visual Basic 6.0 SP6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   2
         Top             =   1080
         Width           =   1950
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
End
Attribute VB_Name = "frmAcerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Versión " + CStr(App.Major) + "." + CStr(App.Minor)
End Sub

