VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menu"
   ClientHeight    =   8580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8580
   ScaleWidth      =   14130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "Damero"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   4815
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   12960
      Top             =   3120
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E67E22&
      Caption         =   "Alumno"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   3375
      Begin VB.Label Label1 
         BackColor       =   &H00E67E22&
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir del Programa"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1440
      Picture         =   "Form1.frx":22581
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E74C3C&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   6120
      TabIndex        =   6
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    End 'salir del programa
End Sub

Private Sub Command2_Click()
    Form1.Hide 'oculta el formulario 1
    Load Form2 'caraga el formulario 2
    Form2.Show 'muestra el formulario 2
End Sub

Private Sub Form_Load()
    'se colocan los textos de los label 1 y 4
    Label1.Caption = "Nombre: Daniel Grazzina" & vbNewLine _
        & "Cedula: 26.254.611" & vbNewLine & "Seccion: V-0401" & vbNewLine
    Label4.Caption = "Presiona el boton para ir al juego" & vbNewLine
End Sub

Private Sub Timer1_Timer()
    Label2.Caption = Date 'este comado coloca la fecha en el label
    Label3.Caption = Time 'este comando coloca la hora en el label
End Sub
