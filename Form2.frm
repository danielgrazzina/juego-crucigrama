VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C000&
   Caption         =   "Damero"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14700
   LinkTopic       =   "Form2"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FINALIZAR JUEGO"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   151
      Top             =   9720
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reiniciar juego"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   148
      Top             =   8520
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reglas"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   240
      TabIndex        =   120
      Top             =   240
      Width           =   3255
      Begin VB.Label Label25 
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   240
         TabIndex        =   121
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   94
      Left            =   15120
      TabIndex        =   95
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   93
      Left            =   14520
      TabIndex        =   94
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   92
      Left            =   13920
      TabIndex        =   93
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   91
      Left            =   13320
      TabIndex        =   92
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   90
      Left            =   12720
      TabIndex        =   91
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   89
      Left            =   12120
      TabIndex        =   90
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   88
      Left            =   11520
      TabIndex        =   89
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   87
      Left            =   10920
      TabIndex        =   88
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   86
      Left            =   10320
      TabIndex        =   87
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   85
      Left            =   9720
      TabIndex        =   86
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   84
      Left            =   9120
      TabIndex        =   85
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   83
      Left            =   8520
      TabIndex        =   84
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   82
      Left            =   7920
      TabIndex        =   83
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   81
      Left            =   7320
      TabIndex        =   82
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   80
      Left            =   6720
      TabIndex        =   81
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   79
      Left            =   6120
      TabIndex        =   80
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   78
      Left            =   5520
      TabIndex        =   79
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   77
      Left            =   4920
      TabIndex        =   78
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   76
      Left            =   4320
      TabIndex        =   77
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   75
      Left            =   15120
      TabIndex        =   76
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   74
      Left            =   14520
      TabIndex        =   75
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   73
      Left            =   13920
      TabIndex        =   74
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   72
      Left            =   13320
      TabIndex        =   73
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   71
      Left            =   12720
      TabIndex        =   72
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   70
      Left            =   12120
      TabIndex        =   71
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   69
      Left            =   11520
      TabIndex        =   70
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   68
      Left            =   10920
      TabIndex        =   69
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   67
      Left            =   10320
      TabIndex        =   68
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   66
      Left            =   9720
      TabIndex        =   67
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   65
      Left            =   9120
      TabIndex        =   66
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   64
      Left            =   8520
      TabIndex        =   65
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   63
      Left            =   7920
      TabIndex        =   64
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   62
      Left            =   7320
      TabIndex        =   63
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   61
      Left            =   6720
      TabIndex        =   62
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   60
      Left            =   6120
      TabIndex        =   61
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   59
      Left            =   5520
      TabIndex        =   60
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   58
      Left            =   4920
      TabIndex        =   59
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   57
      Left            =   4320
      TabIndex        =   58
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   56
      Left            =   15120
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   55
      Left            =   14520
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   54
      Left            =   13920
      TabIndex        =   55
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   53
      Left            =   13320
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   52
      Left            =   12720
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   51
      Left            =   12120
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   50
      Left            =   11520
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   49
      Left            =   10920
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   48
      Left            =   10320
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   47
      Left            =   9720
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   46
      Left            =   9120
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   45
      Left            =   8520
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   44
      Left            =   7920
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   43
      Left            =   7320
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   42
      Left            =   6720
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   41
      Left            =   6120
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   40
      Left            =   5520
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   39
      Left            =   4920
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   38
      Left            =   4320
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   37
      Left            =   15120
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   36
      Left            =   14520
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   35
      Left            =   13920
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   34
      Left            =   13320
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   33
      Left            =   12720
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   32
      Left            =   12120
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   31
      Left            =   11520
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   30
      Left            =   10920
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   29
      Left            =   10320
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   28
      Left            =   9720
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   27
      Left            =   9120
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   26
      Left            =   8520
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   25
      Left            =   7920
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   24
      Left            =   7320
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   23
      Left            =   6720
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   22
      Left            =   6120
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   21
      Left            =   5520
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   20
      Left            =   4920
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   19
      Left            =   4320
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   18
      Left            =   15120
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   17
      Left            =   14520
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   16
      Left            =   13920
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   15
      Left            =   13320
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   14
      Left            =   12720
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   13
      Left            =   12120
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   12
      Left            =   11520
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   11
      Left            =   10920
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   10
      Left            =   10320
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   9
      Left            =   9720
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   8
      Left            =   9120
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   7
      Left            =   8520
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   6
      Left            =   7920
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   5
      Left            =   7320
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   4
      Left            =   6720
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   3
      Left            =   6120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   2
      Left            =   5520
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   1
      Left            =   4920
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      Height          =   615
      Index           =   0
      Left            =   4320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   615
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
      Left            =   720
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18840
      TabIndex        =   150
      Top             =   9960
      Width           =   495
   End
   Begin VB.Label Label52 
      Caption         =   "Label50"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12840
      TabIndex        =   149
      Top             =   9960
      Width           =   5655
   End
   Begin VB.Label Label51 
      Alignment       =   2  'Center
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   17760
      TabIndex        =   147
      Top             =   9120
      Width           =   495
   End
   Begin VB.Label Label50 
      Caption         =   "Label50"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12840
      TabIndex        =   146
      Top             =   9120
      Width           =   4575
   End
   Begin VB.Label Label49 
      Alignment       =   2  'Center
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   17760
      TabIndex        =   145
      Top             =   8160
      Width           =   495
   End
   Begin VB.Label Label48 
      Caption         =   "Label48"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12840
      TabIndex        =   144
      Top             =   8160
      Width           =   4575
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   17760
      TabIndex        =   143
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label Label46 
      Caption         =   "Label46"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12840
      TabIndex        =   142
      Top             =   7200
      Width           =   4575
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   17760
      TabIndex        =   141
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label Label44 
      Caption         =   "Label44"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12840
      TabIndex        =   140
      Top             =   6240
      Width           =   4575
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   19680
      TabIndex        =   139
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label42 
      Caption         =   "Label42"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12840
      TabIndex        =   138
      Top             =   5280
      Width           =   6495
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   17760
      TabIndex        =   137
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label40 
      Caption         =   "Label40"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12840
      TabIndex        =   136
      Top             =   4320
      Width           =   4575
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   135
      Top             =   9960
      Width           =   495
   End
   Begin VB.Label Label38 
      Caption         =   "Label38"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   134
      Top             =   9960
      Width           =   4215
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   133
      Top             =   9000
      Width           =   495
   End
   Begin VB.Label Label36 
      Caption         =   "Label36"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   132
      Top             =   9000
      Width           =   4215
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   131
      Top             =   8040
      Width           =   495
   End
   Begin VB.Label Label34 
      Caption         =   "Label34"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   130
      Top             =   8040
      Width           =   4215
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   129
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label32 
      Caption         =   "Label32"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   128
      Top             =   7080
      Width           =   4215
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   127
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label Label30 
      Caption         =   "Label30"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   126
      Top             =   6120
      Width           =   4215
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   125
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label Label28 
      Caption         =   "Label28"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   124
      Top             =   5160
      Width           =   4215
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      TabIndex        =   123
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label26 
      Caption         =   "Label26"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   122
      Top             =   4320
      Width           =   6975
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   119
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   118
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   117
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   116
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   115
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15120
      TabIndex        =   114
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14520
      TabIndex        =   113
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      TabIndex        =   112
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13320
      TabIndex        =   111
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12720
      TabIndex        =   110
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      TabIndex        =   109
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   108
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   107
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   106
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   105
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   104
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   103
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   102
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   101
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   100
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   99
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   98
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   97
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   96
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim reinicio As Integer

Private Sub Command1_Click() 'boton para regresar al fomulario 1
    Form2.Hide
    Load Form1
    Form1.Show
End Sub

Private Sub Command2_Click() 'boton para reiniciar el juego
    reinicio = 1
End Sub

Private Sub Command3_Click() 'boton que comprueba que finalisaste correctamente el juego
    If (Label27.Caption = "C" And Label29.Caption = "C") And (Label31.Caption = "C" And Label33.Caption = "C") Then
        If (Label35.Caption = "C" And Label37.Caption = "C") And (Label39.Caption = "C" And Label41.Caption = "C") Then
            If (Label43.Caption = "C" And Label45.Caption = "C") And (Label47.Caption = "C" And Label49.Caption = "C") Then
                If Label51.Caption = "C" And Label53.Caption = "C" Then
                    MsgBox "felicidades completaste el damero"
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    
    reinicio = 1 'al iniciar el formulario se coloca la variable reinicio en 1 para que se inicie el tablero
    'los siguentes comandos muestran en los label las preguntas con las codenadas correspondientes
    Label25.Caption = "1. Escriba las letras en mayuscula en el recuadro con las cordenadas correspondientes" & vbNewLine _
        & vbNewLine & "2. Si todas las letras de su respuesta estan correctas saldra una c al lado de la pregunta si no saldra una i" & vbNewLine
    
    Label26.Caption = "1. Dirigente mas destacado del movimiento de independencia indio" & vbNewLine _
        & "     |5A|5B|5C|5D|5E|5F|5G|  |5I|5J|5K|5L|5M|5N|" & vbNewLine
        
    Label28.Caption = "2. Continente" & vbNewLine _
        & "     |1R|1H|1P|2F|2G|1F|2B|" & vbNewLine
        
    Label30.Caption = "3. Contrario de morir" & vbNewLine _
        & "     |1A|1L|1C|4B|1Q|" & vbNewLine
        
    Label32.Caption = "4. Animal de peluche" & vbNewLine _
        & "     |IG|1K|1I|" & vbNewLine
        
    Label34.Caption = "5. No son las tardes" & vbNewLine _
        & "     |2J|2K|2L|2M|2N|2O|1T|" & vbNewLine
        
    Label36.Caption = "6. El fuma y yo" & vbNewLine _
        & "     |1N|1O|2D|2E" & vbNewLine
        
    Label38.Caption = "7. Lugar donde vivieron Adan y Eva" & vbNewLine _
        & "     |3A|3C|3D|3B|" & vbNewLine
        
    Label40.Caption = "8. Secrecion nasal" & vbNewLine _
        & "     |3H|3I|3F|3G|" & vbNewLine
        
    Label42.Caption = "9. Comedia satirica de Will Smith de los años 90 El Rey Del" & vbNewLine _
        & "     |2H|2Q|2R|" & vbNewLine
        
    Label44.Caption = "10. Dejar sin cabello" & vbNewLine _
        & "     |4I|4J|4G|4H|4E|" & vbNewLine
        
    Label46.Caption = "11. Antonimo de muera" & vbNewLine _
        & "     |4C|4D|4A|3R|" & vbNewLine
        
    Label48.Caption = "12. Sinonimo de continuamente" & vbNewLine _
        & "     |4L|4M|4N|4O|4P|4Q|4R|" & vbNewLine
        
    Label50.Caption = "13. Conjugacion del verbo ser en plural" & vbNewLine _
        & "     |3K|3L|3M|3N|3O|3P|" & vbNewLine
        
    Label52.Caption = "14. La tercera persona del presente del verbo reir" & vbNewLine _
        & "     |2T|1B|1D|" & vbNewLine
End Sub

Private Sub Timer1_Timer() 'el timer se usa para comprobar que las respuestas sean correctas y para reiniciar el tablero
    
    If reinicio = 1 Then 'al estar la variable reinicio en 1 configura el tablero a su estado innicial
        For x = 0 To 94 ' con el bucle for se vacia el tablero y se coloca el text segun sea requerido
            If x < 4 Or (x > 4 And x < 9) Then
                Text1(x) = ""
                Text1(x).MaxLength = 1
                Text1(x).Font.Name = "MV Boli"
                Text1(x).Font.Size = 24
                Text1(x).Alignment = 2
            Else
                If (x = 4 Or x = 9) Or (x = 12 Or x = 19) Then
                    Text1(x).BackColor = black
                    Text1(x) = ""
                    Text1(x).Enabled = False
                Else
                    If (x > 9 And x < 12) Or (x > 12 And x < 19) Then
                        Text1(x) = ""
                        Text1(x).MaxLength = 1
                        Text1(x).Font.Name = "MV Boli"
                        Text1(x).Font.Size = 24
                        Text1(x).Alignment = 2
                    Else
                        If x = 20 Or (x > 21 And x < 27) Then
                            Text1(x) = ""
                            Text1(x).MaxLength = 1
                            Text1(x).Font.Name = "MV Boli"
                            Text1(x).Font.Size = 24
                            Text1(x).Alignment = 2
                        Else
                            If (x = 21 Or x = 27) Or (x = 34 Or x = 42) Then
                                Text1(x).BackColor = black
                                Text1(x) = ""
                                Text1(x).Enabled = False
                            Else
                                If (x > 27 And x < 34) Or (x > 34 And x < 42) Then
                                    Text1(x) = ""
                                    Text1(x).MaxLength = 1
                                    Text1(x).Font.Name = "MV Boli"
                                    Text1(x).Font.Size = 24
                                    Text1(x).Alignment = 2
                                Else
                                    If (x = 47 Or x = 54) Or (x = 56 Or x = 62) Then
                                        Text1(x).BackColor = black
                                        Text1(x) = ""
                                        Text1(x).Enabled = False
                                    Else
                                        If (x > 42 And x < 47) Or (x > 47 And x < 54) Then
                                            Text1(x) = ""
                                            Text1(x).MaxLength = 1
                                            Text1(x).Font.Name = "MV Boli"
                                            Text1(x).Font.Size = 24
                                            Text1(x).Alignment = 2
                                        Else
                                            If (x > 54 And x < 56) Or (x > 56 And x < 62) Then
                                                Text1(x) = ""
                                                Text1(x).MaxLength = 1
                                                Text1(x).Font.Name = "MV Boli"
                                                Text1(x).Font.Size = 24
                                                Text1(x).Alignment = 2
                                            Else
                                                If (x = 67 Or x = 75) Or (x = 83 Or x > 89) Then
                                                    Text1(x).BackColor = black
                                                    Text1(x) = ""
                                                    Text1(x).Enabled = False
                                                Else
                                                    Text1(x) = ""
                                                    Text1(x).MaxLength = 1
                                                    Text1(x).Font.Name = "MV Boli"
                                                    Text1(x).Font.Size = 24
                                                    Text1(x).Alignment = 2
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next x
        reinicio = 0
        Label27.Caption = "i"
        Label29.Caption = "i"
        Label31.Caption = "i"
        Label33.Caption = "i"
        Label35.Caption = "i"
        Label37.Caption = "i"
        Label39.Caption = "i"
        Label41.Caption = "i"
        Label43.Caption = "i"
        Label45.Caption = "i"
        Label47.Caption = "i"
        Label49.Caption = "i"
        Label51.Caption = "i"
        Label53.Caption = "i"
    End If
    'se comprueba que cada letra de cada pregunta sea correcta
    'pregunta 1
    If (Text1(76) = "M" And Text1(77) = "A") And (Text1(78) = "H" And Text1(79) = "A") Then
        If (Text1(80) = "T" And Text1(81) = "M") And (Text1(82) = "A" And Text1(84) = "G") Then
            If (Text1(85) = "A" And Text1(86) = "N") And (Text1(87) = "D" And Text1(88) = "H") Then
                If Text1(89) = "I" Then
                    Label27.Caption = "C"
                Else
                    Label27.Caption = "i"
                End If
            End If
        End If
    End If
    'pregunta 2
    If (Text1(17) = "A" And Text1(7) = "M") And (Text1(15) = "E" And Text1(24) = "R") Then
        If (Text1(25) = "I" And Text1(5) = "C") And Text1(20) = "A" Then
            Label29.Caption = "C"
            ganar = 1
        Else
            Label29.Caption = "i"
        End If
    End If
    'pregunta 3
    If (Text1(0) = "V" And Text1(11) = "I") And (Text1(2) = "V" And Text1(58) = "I") Then
        If Text1(16) = "R" Then
            Label31.Caption = "C"
        Else
            Label31.Caption = "i"
        End If
    End If
    'pregunta 4
    If (Text1(6) = "O" And Text1(10) = "S") And Text1(8) = "O" Then
        Label33.Caption = "C"
    Else
        Label33.Caption = "i"
    End If
    'pregunta 5
    If (Text1(28) = "M" And Text1(29) = "A") And (Text1(30) = "Ñ" And Text1(31) = "A") Then
        If (Text1(32) = "N" And Text1(33) = "A") And Text1(18) = "S" Then
            Label35.Caption = "C"
        Else
            Label35.Caption = "i"
        End If
    End If
    'pregunta 6
    If (Text1(13) = "F" And Text1(14) = "U") And (Text1(22) = "M" And Text1(23) = "O") Then
        Label37.Caption = "C"
    Else
        Label37.Caption = "i"
    End If
    'pregunta 7
    If (Text1(38) = "E" And Text1(40) = "D") And (Text1(41) = "E" And Text1(39) = "N") Then
        Label39.Caption = "C"
    Else
        Label39.Caption = "i"
    End If
    'pregunta 8
    If (Text1(45) = "M" And Text1(46) = "O") And (Text1(43) = "C" And Text1(44) = "O") Then
        Label41.Caption = "C"
    Else
        Label41.Caption = "i"
    End If
    'pregunta 9
    If (Text1(26) = "R" And Text1(35) = "A") And Text1(36) = "P" Then
        Label43.Caption = "C"
    Else
        Label43.Caption = "i"
    End If
    'pregunta 10
    If (Text1(65) = "R" And Text1(66) = "A") And (Text1(63) = "P" And Text1(64) = "A") Then
        If Text1(61) = "R" Then
            Label45.Caption = "C"
        Else
            Label45.Caption = "i"
        End If
    End If
    'pregunta 11
    If (Text1(59) = "V" And Text1(60) = "I") And (Text1(57) = "V" And Text1(55) = "A") Then
        Label47.Caption = "C"
    Else
        Label47.Caption = "i"
    End If
    'pregunta 12
    If (Text1(68) = "S" And Text1(69) = "I") And (Text1(70) = "E" And Text1(71) = "M") Then
        If (Text1(72) = "P" And Text1(73) = "R") And Text1(74) = "E" Then
            Label49.Caption = "C"
        Else
            Label49.Caption = "i"
        End If
    End If
    'pregunta 13
    If (Text1(48) = "F" And Text1(49) = "U") And (Text1(50) = "E" And Text1(51) = "S") Then
        If Text1(52) = "E" And Text1(53) = "S" Then
            Label51.Caption = "C"
        Else
            Label51.Caption = "i"
        End If
    End If
    'pregunta 14
    If (Text1(37) = "R" And Text1(1) = "I") And Text1(3) = "E" Then
        Label53.Caption = "C"
    Else
        Label53.Caption = "i"
    End If
    
End Sub
