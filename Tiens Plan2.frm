VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tiens Plan"
   ClientHeight    =   10845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15120
   FillColor       =   &H00C00000&
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   Icon            =   "Tiens Plan2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10845
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      Picture         =   "Tiens Plan2.frx":324A
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   124
      Top             =   0
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3960
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10335
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   14535
      Begin VB.CommandButton Ranks 
         Caption         =   "1*"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2760
         TabIndex        =   120
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox NM 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   5280
         TabIndex        =   119
         Text            =   "1"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox MP 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   3960
         TabIndex        =   118
         Text            =   "2"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox NP 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   6720
         TabIndex        =   117
         Text            =   "100"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TP 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   8160
         TabIndex        =   116
         Text            =   "200"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox DB 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   9600
         TabIndex        =   115
         Text            =   "500"
         Top             =   960
         Width           =   975
      End
      Begin VB.Frame Form1 
         Caption         =   "New PV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8895
         Index           =   4
         Left            =   6600
         TabIndex        =   94
         Top             =   600
         Width           =   1215
         Begin VB.TextBox NP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   105
            Text            =   "200"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox NP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   104
            Text            =   "400"
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox NP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   103
            Text            =   "800"
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox NP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   102
            Text            =   "1600"
            Top             =   3240
            Width           =   975
         End
         Begin VB.TextBox NP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   101
            Text            =   "3200"
            Top             =   3960
            Width           =   975
         End
         Begin VB.TextBox NP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   100
            Text            =   "6400"
            Top             =   4680
            Width           =   975
         End
         Begin VB.TextBox NP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   99
            Text            =   "12800"
            Top             =   5400
            Width           =   975
         End
         Begin VB.TextBox NP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   98
            Text            =   "25600"
            Top             =   6120
            Width           =   975
         End
         Begin VB.TextBox NP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   97
            Text            =   "51200"
            Top             =   6840
            Width           =   975
         End
         Begin VB.TextBox NP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   96
            Text            =   "102400"
            Top             =   7560
            Width           =   975
         End
         Begin VB.TextBox NP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   95
            Text            =   "204800"
            Top             =   8280
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "    Months"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8895
         Index           =   0
         Left            =   1080
         TabIndex        =   90
         Top             =   600
         Width           =   1215
         Begin VB.CommandButton Months 
            Caption         =   "1st"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   114
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton Months 
            Caption         =   "4th"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   240
            TabIndex        =   113
            Top             =   2640
            Width           =   735
         End
         Begin VB.CommandButton Months 
            Caption         =   "5th"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   240
            TabIndex        =   112
            Top             =   3360
            Width           =   735
         End
         Begin VB.CommandButton Months 
            Caption         =   "6th"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   240
            TabIndex        =   111
            Top             =   4080
            Width           =   735
         End
         Begin VB.CommandButton Months 
            Caption         =   "7th"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   240
            TabIndex        =   110
            Top             =   4800
            Width           =   735
         End
         Begin VB.CommandButton Months 
            Caption         =   "8th"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   240
            TabIndex        =   109
            Top             =   5520
            Width           =   735
         End
         Begin VB.CommandButton Months 
            Caption         =   "9th"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   240
            TabIndex        =   108
            Top             =   6240
            Width           =   735
         End
         Begin VB.CommandButton Months 
            Caption         =   "10th"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   9
            Left            =   240
            TabIndex        =   107
            Top             =   6960
            Width           =   735
         End
         Begin VB.CommandButton Months 
            Caption         =   "11th"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   10
            Left            =   240
            TabIndex        =   106
            Top             =   7680
            Width           =   735
         End
         Begin VB.CommandButton Months 
            Caption         =   "2nd"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   240
            TabIndex        =   93
            Top             =   1080
            Width           =   735
         End
         Begin VB.CommandButton Months 
            Caption         =   "3rd"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   240
            TabIndex        =   92
            Top             =   1800
            Width           =   735
         End
         Begin VB.CommandButton Months 
            Caption         =   "12th"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   11
            Left            =   240
            TabIndex        =   91
            Top             =   8280
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Men Power"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8895
         Index           =   3
         Left            =   3840
         TabIndex        =   78
         Top             =   600
         Width           =   1215
         Begin VB.TextBox MP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   89
            Text            =   "4"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox MP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   88
            Text            =   "8"
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox MP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   87
            Text            =   "16"
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox MP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   86
            Text            =   "32"
            Top             =   3240
            Width           =   975
         End
         Begin VB.TextBox MP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   85
            Text            =   "64"
            Top             =   3960
            Width           =   975
         End
         Begin VB.TextBox MP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   84
            Text            =   "128"
            Top             =   4680
            Width           =   975
         End
         Begin VB.TextBox MP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   83
            Text            =   "256"
            Top             =   5400
            Width           =   975
         End
         Begin VB.TextBox MP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   82
            Text            =   "512"
            Top             =   6120
            Width           =   975
         End
         Begin VB.TextBox MP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   81
            Text            =   "1024"
            Top             =   6840
            Width           =   975
         End
         Begin VB.TextBox MP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   80
            Text            =   "2048"
            Top             =   7560
            Width           =   975
         End
         Begin VB.TextBox MP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   79
            Text            =   "4096"
            Top             =   8280
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "  Ranks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8895
         Index           =   1
         Left            =   2520
         TabIndex        =   65
         Top             =   600
         Width           =   1215
         Begin VB.CommandButton Ranks 
            Caption         =   "2*"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   240
            TabIndex        =   77
            Top             =   1080
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "3*"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   240
            TabIndex        =   76
            Top             =   1800
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "4*"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   240
            TabIndex        =   75
            Top             =   2520
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "4*"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   240
            TabIndex        =   74
            Top             =   3240
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "5*"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   240
            TabIndex        =   73
            Top             =   3960
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "5*"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   240
            TabIndex        =   72
            Top             =   4680
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "6*"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   240
            TabIndex        =   71
            Top             =   5400
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "6*"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   240
            TabIndex        =   70
            Top             =   6120
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "7*"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   9
            Left            =   240
            TabIndex        =   69
            Top             =   6840
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "7*"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   10
            Left            =   240
            TabIndex        =   68
            Top             =   7560
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "8*"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   11
            Left            =   240
            TabIndex        =   67
            Top             =   8280
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "5*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   12
            Left            =   240
            TabIndex        =   66
            Top             =   9000
            Width           =   735
         End
         Begin VB.Shape Shape1 
            Height          =   375
            Left            =   480
            Top             =   1920
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "New Men"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8895
         Index           =   2
         Left            =   5160
         TabIndex        =   52
         Top             =   600
         Width           =   1215
         Begin VB.TextBox NM 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   64
            Text            =   "2"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox NM 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   63
            Text            =   "4"
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox NM 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   62
            Text            =   "8"
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox NM 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   61
            Text            =   "16"
            Top             =   3240
            Width           =   975
         End
         Begin VB.TextBox NM 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   60
            Text            =   "32"
            Top             =   3960
            Width           =   975
         End
         Begin VB.TextBox NM 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   59
            Text            =   "64"
            Top             =   4680
            Width           =   975
         End
         Begin VB.TextBox NM 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   58
            Text            =   "128"
            Top             =   5400
            Width           =   975
         End
         Begin VB.TextBox NM 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   57
            Text            =   "256"
            Top             =   6120
            Width           =   975
         End
         Begin VB.TextBox NM 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   56
            Text            =   "512"
            Top             =   6840
            Width           =   975
         End
         Begin VB.TextBox NM 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   55
            Text            =   "1024"
            Top             =   7560
            Width           =   975
         End
         Begin VB.TextBox NM 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   54
            Text            =   "2048"
            Top             =   8280
            Width           =   975
         End
         Begin VB.TextBox NM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   12
            Left            =   120
            TabIndex        =   53
            Text            =   "0"
            Top             =   9000
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Total  PV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8895
         Index           =   5
         Left            =   8040
         TabIndex        =   40
         Top             =   600
         Width           =   1215
         Begin VB.TextBox TP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Text            =   "400"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox TP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   50
            Text            =   "800"
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox TP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   49
            Text            =   "1600"
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox TP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   48
            Text            =   "3200"
            Top             =   3240
            Width           =   975
         End
         Begin VB.TextBox TP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   47
            Text            =   "6400"
            Top             =   3960
            Width           =   975
         End
         Begin VB.TextBox TP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   46
            Text            =   "12800"
            Top             =   4680
            Width           =   975
         End
         Begin VB.TextBox TP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   45
            Text            =   "25600"
            Top             =   5400
            Width           =   975
         End
         Begin VB.TextBox TP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   44
            Text            =   "51200"
            Top             =   6120
            Width           =   975
         End
         Begin VB.TextBox TP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   43
            Text            =   "102400"
            Top             =   6840
            Width           =   975
         End
         Begin VB.TextBox TP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   42
            Text            =   "204800"
            Top             =   7560
            Width           =   975
         End
         Begin VB.TextBox TP 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   41
            Text            =   "409600"
            Top             =   8280
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Direct Bonus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8895
         Index           =   6
         Left            =   9480
         TabIndex        =   27
         Top             =   600
         Width           =   1215
         Begin VB.TextBox DB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Text            =   "500"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox DB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Text            =   "2000"
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox DB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   37
            Text            =   "2400"
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox DB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   36
            Text            =   "2400"
            Top             =   3240
            Width           =   975
         End
         Begin VB.TextBox DB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   35
            Text            =   "2800"
            Top             =   3960
            Width           =   975
         End
         Begin VB.TextBox DB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   34
            Text            =   "2800"
            Top             =   4680
            Width           =   975
         End
         Begin VB.TextBox DB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   33
            Text            =   "3200"
            Top             =   5400
            Width           =   975
         End
         Begin VB.TextBox DB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   32
            Text            =   "3200"
            Top             =   6120
            Width           =   975
         End
         Begin VB.TextBox DB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   31
            Text            =   "3600"
            Top             =   6840
            Width           =   975
         End
         Begin VB.TextBox DB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   30
            Text            =   "3600"
            Top             =   7560
            Width           =   975
         End
         Begin VB.TextBox DB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   29
            Text            =   "4000"
            Top             =   8280
            Width           =   975
         End
         Begin VB.TextBox DB 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   12
            Left            =   120
            TabIndex        =   28
            Top             =   9000
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Indirect Bonus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8895
         Index           =   7
         Left            =   10920
         TabIndex        =   14
         Top             =   600
         Width           =   1215
         Begin VB.TextBox IB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   122
            Text            =   "0"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox TB 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   14
            Left            =   2400
            TabIndex        =   121
            Text            =   "0"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox IB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Text            =   "0"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox IB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Text            =   "0"
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox IB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   24
            Text            =   "0"
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox IB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   23
            Text            =   "0"
            Top             =   3240
            Width           =   975
         End
         Begin VB.TextBox IB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   22
            Text            =   "0"
            Top             =   3960
            Width           =   975
         End
         Begin VB.TextBox IB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   21
            Text            =   "0"
            Top             =   4680
            Width           =   975
         End
         Begin VB.TextBox IB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   20
            Text            =   "0"
            Top             =   5400
            Width           =   975
         End
         Begin VB.TextBox IB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   19
            Text            =   "0"
            Top             =   6120
            Width           =   975
         End
         Begin VB.TextBox IB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   18
            Text            =   "0"
            Top             =   6840
            Width           =   975
         End
         Begin VB.TextBox IB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   17
            Text            =   "0"
            Top             =   7560
            Width           =   975
         End
         Begin VB.TextBox IB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   16
            Text            =   "0"
            Top             =   8280
            Width           =   975
         End
         Begin VB.TextBox IB 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   12
            Left            =   120
            TabIndex        =   15
            Text            =   "0"
            Top             =   9000
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Total Bonus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8895
         Index           =   8
         Left            =   12360
         TabIndex        =   1
         Top             =   600
         Width           =   1455
         Begin VB.TextBox TB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   123
            Text            =   "0"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox TB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Text            =   "0"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox TB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Text            =   "0"
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox TB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   11
            Text            =   "0"
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox TB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   10
            Text            =   "0"
            Top             =   3240
            Width           =   1215
         End
         Begin VB.TextBox TB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   9
            Text            =   "0"
            Top             =   3960
            Width           =   1215
         End
         Begin VB.TextBox TB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   8
            Text            =   "0"
            Top             =   4680
            Width           =   1215
         End
         Begin VB.TextBox TB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   7
            Text            =   "0"
            Top             =   5400
            Width           =   1215
         End
         Begin VB.TextBox TB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   6
            Text            =   "0"
            Top             =   6120
            Width           =   1215
         End
         Begin VB.TextBox TB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   5
            Text            =   "0"
            Top             =   6840
            Width           =   1215
         End
         Begin VB.TextBox TB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   4
            Text            =   "0"
            Top             =   7560
            Width           =   1215
         End
         Begin VB.TextBox TB 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   3
            Text            =   "0"
            Top             =   8280
            Width           =   1215
         End
         Begin VB.TextBox TB 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   12
            Left            =   120
            TabIndex        =   2
            Text            =   "0"
            Top             =   9000
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, i, d As Double

Private Sub Timer1_Timer()
On Error GoTo error

Frame1.Caption = "       Tiens            " & Time & "             " & Date
For i = 0 To 11

a = Val(NP(i).Text) - 100
b = a * 100

c = Val((4 * b) / 100)
IB(i).Text = c
TB(i).Text = Val(IB(i).Text) + Val(DB(i).Text)


IB(0).Text = "1500"

Next
error:
Err.Clear

End Sub
