VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   Caption         =   "CGPA Calculator "
   ClientHeight    =   10110
   ClientLeft      =   75
   ClientTop       =   855
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10110
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.TextBox Lcgp 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   0
      Left            =   2880
      TabIndex        =   79
      ToolTipText     =   "Enter Here Marks"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Lcgp 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   7
      Left            =   12120
      TabIndex        =   78
      ToolTipText     =   "Enter Here Marks"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Lcgp 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   6
      Left            =   10800
      TabIndex        =   77
      ToolTipText     =   "Enter Here Marks"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Lcgp 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   5
      Left            =   9480
      TabIndex        =   76
      ToolTipText     =   "Enter Here Marks"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Lcgp 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   4
      Left            =   8160
      TabIndex        =   75
      ToolTipText     =   "Enter Here Marks"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Lcgp 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   3
      Left            =   6840
      TabIndex        =   74
      ToolTipText     =   "Enter Here Marks"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Lcgp 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   2
      Left            =   5520
      TabIndex        =   73
      ToolTipText     =   "Enter Here Marks"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Lcgp 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   1
      Left            =   4200
      TabIndex        =   72
      ToolTipText     =   "Enter Here Marks"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox sgp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   7
      Left            =   12240
      TabIndex        =   71
      ToolTipText     =   "Total GPs Of All Cources"
      Top             =   8280
      Width           =   855
   End
   Begin VB.TextBox sgp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   3000
      TabIndex        =   70
      ToolTipText     =   "Total GPs Of All Cources"
      Top             =   8280
      Width           =   855
   End
   Begin VB.TextBox sch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   3000
      TabIndex        =   69
      ToolTipText     =   "Total Credit Hores Of  All Semisters"
      Top             =   8760
      Width           =   855
   End
   Begin VB.TextBox sgp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   4320
      TabIndex        =   68
      ToolTipText     =   "Total GPs Of All Cources"
      Top             =   8280
      Width           =   855
   End
   Begin VB.TextBox sgp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   5640
      TabIndex        =   67
      ToolTipText     =   "Total GPs Of All Cources"
      Top             =   8280
      Width           =   855
   End
   Begin VB.TextBox sgp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   6960
      TabIndex        =   66
      ToolTipText     =   "Total GPs Of All Cources"
      Top             =   8280
      Width           =   855
   End
   Begin VB.TextBox sgp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   8280
      TabIndex        =   65
      ToolTipText     =   "Total GPs Of All Cources"
      Top             =   8280
      Width           =   855
   End
   Begin VB.TextBox sgp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   9600
      TabIndex        =   64
      ToolTipText     =   "Total GPs Of All Cources"
      Top             =   8280
      Width           =   855
   End
   Begin VB.TextBox sgp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   6
      Left            =   10920
      TabIndex        =   63
      ToolTipText     =   "Total GPs Of All Cources"
      Top             =   8280
      Width           =   855
   End
   Begin VB.TextBox sch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   4320
      TabIndex        =   62
      ToolTipText     =   "Total Credit Hores Of  All Semisters"
      Top             =   8760
      Width           =   855
   End
   Begin VB.TextBox sch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   5640
      TabIndex        =   61
      ToolTipText     =   "Total Credit Hores Of  All Semisters"
      Top             =   8760
      Width           =   855
   End
   Begin VB.TextBox sch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   6960
      TabIndex        =   60
      ToolTipText     =   "Total Credit Hores Of  All Semisters"
      Top             =   8760
      Width           =   855
   End
   Begin VB.TextBox sch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   8280
      TabIndex        =   59
      ToolTipText     =   "Total Credit Hores Of  All Semisters"
      Top             =   8760
      Width           =   855
   End
   Begin VB.TextBox sch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   9600
      TabIndex        =   58
      ToolTipText     =   "Total Credit Hores Of  All Semisters"
      Top             =   8760
      Width           =   855
   End
   Begin VB.TextBox sch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   6
      Left            =   10920
      TabIndex        =   57
      ToolTipText     =   "Total Credit Hores Of  All Semisters"
      Top             =   8760
      Width           =   855
   End
   Begin VB.TextBox sch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   7
      Left            =   12240
      TabIndex        =   56
      ToolTipText     =   "Total Credit Hores Of  All Semisters"
      Top             =   8760
      Width           =   855
   End
   Begin VB.CommandButton sCmd 
      BackColor       =   &H8000000D&
      Caption         =   "Semister 6"
      Height          =   495
      Index           =   5
      Left            =   9480
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Go To Semister 6"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton sCmd 
      BackColor       =   &H8000000D&
      Caption         =   "Semister 1"
      Height          =   495
      Index           =   0
      Left            =   2880
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Go To Semister 1 "
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton sCmd 
      BackColor       =   &H8000000D&
      Caption         =   "Semister 2"
      Height          =   495
      Index           =   1
      Left            =   4200
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Go To Semister 2"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton sCmd 
      BackColor       =   &H8000000D&
      Caption         =   "Semister 3"
      Height          =   495
      Index           =   2
      Left            =   5520
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Go To Semister 3"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton sCmd 
      BackColor       =   &H8000000D&
      Caption         =   "Semister 4"
      Height          =   495
      Index           =   3
      Left            =   6840
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Go To Semister 4"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton sCmd 
      BackColor       =   &H8000000D&
      Caption         =   "Semister 5"
      Height          =   495
      Index           =   4
      Left            =   8160
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Go To Semister 5"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton sCmd 
      BackColor       =   &H8000000D&
      Caption         =   "Semister 7"
      Height          =   495
      Index           =   6
      Left            =   10800
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Go To Semister 7"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton sCmd 
      BackColor       =   &H8000000D&
      Caption         =   "Semister 8"
      Height          =   495
      Index           =   7
      Left            =   12120
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Go To Semister 8 "
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "C GPA :"
      Height          =   2295
      Left            =   10920
      TabIndex        =   42
      ToolTipText     =   "CGPA Of Calculated E(GP x Credit Hores)/Total Credit Hores Of All Cources"
      Top             =   960
      Width           =   3255
      Begin VB.TextBox cgpa 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   600
         TabIndex        =   43
         ToolTipText     =   "CGPA Of Calculated E(GP x Credit Hores)/Total Credit Hores Of All Cources"
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000000&
      Caption         =   "Next Semister"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Go To Next Semister"
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000000&
      Caption         =   "Back Semister"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Go To Back Semister"
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Semister 1"
      Height          =   5895
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   9615
      Begin VB.TextBox sbj 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   2
         Left            =   2160
         TabIndex        =   17
         ToolTipText     =   "Enter Here Marks"
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox sbj 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   3
         Left            =   2160
         TabIndex        =   16
         ToolTipText     =   "Enter Here Marks"
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox sbj 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   4
         Left            =   2160
         TabIndex        =   15
         ToolTipText     =   "Enter Here Marks"
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox sbj 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   5
         Left            =   2160
         TabIndex        =   14
         ToolTipText     =   "Enter Here Marks"
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox ch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   5
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   13
         ToolTipText     =   "Write Credit Hores"
         Top             =   5160
         Width           =   855
      End
      Begin VB.TextBox ch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   0
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   12
         ToolTipText     =   "Write Credit Hores"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox ch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   1
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   11
         ToolTipText     =   "Write Credit Hores"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox ch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   2
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   10
         ToolTipText     =   "Write Credit Hores"
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox ch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   3
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   9
         ToolTipText     =   "Write Credit Hores"
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox ch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   4
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   8
         ToolTipText     =   "Write Credit Hores"
         Top             =   4320
         Width           =   855
      End
      Begin VB.TextBox sbj 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   1
         Left            =   2160
         TabIndex        =   7
         ToolTipText     =   "Enter Here Marks"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox sbj 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   0
         Left            =   2160
         TabIndex        =   6
         ToolTipText     =   "Enter Here Marks"
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "C.H:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6600
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Marks:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   39
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label sb 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject 1:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label sb 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject 2:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   37
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label sb 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject 3:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   36
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label sb 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject 4:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   35
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label sb 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject 5:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   240
         TabIndex        =   34
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label sb 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject 6:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   240
         TabIndex        =   33
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "G.P:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   32
         Top             =   240
         Width           =   975
      End
      Begin VB.Label gp 
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   3360
         TabIndex        =   31
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label gp 
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   3360
         TabIndex        =   29
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label gp 
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   3360
         TabIndex        =   28
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label gp 
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   3360
         TabIndex        =   27
         Top             =   4200
         Width           =   2895
      End
      Begin VB.Label gp 
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   3360
         TabIndex        =   26
         Top             =   5040
         Width           =   2895
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "G.P x C.H"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   0
         Left            =   7920
         TabIndex        =   24
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   1
         Left            =   7920
         TabIndex        =   23
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   2
         Left            =   7920
         TabIndex        =   22
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   3
         Left            =   7920
         TabIndex        =   21
         Top             =   3480
         Width           =   1200
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   4
         Left            =   7920
         TabIndex        =   20
         Top             =   4320
         Width           =   1200
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   5
         Left            =   7920
         TabIndex        =   19
         Top             =   5160
         Width           =   1200
      End
      Begin VB.Label gp 
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   3360
         TabIndex        =   18
         Top             =   3360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Total  GPA "
      Height          =   2535
      Left            =   11280
      TabIndex        =   0
      Top             =   3720
      Width           =   3135
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   600
         MaxLength       =   4
         TabIndex        =   1
         ToolTipText     =   "GPA Of This Semister"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label glbl 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   4
         ToolTipText     =   "Total Grade Points & Credit Hores"
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label clbl 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         TabIndex        =   3
         ToolTipText     =   "Total Grade Points & Credit Hores"
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   " T G.P     T C.H"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   2
         ToolTipText     =   "Total Grade Points & Credit Hores"
         Top             =   1320
         Width           =   2175
      End
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "T G.P"
      Height          =   255
      Left            =   2040
      TabIndex        =   55
      Top             =   8400
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "T C.H"
      Height          =   255
      Left            =   2040
      TabIndex        =   54
      Top             =   8880
      Width           =   615
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Go To"
      Height          =   375
      Left            =   2040
      TabIndex        =   52
      Top             =   6960
      Width           =   735
   End
   Begin VB.Menu file 
      Caption         =   "File"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu s1 
         Caption         =   "Semister 1"
         Index           =   1
      End
      Begin VB.Menu s1 
         Caption         =   "Semister 2"
         Index           =   2
      End
      Begin VB.Menu s1 
         Caption         =   "Semister 3"
         Index           =   3
      End
      Begin VB.Menu s1 
         Caption         =   "Semister 4"
         Index           =   4
      End
      Begin VB.Menu s1 
         Caption         =   "Semister 5"
         Index           =   5
      End
      Begin VB.Menu s1 
         Caption         =   "Semister 6"
         Index           =   6
      End
      Begin VB.Menu s1 
         Caption         =   "Semister 7"
         Index           =   7
      End
      Begin VB.Menu s1 
         Caption         =   "Semister 8"
         Index           =   8
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
   End
   Begin VB.Menu aboutme 
      Caption         =   "About Me"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public d, i, j, k, l, Semister, p, th, q, z, ii, ij, ik, t As Integer
Public r
Dim b, c, M(5), n(5), o(5), Edis(8, 8), Edic(8, 8)

Function tgp(k) As Integer
On Error GoTo error
Select Case ch(k).Text
Case Is <> ""
If sbj(k) >= 50 Then
M(k) = ch(k).Text
Label6(k).Caption = M(k) * n(k)

Else
Label6(k).Caption = ""
End If
Case Else
Label6(k).Caption = ""
End Select
error:
Err.Clear
End Function

Function glb(r) As Integer

glbl.Caption = ""
For i = 0 To 5
p = Label6(i).Caption

Select Case p
Case Is > 0
glbl.Caption = Val(glbl.Caption) + Val(Label6(i).Caption)

Case Is > 100
glbl.Caption = " >100"

End Select
Next
End Function


Private Sub aboutme_Click()
Form9.Show
End Sub

Private Sub clbl_Click()
'MsgBox (clbl.Caption & "     " & glbl.Caption)
End Sub

Private Sub Form_Activate()
sbj(0).SetFocus
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
PopupMenu file
End If
End Sub

Private Sub Form_Load()
For k = 0 To 5
Label6(k).BackColor = vbGreen
Next
Frame1.BackColor = vbGreen
Frame2.BackColor = vbGreen
Frame4.BackColor = vbGreen
clbl.BackColor = vbGreen
glbl.BackColor = vbGreen
i = dc()
ii = 1
ik = 1
For j = 0 To 5
gp(j).BackColor = vbGreen
Next
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
PopupMenu file
End If
End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
PopupMenu file
End If
End Sub


Private Sub Frame4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
PopupMenu file
End If
End Sub

Private Sub gp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
PopupMenu file
End If
End Sub

Private Sub Label6_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
PopupMenu file
End If
End Sub



Private Sub sbj_Change(j As Integer)
Label3.BackColor = vbGreen
b = sbj(j).Text
Select Case b
Case Is < 0
sbj(j).Text = ""
Case Is > 100
sbj(j).Text = ""
End Select


k = sbj(j).Text
Select Case k
Case Is >= 101
Label3.BackColor = vbRed
gp(j).Caption = ""
n(j) = 0

 Case Is >= 85
gp(j).Caption = "4.00      A+"
i = dc()
n(j) = 4
gp(j).BackColor = &H8000000E
Case Is >= 80
i = dc()
gp(j).Caption = "3.70      A"
n(j) = 3.7
gp(j).BackColor = &H80000014
Case Is >= 75
gp(j).Caption = "3.30      B+"
i = dc()
n(j) = 3.3
gp(j).BackColor = &H80000010
Case Is >= 70
gp(j).Caption = "3.00      B"
i = dc()
n(j) = 3
gp(j).BackColor = &H80000016
Case Is >= 65
i = dc()
gp(j).Caption = "2.70      B-"
n(j) = 2.7
gp(j).BackColor = &H80000016
Case Is >= 61
i = dc()
gp(j).Caption = "2.30      C+"
n(j) = 2.3
gp(j).BackColor = &H80000000
Case Is >= 58
i = dc()
gp(j).Caption = "2.00      C"
n(j) = 2
gp(j).BackColor = &H80000000
Case Is >= 55
i = dc()
gp(j).Caption = "1.70       C-"
n(j) = 1.7
gp(j).BackColor = &H80000000
Case Is >= 53
i = dc()
gp(j).Caption = "1.30       D+"
n(j) = 1.3
gp(j).BackColor = &H80000010
Case Is >= 50
i = dc()
gp(j).Caption = "1.00       D"
n(j) = 1
 gp(j).BackColor = &H80000010
Case Is >= 11
gp(j).Caption = " Fail!      F"
gp(j).BackColor = vbRed
Label1.BackColor = vbRed
n(j) = 0

Case Is <= 10
i = dc()
gp(j).Caption = ""
Case Is = ""
gp(j).Caption = ""
End Select

For k = 0 To 5
z = 2
Select Case z
Case Is = Len(sbj(k).Text)
If ch(k).Text = "" Then
ch(k).SetFocus
End If
End Select

Next
End Sub


Private Sub ch_Change(r As Integer)
On Error GoTo error

If Len(ch(5).Text) = 1 Then
Command1.SetFocus
End If

For k = 0 To 4
Select Case 1
Case Is = Len(ch(k).Text)
If sbj(k + 1).Text = "" Then
sbj(k + 1).SetFocus

End If
End Select
Next


For k = 0 To 5
i = tgp(k)
i = glb(k)
Next
If Label6(r).Caption = "" Then
Else
clbl.Caption = Val(clbl.Caption) + Val(ch(r))
End If
Select Case clbl.Caption
Case Is > 0
If ch(r).Text = "" Then
Else
Text1.Text = glbl.Caption / clbl.Caption
End If
Case Is < 0
End Select


sgp(ii - 1) = glbl.Caption
sch(ii - 1) = clbl.Caption

b = Val(sgp(0).Text) + Val(sgp(1).Text) + Val(sgp(2).Text) + Val(sgp(3).Text) + Val(sgp(4).Text) + Val(sgp(5).Text) + Val(sgp(6).Text) + Val(sgp(7).Text)
c = Val(sch(0).Text) + Val(sch(1).Text) + Val(sch(2).Text) + Val(sch(3).Text) + Val(sch(4).Text) + Val(sch(5).Text) + Val(sch(6).Text) + Val(sch(7).Text)
cgpa.Text = ""
cgpa.Text = Round((b / c), 2)

Exit Sub
error:
'MsgBox (b & "        " & c)
End Sub


Function dc() As Integer
Form1.BackColor = vbYellow
Label3.BackColor = vbRed
For i = 0 To 5

sbj(i).BackColor = vbWhite

Next
Label1.BackColor = vbRed
End Function

Public Function Sf(i)
If i = 0 Then
ik = ii
ii = 1
ElseIf i = 1 Then
ik = ii
ii = 2
ElseIf i = 2 Then
ik = ii
ii = 3
ElseIf i = 3 Then
ik = ii
ii = 4
ElseIf i = 4 Then
ik = ii
ii = 5
ElseIf i = 5 Then
ik = ii
ii = 6
ElseIf i = 6 Then
ik = ii
ii = 7
ElseIf i = 7 Then
ik = ii
ii = 8
End If
End Function

Public Function Edia(ii, ik)
' The REAL CODE
clbl.Caption = ""

For ij = 0 To 5
Edis(ik, ij) = sbj(ij)
sbj(ij) = ""
sbj(ij) = Edis(ii, ij)

Edic(ik, ij) = ch(ij)
ch(ij) = ""
ch(ij) = Edic(ii, ij)
Next
Frame1.Caption = "Semister " & ii
For j = 0 To 5
gp(j).BackColor = vbGreen
Next
End Function

Private Sub s1_Click(Index As Integer)
j = Sf(Index)

If ik < 0 & ik < 9 Then
sgp(ik - 1) = glbl.Caption
sch(ik - 1) = clbl.Caption
End If

j = Edia(ii, ik)

If Len(sbj(0).Text) = 0 Then
sbj(0).SetFocus
End If

End Sub

Private Sub sch_Change(Index As Integer)
Lcgp(Index).Text = Text1.Text
End Sub

Private Sub sCmd_Click(i As Integer)
'On Error GoTo error
j = Sf(i)

If ik < 0 & ik < 9 Then
sgp(ik - 1) = glbl.Caption
sch(ik - 1) = clbl.Caption
End If

j = Edia(ii, ik)

If Len(sbj(0).Text) = 0 Then
sbj(0).SetFocus
End If
error:
End Sub

Private Sub Command1_Click()
'On Error GoTo error:
j = Sf(i)

If ik > 0 & ik < 9 Then
sgp(ik - 1) = glbl.Caption
sch(ik - 1) = clbl.Caption
End If

If ik < 8 Then
ii = ik + 1
Text1.Text = ""
Else
ii = 1
End If
j = Edia(ii, ik)

If Len(sbj(0).Text) = 0 Then
sbj(0).SetFocus
End If

error:

End Sub

Private Sub Command2_Click()
j = Sf(i)

If ik < 0 & ik < 9 Then
sgp(ik - 1) = glbl.Caption
sch(ik - 1) = clbl.Caption
End If

If ik > 1 Then
ii = ik - 1
Else: ii = 8
End If

j = Edia(ii, ik)
If Len(sbj(0).Text) = 0 Then
sbj(0).SetFocus
End If

End Sub


