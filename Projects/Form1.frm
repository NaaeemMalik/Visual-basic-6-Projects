VERSION 5.00
Begin VB.Form Form1 
   Caption         =   $"Form1.frx":0000
   ClientHeight    =   10110
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":008A
   LinkTopic       =   "Form1"
   ScaleHeight     =   10110
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000002&
      Caption         =   "Calculate CGPA"
      Height          =   1095
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "To Find CGPA Of Calculated E(GP x Credit Hores)/Total Credit Hores Of All Cources Click Here"
      Top             =   240
      Width           =   2895
   End
   Begin VB.Frame Frame4 
      Caption         =   "C GPA :"
      Height          =   2295
      Left            =   11280
      TabIndex        =   70
      ToolTipText     =   "CGPA Of Calculated E(GP x Credit Hores)/Total Credit Hores Of All Cources"
      Top             =   1320
      Width           =   3255
      Begin VB.TextBox cgpa 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   600
         TabIndex        =   71
         ToolTipText     =   "CGPA Of Calculated E(GP x Credit Hores)/Total Credit Hores Of All Cources"
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000000&
      Caption         =   "Next Semister"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "Go To Next Semister"
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Accesing  Others Semisters "
      Height          =   2895
      Left            =   2160
      TabIndex        =   41
      Top             =   6360
      Visible         =   0   'False
      Width           =   11895
      Begin VB.CommandButton sCmd 
         BackColor       =   &H8000000D&
         Caption         =   "Semister 1"
         Height          =   495
         Index           =   0
         Left            =   1200
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Go To Semister 1 "
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox sgp 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   0
         Left            =   1200
         TabIndex        =   64
         ToolTipText     =   "Total GPs Of All Cources"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox sch 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   0
         Left            =   1200
         TabIndex        =   63
         ToolTipText     =   "Total Credit Hores Of  All Semisters"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox sgp 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   1
         Left            =   2520
         TabIndex        =   62
         ToolTipText     =   "Total GPs Of All Cources"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox sgp 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   2
         Left            =   3840
         TabIndex        =   61
         ToolTipText     =   "Total GPs Of All Cources"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox sgp 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   3
         Left            =   5160
         TabIndex        =   60
         ToolTipText     =   "Total GPs Of All Cources"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox sgp 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   4
         Left            =   6480
         TabIndex        =   59
         ToolTipText     =   "Total GPs Of All Cources"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox sgp 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   5
         Left            =   7800
         TabIndex        =   58
         ToolTipText     =   "Total GPs Of All Cources"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox sgp 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   6
         Left            =   9120
         TabIndex        =   57
         ToolTipText     =   "Total GPs Of All Cources"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox sgp 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   7
         Left            =   10440
         TabIndex        =   56
         ToolTipText     =   "Total GPs Of All Cources"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox sch 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   1
         Left            =   2520
         TabIndex        =   55
         ToolTipText     =   "Total Credit Hores Of  All Semisters"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox sch 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   2
         Left            =   3840
         TabIndex        =   54
         ToolTipText     =   "Total Credit Hores Of  All Semisters"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox sch 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   3
         Left            =   5160
         TabIndex        =   53
         ToolTipText     =   "Total Credit Hores Of  All Semisters"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox sch 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   4
         Left            =   6480
         TabIndex        =   52
         ToolTipText     =   "Total Credit Hores Of  All Semisters"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox sch 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   5
         Left            =   7800
         TabIndex        =   51
         ToolTipText     =   "Total Credit Hores Of  All Semisters"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox sch 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   6
         Left            =   9120
         TabIndex        =   50
         ToolTipText     =   "Total Credit Hores Of  All Semisters"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox sch 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   7
         Left            =   10440
         TabIndex        =   49
         ToolTipText     =   "Total Credit Hores Of  All Semisters"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton sCmd 
         BackColor       =   &H8000000D&
         Caption         =   "Semister 2"
         Height          =   495
         Index           =   1
         Left            =   2520
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Go To Semister 2"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton sCmd 
         BackColor       =   &H8000000D&
         Caption         =   "Semister 3"
         Height          =   495
         Index           =   2
         Left            =   3840
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Go To Semister 3"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton sCmd 
         BackColor       =   &H8000000D&
         Caption         =   "Semister 4"
         Height          =   495
         Index           =   3
         Left            =   5160
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Go To Semister 4"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton sCmd 
         BackColor       =   &H8000000D&
         Caption         =   "Semister 5"
         Height          =   495
         Index           =   4
         Left            =   6480
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Go To Semister 5"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton sCmd 
         BackColor       =   &H8000000D&
         Caption         =   "Semister 6"
         Height          =   495
         Index           =   5
         Left            =   7800
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Go To Semister 6"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton sCmd 
         BackColor       =   &H8000000D&
         Caption         =   "Semister 7"
         Height          =   495
         Index           =   6
         Left            =   9120
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Go To Semister 7"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton sCmd 
         BackColor       =   &H8000000D&
         Caption         =   "Semister 8"
         Height          =   495
         Index           =   7
         Left            =   10440
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Go To Semister 8 "
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "T G.P"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   68
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "T C.H"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Go To"
         Height          =   375
         Left            =   240
         TabIndex        =   66
         Top             =   480
         Width           =   735
      End
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
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
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
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   2760
         Top             =   0
      End
      Begin VB.TextBox ch 
         BeginProperty Font 
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
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
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label sb 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject 2:"
         BeginProperty Font 
            Name            =   "Book Antiqua"
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
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label sb 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject 3:"
         BeginProperty Font 
            Name            =   "Book Antiqua"
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
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label sb 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject 4:"
         BeginProperty Font 
            Name            =   "Book Antiqua"
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
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label sb 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject 5:"
         BeginProperty Font 
            Name            =   "Book Antiqua"
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
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label sb 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject 6:"
         BeginProperty Font 
            Name            =   "Book Antiqua"
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
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "G.P:"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3720
         TabIndex        =   32
         Top             =   240
         Width           =   975
      End
      Begin VB.Label gp 
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   3600
         TabIndex        =   31
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade:"
         BeginProperty Font 
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   3600
         TabIndex        =   29
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label gp 
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   3600
         TabIndex        =   28
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label gp 
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   3600
         TabIndex        =   27
         Top             =   4200
         Width           =   2655
      End
      Begin VB.Label gp 
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   3600
         TabIndex        =   26
         Top             =   5040
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "G.P x C.H"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   25
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   0
         Left            =   7800
         TabIndex        =   24
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   1
         Left            =   7800
         TabIndex        =   23
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   2
         Left            =   7800
         TabIndex        =   22
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   3
         Left            =   7800
         TabIndex        =   21
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   4
         Left            =   7800
         TabIndex        =   20
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   5
         Left            =   7800
         TabIndex        =   19
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label gp 
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   3600
         TabIndex        =   18
         Top             =   3360
         Width           =   2655
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
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
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
            Name            =   "Book Antiqua"
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
   Begin VB.Menu file 
      Caption         =   "File"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu goto 
         Caption         =   "Go To"
      End
      Begin VB.Menu s1 
         Caption         =   "Semister 1"
         Shortcut        =   ^A
      End
      Begin VB.Menu s2 
         Caption         =   "Semister 2"
         Shortcut        =   ^B
      End
      Begin VB.Menu s3 
         Caption         =   "Semister 3"
         Shortcut        =   ^C
      End
      Begin VB.Menu s4 
         Caption         =   "Semister 4"
         Shortcut        =   ^D
      End
      Begin VB.Menu s5 
         Caption         =   "Semister 5"
         Shortcut        =   ^E
      End
      Begin VB.Menu s6 
         Caption         =   "Semister 6"
         Shortcut        =   ^F
      End
      Begin VB.Menu s7 
         Caption         =   "Semister 7"
         Shortcut        =   ^G
      End
      Begin VB.Menu s8 
         Caption         =   "Semister 8"
         Shortcut        =   ^H
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

Dim b, c, d, i, j, k, l, M(5), n(5), o(5), p, th, r, q, s, z, t As Integer

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
Form1.Hide
Form9.Show
End Sub

Private Sub cgpa_Click()
On Error GoTo error
b = Val(sgp(0).Text) + Val(sgp(1).Text) + Val(sgp(2).Text) + Val(sgp(3).Text) + Val(sgp(4).Text) + Val(sgp(5).Text) + Val(sgp(6).Text) + Val(sgp(7).Text)
c = Val(sch(0).Text) + Val(sch(1).Text) + Val(sch(2).Text) + Val(sch(3).Text) + Val(sch(4).Text) + Val(sch(5).Text) + Val(sch(6).Text) + Val(sch(7).Text)
cgpa.Text = ""
cgpa.Text = (b / c)
error:
Err.Clear
End Sub

Private Sub Command1_Click()
Form1.Hide
Form2.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = M1.sSgp(0)
End Sub

Private Sub Command2_Click()
On Error GoTo error
b = Val(sgp(0).Text) + Val(sgp(1).Text) + Val(sgp(2).Text) + Val(sgp(3).Text) + Val(sgp(4).Text) + Val(sgp(5).Text) + Val(sgp(6).Text) + Val(sgp(7).Text)
c = Val(sch(0).Text) + Val(sch(1).Text) + Val(sch(2).Text) + Val(sch(3).Text) + Val(sch(4).Text) + Val(sch(5).Text) + Val(sch(6).Text) + Val(sch(7).Text)
cgpa.Text = ""
cgpa.Text = (b / c)
error:
Err.Clear
End Sub

Private Sub Command3_Click()
For i = 0 To 5
k = Len(sbj(i).Text)
If k = 2 Then
Print k

End If
Next
End Sub

Private Sub Form_Activate()
sbj(0).SetFocus
For k = 0 To 5
Label6(k).BackColor = vbGreen
Next

End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub opt_Click(i As Integer)
For k = 0 To 5

ch(k).Text = i
Next
End Sub


Private Sub Form_Load()
Frame1.BackColor = vbGreen
Frame2.BackColor = vbGreen
Frame3.BackColor = vbGreen
Frame4.BackColor = vbGreen
clbl.BackColor = vbGreen
glbl.BackColor = vbGreen
i = clb(r)
i = dc()

End Sub

Private Sub Label2_Click()
On Error GoTo error
If glbl.Caption > 0 Then
If clbl.Caption > 0 Then

Label2.Caption = glbl / clbl
End If
End If
error:
Err.Clear
End Sub





Private Sub sbj_Change(j As Integer)


Label3.BackColor = vbGreen
Label3.Caption = "G.P:"




k = sbj(j).Text
Select Case k
Case Is >= 101
Label3.BackColor = vbRed
gp(j).Caption = ""
n(j) = 0

 Case Is >= 85
gp(j).Caption = "4.00     A+"
i = dc()
n(j) = 4
Case Is >= 80
i = dc()
gp(j).Caption = "3.70      A"
n(j) = 3.7

Case Is >= 75
gp(j).Caption = "3.30     B+"
i = dc()
n(j) = 3.3

Case Is >= 70
gp(j).Caption = "3.00      B"
i = dc()
n(j) = 3

Case Is >= 65
i = dc()
gp(j).Caption = "2.70     B-"
n(j) = 2.7

Case Is >= 61
i = dc()
gp(j).Caption = "2.30     C+"
n(j) = 2.3
 
Case Is >= 58
i = dc()
gp(j).Caption = "2.00      C"
n(j) = 2

Case Is >= 55
i = dc()
gp(j).Caption = "1.70      C-"
n(j) = 1.7

Case Is >= 53
i = dc()
gp(j).Caption = "1.30     D+"
n(j) = 1.3

Case Is >= 50
i = dc()
gp(j).Caption = "1.00      D"
n(j) = 1
 
Case Is >= 11
gp(j).Caption = "Fail!      F"
gp(j).BackColor = vbGreen
Label3.BackColor = vbRed
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
i = clb(r)
For k = 0 To 4
Select Case 1
Case Is = Len(ch(k).Text)
If sbj(k + 1).Text = "" Then
sbj(k + 1).SetFocus
End If
End Select
Next

End Sub


Function dc() As Integer
Label6(i).BackColor = vbGreen
Form1.BackColor = vbYellow
Label3.BackColor = vbRed
For i = 0 To 5

gp(i).BackColor = vbGreen
sbj(i).BackColor = vbWhite

Next
Label1.BackColor = vbRed
End Function

Function clb(r) As Integer
clbl.Caption = ""
b = ch(r).Text
For i = 0 To 5
Select Case b

Case Is = ""
ch(r).Text = ""
clbl.Caption = ""

Case Is < 100
clbl.Caption = Val(clbl.Caption) + Val(ch(i))

Case Else
ch(r).Text = ""
clbl.Caption = ""
End Select
Next
End Function



Private Sub Timer1_Timer()

On Error GoTo error
For k = 0 To 5
i = tgp(k)
i = glb(k)
Label6(k).BackColor = vbGreen
Next

b = clbl.Caption
Select Case b
Case Is > 0

If glbl.Caption > 0 Then

r = glbl / clbl

Text1.Text = r
End If


Case Is < 0
Label2.Caption = ""
End Select
Form1.Caption = "CGPA Calculator                                                   " & Time & "                                      Naeem"

M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
i = sSgp(0)
b = Text1.Text

error:

End Sub

Private Sub s1_Click()
Form1.Hide
Form1.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = M1.sSgp(0)
End Sub
Private Sub s2_Click()
Form1.Hide
Form2.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = M1.sSgp(0)
End Sub
Private Sub s3_Click()
Form1.Hide
Form3.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = M1.sSgp(0)
End Sub
Private Sub s4_Click()
Form1.Hide
Form4.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = M1.sSgp(0)
End Sub
Private Sub s5_Click()
Form1.Hide
Form5.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = M1.sSgp(0)
End Sub
Private Sub s6_Click()
Form1.Hide
Form6.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = M1.sSgp(0)
End Sub
Private Sub s7_Click()
Form1.Hide
Form7.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = M1.sSgp(0)
End Sub
Private Sub s8_Click()
Form1.Hide
Form8.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = M1.sSgp(0)
End Sub





Private Sub sCmd_Click(i As Integer)
On Error GoTo error
If i = 0 Then
Form1.Hide
Form1.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = M1.sSgp(0)
Form1.Frame3.Visible = True
ElseIf i = 1 Then
Form1.Hide
Form2.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = M1.sSgp(0)
ElseIf i = 2 Then
Form1.Hide
Form3.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = M1.sSgp(0)
ElseIf i = 3 Then
Form1.Hide
Form4.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = sSgp(0)
ElseIf i = 4 Then
Form1.Hide
Form5.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = sSgp(0)
ElseIf i = 5 Then
Form1.Hide
Form6.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = sSgp(0)
ElseIf i = 6 Then
Form1.Hide
Form7.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = sSgp(0)
ElseIf i = 7 Then
Form1.Hide
Form8.Show
M1.f(0) = glbl.Caption
M1.h(0) = clbl.Caption
j = sSgp(0)

End If
error:

End Sub
