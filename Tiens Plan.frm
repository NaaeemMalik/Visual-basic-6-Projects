VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox TB 
      Height          =   495
      Index           =   0
      Left            =   9960
      TabIndex        =   110
      Text            =   "0"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox IB 
      Height          =   495
      Index           =   0
      Left            =   8760
      TabIndex        =   97
      Text            =   "0"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox DB 
      Height          =   495
      Index           =   0
      Left            =   7560
      TabIndex        =   84
      Text            =   "0"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox TP 
      Height          =   495
      Index           =   0
      Left            =   6360
      TabIndex        =   72
      Text            =   "100"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox NP 
      Height          =   495
      Index           =   0
      Left            =   5160
      TabIndex        =   60
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox MP 
      Height          =   495
      Index           =   0
      Left            =   3960
      TabIndex        =   48
      Text            =   "1"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox NM 
      Height          =   495
      Index           =   0
      Left            =   2760
      TabIndex        =   35
      Text            =   "0"
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Ranks 
      Caption         =   "1*"
      Height          =   495
      Index           =   0
      Left            =   1680
      TabIndex        =   22
      Top             =   840
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1200
      Top             =   0
   End
   Begin VB.CommandButton Months 
      Caption         =   "11th"
      Height          =   495
      Index           =   10
      Left            =   480
      TabIndex        =   20
      Top             =   8040
      Width           =   735
   End
   Begin VB.CommandButton Months 
      Caption         =   "10th"
      Height          =   495
      Index           =   9
      Left            =   480
      TabIndex        =   19
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton Months 
      Caption         =   "9th"
      Height          =   495
      Index           =   8
      Left            =   480
      TabIndex        =   18
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Months 
      Caption         =   "8th"
      Height          =   495
      Index           =   7
      Left            =   480
      TabIndex        =   17
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton Months 
      Caption         =   "7th"
      Height          =   495
      Index           =   6
      Left            =   480
      TabIndex        =   16
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Months 
      Caption         =   "6th"
      Height          =   495
      Index           =   5
      Left            =   480
      TabIndex        =   15
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Months 
      Caption         =   "5th"
      Height          =   495
      Index           =   4
      Left            =   480
      TabIndex        =   14
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Months 
      Caption         =   "4th"
      Height          =   495
      Index           =   3
      Left            =   480
      TabIndex        =   13
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Months 
      Caption         =   "1st"
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   10
      Top             =   840
      Width           =   735
   End
   Begin VB.Frame Form1 
      Caption         =   "New PV"
      Height          =   8895
      Index           =   4
      Left            =   5040
      TabIndex        =   5
      Top             =   480
      Width           =   1215
      Begin VB.TextBox NP 
         Height          =   495
         Index           =   11
         Left            =   120
         TabIndex        =   71
         Top             =   8280
         Width           =   975
      End
      Begin VB.TextBox NP 
         Height          =   495
         Index           =   10
         Left            =   120
         TabIndex        =   70
         Top             =   7560
         Width           =   975
      End
      Begin VB.TextBox NP 
         Height          =   495
         Index           =   9
         Left            =   120
         TabIndex        =   69
         Top             =   6840
         Width           =   975
      End
      Begin VB.TextBox NP 
         Height          =   495
         Index           =   8
         Left            =   120
         TabIndex        =   68
         Top             =   6120
         Width           =   975
      End
      Begin VB.TextBox NP 
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   67
         Top             =   5400
         Width           =   975
      End
      Begin VB.TextBox NP 
         Height          =   495
         Index           =   6
         Left            =   120
         TabIndex        =   66
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox NP 
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   65
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox NP 
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   64
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox NP 
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   63
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox NP 
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   62
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox NP 
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Men Power"
      Height          =   8895
      Index           =   3
      Left            =   3840
      TabIndex        =   4
      Top             =   480
      Width           =   1215
      Begin VB.TextBox MP 
         Height          =   495
         Index           =   11
         Left            =   120
         TabIndex        =   59
         Text            =   "1"
         Top             =   8280
         Width           =   975
      End
      Begin VB.TextBox MP 
         Height          =   495
         Index           =   10
         Left            =   120
         TabIndex        =   58
         Text            =   "1"
         Top             =   7560
         Width           =   975
      End
      Begin VB.TextBox MP 
         Height          =   495
         Index           =   9
         Left            =   120
         TabIndex        =   57
         Text            =   "1"
         Top             =   6840
         Width           =   975
      End
      Begin VB.TextBox MP 
         Height          =   495
         Index           =   8
         Left            =   120
         TabIndex        =   56
         Text            =   "1"
         Top             =   6120
         Width           =   975
      End
      Begin VB.TextBox MP 
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   55
         Text            =   "1"
         Top             =   5400
         Width           =   975
      End
      Begin VB.TextBox MP 
         Height          =   495
         Index           =   6
         Left            =   120
         TabIndex        =   54
         Text            =   "1"
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox MP 
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   53
         Text            =   "1"
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox MP 
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   52
         Text            =   "1"
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox MP 
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   51
         Text            =   "1"
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox MP 
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   50
         Text            =   "1"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox MP 
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   49
         Text            =   "1"
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "    Months"
      Height          =   8895
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1215
      Begin VB.CommandButton Months 
         Caption         =   "12th"
         Height          =   495
         Index           =   11
         Left            =   240
         TabIndex        =   21
         Top             =   8280
         Width           =   735
      End
      Begin VB.CommandButton Months 
         Caption         =   "3rd"
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton Months 
         Caption         =   "2nd"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   9735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.Frame Frame2 
         Caption         =   "Total Bonus"
         Height          =   8895
         Index           =   8
         Left            =   9720
         TabIndex        =   9
         Top             =   480
         Width           =   1455
         Begin VB.TextBox TB 
            Height          =   495
            Index           =   12
            Left            =   120
            TabIndex        =   122
            Text            =   "0"
            Top             =   9000
            Width           =   1215
         End
         Begin VB.TextBox TB 
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   121
            Text            =   "0"
            Top             =   8280
            Width           =   1215
         End
         Begin VB.TextBox TB 
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   120
            Text            =   "0"
            Top             =   7560
            Width           =   1215
         End
         Begin VB.TextBox TB 
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   119
            Text            =   "0"
            Top             =   6840
            Width           =   1215
         End
         Begin VB.TextBox TB 
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   118
            Text            =   "0"
            Top             =   6120
            Width           =   1215
         End
         Begin VB.TextBox TB 
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   117
            Text            =   "0"
            Top             =   5400
            Width           =   1215
         End
         Begin VB.TextBox TB 
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   116
            Text            =   "0"
            Top             =   4680
            Width           =   1215
         End
         Begin VB.TextBox TB 
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   115
            Text            =   "0"
            Top             =   3960
            Width           =   1215
         End
         Begin VB.TextBox TB 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   114
            Text            =   "0"
            Top             =   3240
            Width           =   1215
         End
         Begin VB.TextBox TB 
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   113
            Text            =   "0"
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox TB 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   112
            Text            =   "0"
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox TB 
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   111
            Text            =   "0"
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Indirect Bonus"
         Height          =   8895
         Index           =   7
         Left            =   8520
         TabIndex        =   8
         Top             =   480
         Width           =   1215
         Begin VB.TextBox IB 
            Height          =   495
            Index           =   12
            Left            =   120
            TabIndex        =   109
            Text            =   "0"
            Top             =   9000
            Width           =   975
         End
         Begin VB.TextBox IB 
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   108
            Text            =   "0"
            Top             =   8280
            Width           =   975
         End
         Begin VB.TextBox IB 
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   107
            Text            =   "0"
            Top             =   7560
            Width           =   975
         End
         Begin VB.TextBox IB 
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   106
            Text            =   "0"
            Top             =   6840
            Width           =   975
         End
         Begin VB.TextBox IB 
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   105
            Text            =   "0"
            Top             =   6120
            Width           =   975
         End
         Begin VB.TextBox IB 
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   104
            Text            =   "0"
            Top             =   5400
            Width           =   975
         End
         Begin VB.TextBox IB 
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   103
            Text            =   "0"
            Top             =   4680
            Width           =   975
         End
         Begin VB.TextBox IB 
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   102
            Text            =   "0"
            Top             =   3960
            Width           =   975
         End
         Begin VB.TextBox IB 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   101
            Text            =   "0"
            Top             =   3240
            Width           =   975
         End
         Begin VB.TextBox IB 
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   100
            Text            =   "0"
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox IB 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   99
            Text            =   "0"
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox IB 
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   98
            Text            =   "0"
            Top             =   1080
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Direct Bonus"
         Height          =   8895
         Index           =   6
         Left            =   7320
         TabIndex        =   7
         Top             =   480
         Width           =   1215
         Begin VB.TextBox DB 
            Height          =   495
            Index           =   12
            Left            =   120
            TabIndex        =   96
            Top             =   9000
            Width           =   975
         End
         Begin VB.TextBox DB 
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   95
            Top             =   8280
            Width           =   975
         End
         Begin VB.TextBox DB 
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   94
            Top             =   7560
            Width           =   975
         End
         Begin VB.TextBox DB 
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   93
            Top             =   6840
            Width           =   975
         End
         Begin VB.TextBox DB 
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   92
            Top             =   6120
            Width           =   975
         End
         Begin VB.TextBox DB 
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   91
            Top             =   5400
            Width           =   975
         End
         Begin VB.TextBox DB 
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   90
            Top             =   4680
            Width           =   975
         End
         Begin VB.TextBox DB 
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   89
            Top             =   3960
            Width           =   975
         End
         Begin VB.TextBox DB 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   88
            Top             =   3240
            Width           =   975
         End
         Begin VB.TextBox DB 
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   87
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox DB 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   86
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox DB 
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   85
            Top             =   1080
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Total  PV"
         Height          =   8895
         Index           =   5
         Left            =   6120
         TabIndex        =   6
         Top             =   480
         Width           =   1215
         Begin VB.TextBox TP 
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   83
            Text            =   "100"
            Top             =   8280
            Width           =   975
         End
         Begin VB.TextBox TP 
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   82
            Text            =   "100"
            Top             =   7560
            Width           =   975
         End
         Begin VB.TextBox TP 
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   81
            Text            =   "100"
            Top             =   6840
            Width           =   975
         End
         Begin VB.TextBox TP 
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   80
            Text            =   "100"
            Top             =   6120
            Width           =   975
         End
         Begin VB.TextBox TP 
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   79
            Text            =   "100"
            Top             =   5400
            Width           =   975
         End
         Begin VB.TextBox TP 
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   78
            Text            =   "100"
            Top             =   4680
            Width           =   975
         End
         Begin VB.TextBox TP 
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   77
            Text            =   "100"
            Top             =   3960
            Width           =   975
         End
         Begin VB.TextBox TP 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   76
            Text            =   "100"
            Top             =   3240
            Width           =   975
         End
         Begin VB.TextBox TP 
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   75
            Text            =   "100"
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox TP 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   74
            Text            =   "100"
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox TP 
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   73
            Text            =   "100"
            Top             =   1080
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "New Men"
         Height          =   8895
         Index           =   2
         Left            =   2520
         TabIndex        =   3
         Top             =   480
         Width           =   1215
         Begin VB.TextBox NM 
            Height          =   495
            Index           =   12
            Left            =   120
            TabIndex        =   47
            Text            =   "0"
            Top             =   9000
            Width           =   975
         End
         Begin VB.TextBox NM 
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   46
            Text            =   "0"
            Top             =   8280
            Width           =   975
         End
         Begin VB.TextBox NM 
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   45
            Text            =   "0"
            Top             =   7560
            Width           =   975
         End
         Begin VB.TextBox NM 
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   44
            Text            =   "0"
            Top             =   6840
            Width           =   975
         End
         Begin VB.TextBox NM 
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   43
            Text            =   "0"
            Top             =   6120
            Width           =   975
         End
         Begin VB.TextBox NM 
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   42
            Text            =   "0"
            Top             =   5400
            Width           =   975
         End
         Begin VB.TextBox NM 
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   41
            Text            =   "0"
            Top             =   4680
            Width           =   975
         End
         Begin VB.TextBox NM 
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   40
            Text            =   "0"
            Top             =   3960
            Width           =   975
         End
         Begin VB.TextBox NM 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   39
            Text            =   "0"
            Top             =   3240
            Width           =   975
         End
         Begin VB.TextBox NM 
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   38
            Text            =   "0"
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox NM 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Text            =   "0"
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox NM 
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   36
            Text            =   "0"
            Top             =   1080
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "  Ranks"
         Height          =   8895
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   1215
         Begin VB.CommandButton Ranks 
            Caption         =   "5*"
            Height          =   495
            Index           =   12
            Left            =   240
            TabIndex        =   34
            Top             =   9000
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "12*"
            Height          =   495
            Index           =   11
            Left            =   240
            TabIndex        =   33
            Top             =   8280
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "11*"
            Height          =   495
            Index           =   10
            Left            =   240
            TabIndex        =   32
            Top             =   7560
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "10*"
            Height          =   495
            Index           =   9
            Left            =   240
            TabIndex        =   31
            Top             =   6840
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "9*"
            Height          =   495
            Index           =   8
            Left            =   240
            TabIndex        =   30
            Top             =   6120
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "8*"
            Height          =   495
            Index           =   7
            Left            =   240
            TabIndex        =   29
            Top             =   5400
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "7*"
            Height          =   495
            Index           =   6
            Left            =   240
            TabIndex        =   28
            Top             =   4680
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "6*"
            Height          =   495
            Index           =   5
            Left            =   240
            TabIndex        =   27
            Top             =   3960
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "5*"
            Height          =   495
            Index           =   4
            Left            =   240
            TabIndex        =   26
            Top             =   3240
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "4*"
            Height          =   495
            Index           =   3
            Left            =   240
            TabIndex        =   25
            Top             =   2520
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "3*"
            Height          =   495
            Index           =   2
            Left            =   240
            TabIndex        =   24
            Top             =   1800
            Width           =   735
         End
         Begin VB.CommandButton Ranks 
            Caption         =   "2*"
            Height          =   495
            Index           =   1
            Left            =   240
            TabIndex        =   23
            Top             =   1080
            Width           =   735
         End
         Begin VB.Shape Shape1 
            Height          =   375
            Left            =   480
            Top             =   1920
            Width           =   495
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   3645
      Left            =   5520
      Picture         =   "Tiens Plan.frx":0000
      Top             =   4560
      Width           =   4800
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   10320
      Y1              =   1440
      Y2              =   2640
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
Frame1.FontSize = 18
Frame1.Caption = "    " & Time & "                   " & Date
End Sub
