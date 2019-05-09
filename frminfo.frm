VERSION 5.00
Begin VB.Form frminfo 
   BackColor       =   &H80000009&
   Caption         =   "Datos del  Sismo"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List5 
      Height          =   2595
      Left            =   8520
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox List4 
      Height          =   2595
      Left            =   7560
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox List3 
      Height          =   2595
      Left            =   6600
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   5760
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   5040
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblmag1 
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblprof1 
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label lbllon1 
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label lbllat1 
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lblfecha1 
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblprof 
      Caption         =   "Profundidad:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lbllon 
      Caption         =   "Longitud:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lbllat 
      Caption         =   "Latitud:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblmag 
      Caption         =   "Magnitud:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblfecha 
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frminfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i
Dim X

X = num - 1
For i = 0 To num - 1
    Me.List1.AddItem fech(X)
    X = X - 1
Next
lblfecha1.Caption = List1.List(posicionlista)

X = num - 1
For i = 0 To num - 1
    Me.List2.AddItem lat(X)
    X = X - 1
Next
lbllat1.Caption = List2.List(posicionlista)

X = num - 1
For i = 0 To num - 1
    Me.List3.AddItem lon(X)
    X = X - 1
Next
lbllon1.Caption = List3.List(posicionlista)

X = num - 1
For i = 0 To num - 1
    Me.List4.AddItem mag(X)
    X = X - 1
Next
lblmag1.Caption = List4.List(posicionlista)

X = num - 1
For i = 0 To num - 1
    Me.List5.AddItem prof(X)
    X = X - 1
Next
lblprof1.Caption = List5.List(posicionlista)

End Sub

