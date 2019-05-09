VERSION 5.00
Object = "{0002E550-0000-0000-C000-000000000046}#1.0#0"; "OWC10.DLL"
Begin VB.Form frm_est_alc_Ingresos_Rubros_Fechas 
   Caption         =   "Ingresos Rubros Fechas"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   11790
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8175
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   11295
      Begin OWC10.ChartSpace ChartSpace1 
         Height          =   7095
         Left            =   0
         OleObjectBlob   =   "frm_ALC_Ingresos_Rubros_Fechas.frx":0000
         TabIndex        =   5
         Top             =   0
         Width           =   10815
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   9240
         TabIndex        =   4
         Tag             =   "Cerrar relación ingresos por rubros por caja"
         Top             =   7320
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   8295
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   "Relación Ingresos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   " Por Rubros Por Caja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frm_est_alc_Ingresos_Rubros_Fechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MATRIZ(4, 4, 4)

Dim rubros(95, 3)

Dim Lista_Rubros As String


Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = True
'    Call Descripcion(Me.cmd_cerrar.Tag)
End Sub

Private Sub Form_Resize()
    Call Mover_der(Me, Frame2, 0)
    Call Mover_centrado(Me, Frame1)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
'    Call Descripcion("")
End Sub

