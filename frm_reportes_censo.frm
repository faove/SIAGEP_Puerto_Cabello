VERSION 5.00
Begin VB.Form frm_reportes_censo 
   Caption         =   "Reportes Generales del Censo 2003"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   8085
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   1200
      TabIndex        =   6
      Top             =   1320
      Width           =   5775
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "C&errar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2280
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmd_Sector 
         Caption         =   "Por &Sector"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3840
         TabIndex        =   2
         ToolTipText     =   "Reporte de Contribuyentes por Sector General"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmd_contri_con_pat 
         Caption         =   "Contribuyentes con &Patente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2280
         TabIndex        =   1
         ToolTipText     =   "Reporte de Contribuyentes con Patente por Sector"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmd_pub 
         Caption         =   "Publicidad no &Registrada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   720
         TabIndex        =   3
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmd_contri_sin_pat 
         Caption         =   "&Contribuyentes sin Patente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   720
         TabIndex        =   0
         ToolTipText     =   "Reporte de Contribuyentes sin Patente por Sector"
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "Área  de  Reportes del Censo 2003"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   7695
   End
End
Attribute VB_Name = "frm_reportes_censo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_contri_con_pat_Click()
censo_sector = 2
frm_rpt_censo_tab_sector.Show
End Sub

Private Sub cmd_contri_sin_pat_Click()
censo_sector = 1
frm_rpt_censo_tab_sector.Show
End Sub

Private Sub cmd_pub_Click()
censo_sector = 4
frm_rpt_censo_tab_sector.Show

End Sub

Private Sub cmd_Sector_Click()
censo_sector = 3
frm_rpt_censo_tab_sector.Show
End Sub

Private Sub cmdCerrar_Click()
censo_sector = 0
Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Height = 5280
    Me.Width = 8850
End Sub
