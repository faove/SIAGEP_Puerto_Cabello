VERSION 5.00
Begin VB.Form Atencion_Contri_Despachador 
   Caption         =   "Atención al Contribuyente"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   8655
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   5775
      Begin VB.CommandButton cmdPic 
         Caption         =   "Patente Industria y Comercio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdVeh 
         Caption         =   "Vehiculos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdInm 
         Caption         =   "Inmuebles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdPub 
         Caption         =   "Publicidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton CmdOtros 
         Caption         =   "Tasas y Otros Tributos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         TabIndex        =   3
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         TabIndex        =   2
         Top             =   1680
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "Área  de  Atención  al  Contribuyente"
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
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   7695
   End
End
Attribute VB_Name = "Atencion_Contri_Despachador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdInm_Click()
    Unload Me
    Gid_obj = "INM"
    frm_inm_perfil.Show
End Sub

Private Sub cmdPic_Click()
frm_pic_perfil.Show
End Sub

Private Sub Form_Load()
    Me.Height = 5280
    Me.Width = 8850
End Sub

'   Gid_obj = "INM"
'    DoCmd.Close
'    DoCmd.OpenForm "CUM_INM_PERFIL"
