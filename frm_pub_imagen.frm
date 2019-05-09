VERSION 5.00
Begin VB.Form frm_pub_imagen 
   Caption         =   "Publicidad Zoom"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame16 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.Image imgProf 
         BorderStyle     =   1  'Fixed Single
         Height          =   6495
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   10335
      End
   End
End
Attribute VB_Name = "frm_pub_imagen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

If foto_pub = "editar" Then
    imgProf.Picture = LoadPicture(frm_pub_editar.txt_path.Text)
End If

If foto_pub = "liquidar" Then
    imgProf.Picture = LoadPicture(frm_pub_liqui_anual.txt_path.Text)
End If

End Sub

Private Sub imgProf_Click()
Unload Me
End Sub
