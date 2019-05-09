VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdisiiss 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema Integrado de Información Sismológica"
   ClientHeight    =   6810
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9210
   Icon            =   "mdisiiss.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdisiiss.frx":08CA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6720
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisiiss.frx":3B6ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisiiss.frx":3B7007
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisiiss.frx":3B7469
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisiiss.frx":3B78C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisiiss.frx":3B7D2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LOC"
            Object.ToolTipText     =   "Búsqueda por Coordenadas y/o Fechas"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MAG"
            Object.ToolTipText     =   "Búsqueda por Magnitudes y/o Fechas"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PROF"
            Object.ToolTipText     =   "Búsqueda por Profundidad y/o Fecha"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LOCMAG"
            Object.ToolTipText     =   "Búsqueda Avanzada"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MAPA"
            Object.ToolTipText     =   "Mostrar Mapa"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SALIR"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisiiss.frx":3B833D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisiiss.frx":3B88A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisiiss.frx":3B8D21
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisiiss.frx":3B9176
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisiiss.frx":3B95DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdisiiss.frx":3B99AF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuarchivosalir 
         Caption         =   "Salir"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuvisualizaciones 
      Caption         =   "&Visualizaciones"
      Begin VB.Menu mnuvisualloc 
         Caption         =   "Localización"
      End
      Begin VB.Menu mnuvisualmag 
         Caption         =   "Magnitud"
      End
      Begin VB.Menu mnuvisuallocmag 
         Caption         =   "Localización Versus Magnitud"
      End
   End
   Begin VB.Menu proceso 
      Caption         =   "&Procesamiento"
      Begin VB.Menu seisan 
         Caption         =   "Archivo SEISAN"
      End
   End
End
Attribute VB_Name = "mdisiiss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
completa = False
End Sub

Private Sub mnuarchivosalir_Click()
End
End Sub

Private Sub mnuvisualloc_Click()
frmsiisvloc.Show
End Sub

Private Sub mnuvisuallocmag_Click()
frmsiisv.Show
End Sub

Private Sub mnuvisualmag_Click()
frmsiisvmag.Show
End Sub

Private Sub seisan_Click()
frm_buscar_archivo.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "LOC"
    Unload frmsiisv
    Unload frmsiisvmag
    Unload frmsiisvprof
    Unload frmsiismapa
    frmsiisvloc.Show
    
Case "MAG"
    Unload frmsiisv
    Unload frmsiisvloc
    Unload frmsiisvprof
     Unload frmsiismapa
    frmsiisvmag.Show
    
Case "LOCMAG"
    Unload frmsiisvmag
    Unload frmsiisvloc
    Unload frmsiisvprof
     Unload frmsiismapa
    frmsiisv.Show
    
Case "PROF"
    Unload frmsiisvmag
    Unload frmsiisvloc
    Unload frmsiisv
     Unload frmsiismapa
    frmsiisvprof.Show
    
Case "MAPA"
    Unload frmsiisvmag
    Unload frmsiisvloc
    Unload frmsiisv
    frmsiismapa.Show
    
Case "SALIR"
    End
    
End Select
End Sub
