VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm Alcalsis 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema Automatizado de Gestión Pública "
   ClientHeight    =   9450
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13230
   Icon            =   "alcalsis.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "alcalsis.frx":08CA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer_Scroll 
      Left            =   3960
      Top             =   4680
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   5400
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":E947
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":F621
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":102FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":10FD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":11CAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":12989
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":13663
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":1433D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":15017
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":15CF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":169CB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   9180
      Width           =   13230
      _ExtentX        =   23336
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   12991
            MinWidth        =   1587
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   "BAR"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   706
            MinWidth        =   706
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "17/03/2010"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1535
            MinWidth        =   1535
            TextSave        =   "10:18 p.m."
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4560
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":176A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":1837F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":19059
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":19D33
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":1AA0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":1B6E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":1C3C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":1D09B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":1DD75
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":1EA4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":1F729
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
      Width           =   13230
      _ExtentX        =   23336
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EST"
            Description     =   "EST"
            Object.ToolTipText     =   "Estadística"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PIC"
            Description     =   "PIC"
            Object.ToolTipText     =   "Patente de Industria y Comercio"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "INM"
            Description     =   "INM"
            Object.ToolTipText     =   "Propiedad Inmobiliaria"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "VEH"
            Description     =   "VEH"
            Object.ToolTipText     =   "Patente Vehículo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PUB"
            Description     =   "PUB"
            Object.ToolTipText     =   "Publicidad"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TAS"
            Description     =   "TAS"
            Object.ToolTipText     =   "Tasas y Otros Tributos"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "REC"
            Description     =   "REC"
            Object.ToolTipText     =   "Recaudador"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PRO"
            Object.ToolTipText     =   "Procesos"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LIQ"
            Description     =   "Liquidación Genérica"
            Object.ToolTipText     =   "Liquidación"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "INF"
            Description     =   "Informe"
            Object.ToolTipText     =   "Informes"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Description     =   "Salir a Windows"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":20403
            Key             =   "VEH"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":210DD
            Key             =   "EST"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":21DB7
            Key             =   "INM"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":22A91
            Key             =   "PIC"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":2376B
            Key             =   "INF"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":24445
            Key             =   "PRO"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":2511F
            Key             =   "PUB"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":25DF9
            Key             =   "REC"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":26AD3
            Key             =   "SAL"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":277AD
            Key             =   "LIQ"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "alcalsis.frx":28487
            Key             =   "TAS"
         EndProperty
      EndProperty
   End
   Begin VB.Menu archivo 
      Caption         =   "&Archivo"
      Index           =   1
      Begin VB.Menu cerrar 
         Caption         =   "Cerrar Sesión"
      End
      Begin VB.Menu salir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu herramientas 
      Caption         =   "&Herramientas"
      Index           =   2
      Begin VB.Menu pic 
         Caption         =   "Patente de Industria y Comercio"
      End
      Begin VB.Menu INM 
         Caption         =   "Inmuebles Urbanos"
      End
      Begin VB.Menu VEH 
         Caption         =   "Patente de Vehículo"
      End
      Begin VB.Menu PUB 
         Caption         =   "Publicidad Comercial"
      End
      Begin VB.Menu TAS 
         Caption         =   "Tasas y Otros Tributos"
      End
   End
   Begin VB.Menu reporte 
      Caption         =   "&Reportes y Gestión"
      Index           =   3
      Begin VB.Menu estadistica 
         Caption         =   "Estadística Gestión"
      End
      Begin VB.Menu informes 
         Caption         =   "Informes de Recaudación"
      End
   End
   Begin VB.Menu reportegen 
      Caption         =   "Reportes Generales"
      Index           =   4
      Begin VB.Menu rpt_declara_jurada 
         Caption         =   "Reportes Declaración Jurada"
      End
      Begin VB.Menu rpt_censo_2003 
         Caption         =   "Reportes del Censo 2003"
         Begin VB.Menu contri_sin_pat 
            Caption         =   "&Contribuyentes sin Patente"
         End
         Begin VB.Menu contri_con_pat 
            Caption         =   "Contribuyentes con &Patente"
         End
         Begin VB.Menu por_sector 
            Caption         =   "Por &Sector"
         End
         Begin VB.Menu pub_no_registrada 
            Caption         =   "Publicidad no &Registrada"
         End
      End
      Begin VB.Menu rpt_modificaciones 
         Caption         =   "Reporte de Modificaciones"
         Begin VB.Menu rpt_modificaciones_todas 
            Caption         =   "Todas las Modificaciones Diarias"
         End
      End
      Begin VB.Menu rpt_notif 
         Caption         =   "Reporte de Notificaciones"
         Begin VB.Menu rpt_notif_inm 
            Caption         =   "Reporte de Notificaciones de INM"
         End
      End
   End
   Begin VB.Menu administracion 
      Caption         =   "A&dministración"
      Index           =   5
      Begin VB.Menu seguridad 
         Caption         =   "Usuarios del Sistema"
      End
      Begin VB.Menu password 
         Caption         =   "Password del Sistema (Reimpresión, Editar ...)"
      End
      Begin VB.Menu liberar_planilla 
         Caption         =   "Liberar Planilla"
      End
      Begin VB.Menu modificar_facturas 
         Caption         =   "Modificar Facturación"
      End
      Begin VB.Menu cargar_archivo_ban 
         Caption         =   "Cargar Archivo del Banco"
      End
   End
   Begin VB.Menu ayuda 
      Caption         =   "A&yuda"
      Index           =   6
      Begin VB.Menu contenido 
         Caption         =   "Contenido"
      End
      Begin VB.Menu acerca 
         Caption         =   "Acerca de ..."
      End
   End
End
Attribute VB_Name = "Alcalsis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sesion As String

Private Sub acerca_Click()
frm_acerca.Show
End Sub

Private Sub atencion_Click()
Atencion_Contri_Despachador.Show
End Sub

Private Sub censo_Click()
frm_censo.Show
End Sub



Private Sub cargar_archivo_ban_Click()
frm_alc_recaudacion_banco.Show
End Sub

Private Sub Cerrar_Click()
    Unload Me
    frm_inicio.Show
End Sub

Private Sub contenido_Click()
'frm_ayuda.Show
End Sub

Private Sub contri_con_pat_Click()
censo_sector = 2
frm_rpt_censo_tab_sector.Show
End Sub

Private Sub contri_sin_pat_Click()
censo_sector = 1
frm_rpt_censo_tab_sector.Show
End Sub

Private Sub estadistica_Click()
        If user_grupo = "04" Then
            frm_estadistica_gestion.Show
        Else
            MsgBox "Estadística y gestión solo esta destinado  " & Chr(13) & " a usuarios del grupo 04", vbInformation, "Seguridad -Alcalsis-"
            Exit Sub
        End If
End Sub

Private Sub informes_Click()
frm_informes_recaudacion.Show
End Sub

Private Sub INM_Click()
frm_inm_perfil.Show
End Sub

Private Sub liberar_planilla_Click()
frm_liberar_planilla.Show
End Sub

Private Sub MDIForm_Load()

actualizar_conex

Init_Globals

sesion = "Sesión Iniciada por: " & user_name

Me.Tag = sesion

Call Descripcion(Me.Tag)

'If Usuario <> 4 Then
'    Me.Toolbar1.Buttons.Item(11).Enabled = False
'End If

If user_grupo <> "04" Then
    'En el menu
    Me.informes.Enabled = False
    estadistica.Enabled = False
    'En el toolbar
    Me.Toolbar1.Buttons.Item(13).Enabled = False
    Me.Toolbar1.Buttons.Item(2).Enabled = False
    
End If

If user_grupo <> "01" Then
    administracion(5).Enabled = False
End If

End Sub

Private Sub prueba_Click()
Form1.Show
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Descripcion(Me.Tag)
End Sub

Private Sub modificar_facturas_Click()
frm_liberar_tablas.Show
End Sub

Private Sub password_Click()

ident = "SIA"

frm_seguridad_de_datos.Show


End Sub

Private Sub pic_Click()
frm_pic_perfil.Show
End Sub

Private Sub por_sector_Click()
censo_sector = 3
frm_rpt_censo_tab_sector.Show
End Sub

Private Sub PUB_Click()
frm_pub_perfil.Show
End Sub

Private Sub pub_no_registrada_Click()
censo_sector = 4
frm_rpt_censo_tab_sector.Show
End Sub

Private Sub recaudacion_Click()
    frm_alc_recaudador_micasa.Show
End Sub

Private Sub rpt_declara_jurada_Click()

'rpt_declara_jurada.
End Sub

Private Sub salir_Click()
End
End Sub

Private Sub seguridad_Click()
frm_seguridad_sistema.Show
End Sub

Private Sub TAS_Click()
frm_liquidacion_tasas.Show
End Sub

Private Sub Timer_Scroll_Timer()
Alcalsis.StatusBar1.Panels.Item(2).Text = ""
Timer_Scroll.Interval = 0
End Sub

'Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'
'Select Case Button.Key
'    Case "Salir"
'        MsgBox "mensaje"
'End Select
'
'
'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Impresora"
            MsgBox Printer.DeviceName
        Case "EST"
        If user_grupo = "04" Then
            frm_estadistica_gestion.Show
        Else
            MsgBox "Estadística y gestión solo esta destinado  " & Chr(13) & " a usuarios del grupo 04", vbInformation, "Seguridad -Alcalsis-"
            Exit Sub
        End If
        Case "PIC"
            frm_pic_perfil.Show
        Case "INM"
            frm_inm_perfil.Show
        Case "VEH"
            frm_veh_perfil.Show
        Case "PUB"
            frm_pub_perfil.Show
        Case "REC"
            frm_alc_recaudador_micasa.Show
        Case "LIQ"
            frm_liquidacion.Show
        Case "PRO"
            frm_procesos.Show 1
        Case "INF"
        If user_grupo = "04" Then
            frm_informes_recaudacion.Show
        Else
            MsgBox "Informe solo esta destinado " & Chr(13) & " a usuarios del grupo 04", vbInformation, "Seguridad -Alcalsis-"
            Exit Sub
        End If
        Case "Salir"
            End
        Case "TAS"
            frm_liquidacion_tasas.Show
    End Select
End Sub

Private Sub VEH_Click()
frm_veh_perfil.Show
End Sub
