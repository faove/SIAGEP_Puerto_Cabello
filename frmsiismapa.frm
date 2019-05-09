VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9BD6A640-CE75-11D1-AF04-204C4F4F5020}#2.0#0"; "Mo20.ocx"
Object = "{C7FC2F7C-0688-11D5-B2F8-000102D87123}#1.0#0"; "MO21Legend.ocx"
Object = "{6C20C089-0689-11D5-B2F8-000102D87123}#2.0#0"; "MO21ScaleBar.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmsiismapa 
   Caption         =   "Reporte"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdbuscar 
      Caption         =   "Cargar Capa"
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   8040
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   5325
      ItemData        =   "frmsiismapa.frx":0000
      Left            =   13200
      List            =   "frmsiismapa.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   6495
      ItemData        =   "frmsiismapa.frx":0004
      Left            =   12720
      List            =   "frmsiismapa.frx":0006
      TabIndex        =   14
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9480
      TabIndex        =   12
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   2880
      TabIndex        =   7
      Top             =   7800
      Width           =   6135
      Begin VB.CommandButton cmd_reporte 
         Caption         =   "Reporte"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdborrar 
         Caption         =   "Borrar Capas"
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdpuntos 
         Caption         =   "Graficar Puntos"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdborrarptos 
         Caption         =   "Borrar Puntos"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1545
         Left            =   3480
         Picture         =   "frmsiismapa.frx":0008
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
   End
   Begin MO21ScaleBar.ScaleBar ScaleBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   6960
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   873
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleBarUnits   =   3
      ScreenUnits     =   1
   End
   Begin MO21legend.legend legend1 
      Height          =   6495
      Left            =   10800
      TabIndex        =   5
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   11456
      BackColor       =   14737632
      ForeColor       =   -2147483630
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
      Left            =   11040
      Top             =   7680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsiismapa.frx":C01A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsiismapa.frx":C12C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsiismapa.frx":C23E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsiismapa.frx":C7A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsiismapa.frx":C8B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsiismapa.frx":CBD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsiismapa.frx":CFC8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Zoom"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mover Mapa"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ubicar Latitud y Longitud"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Colocar Marca"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "datos del sismo"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Crear Archivo JPG"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mapa Completo"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   11160
      Top             =   9360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11160
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      InitDir         =   "C:\PROYECTO\Imágenes\Regi"
   End
   Begin VB.Frame frmpredet 
      Caption         =   "Búsqueda Predeterminada"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   7800
      Width           =   2655
      Begin VB.CommandButton cmdmapa 
         Caption         =   "Mostrar Mapa"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdcapa 
         Caption         =   "Mostrar Fallas"
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
      End
   End
   Begin MapObjects2.Map Map1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   10695
      _Version        =   131072
      _ExtentX        =   18865
      _ExtentY        =   11456
      _StockProps     =   225
      BackColor       =   16711680
      BorderStyle     =   1
      Appearance      =   1
      BackColor       =   16711680
      Contents        =   "frmsiismapa.frx":D159
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ISLA LA BLANQUILLA"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frmsiismapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim variable As Boolean

 Dim rect As mapobjects2.Rectangle

 Dim layer As MapLayer
 Dim Loc As New mapobjects2.point
 Dim DC As New DataConnection
 Dim textsym As New mapobjects2.TextSymbol
 Dim newEvt As New mapobjects2.GeoEvent
 Dim newpunto As New mapobjects2.GeoEvent
 Dim foundEvt As mapobjects2.GeoEvent
 
 'texto
Private collGtextStrings As New VBA.Collection
Private collGtextPoints As New VBA.Collection
Private symGtext As New mapobjects2.TextSymbol

Private Sub cmd_reporte_Click()
Map1.ExportMapToJpeg "C:\PROYECTO\Imágenes\Región NorOriental\boletin.jpg"

frmsiisboletin.Show

End Sub

Private Sub cmdborrar_Click()
 On Error GoTo ControlError

  ' Clear all map layers from map
  Map1.Layers.Remove 0
  legend1.RemoveAll
  legend1.setMapSource Map1
  legend1.LoadLegend True
Exit Sub   ' Salir para evitar el controlador.
ControlError:   ' Rutina de control de errores.
   Select Case Err.Number   ' Evalúa el número de error.
      Case 5002   ' Error "Archivo ya está abierto".
         MsgBox "No existen capas para eliminar"   ' Cierra el archivo abierto.
      Case Else
      ' Puede incluir aquí otras situaciones...
            Exit Sub
   End Select
   
   legend1.RemoveAll
  legend1.setMapSource Map1
  legend1.LoadLegend True
End Sub

Private Sub cmdborrarptos_Click()
Map1.TrackingLayer.ClearEvents
Image1.Visible = False
List1.Clear
List2.Clear

End Sub

Private Sub cmdmapa_Click()

  Map1.Layers.Clear

DC.Database = "C:\PROYECTO\Imágenes\Región NorOriental\capas divididas"
Me.Map1.BackColor = RGB(75, 200, 255)   'vbBlue
Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("anzoategui")
layer.Symbol.Color = RGB(255, 204, 162)
Map1.Layers.Add layer

Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("delta")
layer.Symbol.Color = RGB(241, 255, 147)
Map1.Layers.Add layer

Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("guarico")
layer.Symbol.Color = RGB(170, 207, 238)
Map1.Layers.Add layer

Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("islas")
layer.Symbol.Color = RGB(216, 245, 141)
Map1.Layers.Add layer

Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("miranda")
layer.Symbol.Color = RGB(255, 130, 154)
Map1.Layers.Add layer

Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("monagas")
layer.Symbol.Color = RGB(255, 252, 162)
Map1.Layers.Add layer

Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("nva esparta")
layer.Symbol.Color = RGB(156, 143, 255)
Map1.Layers.Add layer

Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("sucre")
layer.Symbol.Color = RGB(125, 212, 143)
Map1.Layers.Add layer

Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("trinidad")
layer.Symbol.Color = RGB(244, 155, 191)
Map1.Layers.Add layer
Set Map1.Extent = Map1.FullExtent

  legend1.RemoveAll
  legend1.setMapSource Map1
  legend1.LoadLegend True

End Sub



Private Sub cmdsalir_Click()
Unload Me
End Sub
Private Sub cmdpuntos_Click()
  Dim point As New point
  Dim X, Y, i, j, VAR
  Dim salto, n
  Dim sw
  Dim swap, swaplat, swaplon, swapfech, swapprof
      Set newEvt = Map1.TrackingLayer.AddEvent(point, VAR)

  
  For i = 0 To num - 1
  
List1.AddItem fech(i)
  Next

  '''''''''''''''''''''''''''''''
  n = num
    salto = num
  Do While salto > 1
  salto = salto / 2
  Do
  sw = 1
  
  
  For j = 0 To n - salto
  i = j + salto
  If mag(j) < mag(i) Then
  ' magnitudes
  swap = mag(i)
  mag(i) = mag(j)
  mag(j) = swap
 
  ' latitudes
  swaplat = lat(i)
  lat(i) = lat(j)
  lat(j) = swaplat
  ' longitudes
  swaplon = lon(i)
  lon(i) = lon(j)
  lon(j) = swaplon
  'fechas
  swapfech = fech(i)
  fech(i) = fech(j)
  fech(j) = swapfech
  
  'profundidad
  swapprof = prof(i)
  prof(i) = prof(j)
  prof(j) = swapprof
  
  sw = 0
  End If
  Next j
  Loop Until sw = 1
  Loop
  
'     Map1.TrackingLayer.SymbolCount = 5

  'put a GeoEvent object on the TrackingLayer
  For i = 0 To num - 1
    

    
    If mag(i) < 3 Then
     With Map1.TrackingLayer.Symbol(0)
     .Style = moCircleMarker
    .Color = moGreen
    .Size = 3
    VAR = 0
    End With

    ElseIf mag(i) < 4 Then
    With Map1.TrackingLayer.Symbol(1)
    .Style = moCircleMarker
    .Color = moYellow
    .Size = 5
    VAR = 1
    End With

    ElseIf mag(i) < 5 Then
    With Map1.TrackingLayer.Symbol(2)
    .Style = moCircleMarker
    .Color = moOrange
    .Size = 8
    VAR = 2
    End With
    
    ElseIf mag(i) >= 5 Then
    With Map1.TrackingLayer.Symbol(3)
    .Style = moCircleMarker
    .Color = moRed
    .Size = 11
    VAR = 3
    End With
    
    End If
   
 X = lon(i)
 Y = lat(i)
    point.X = X
    point.Y = Y
    Set newEvt = Map1.TrackingLayer.AddEvent(point, VAR)
      
    newEvt.Tag = fech(i)
    List2.AddItem newEvt.Tag, 0
    Next
    
    
        Image1.Visible = True
End Sub
Private Sub cmdbuscar_Click()
On Error GoTo ControlError
    
  Dim DC As New mapobjects2.DataConnection
  Dim gds As mapobjects2.GeoDataset
  Dim FName As String
  Dim X, Y, VAR
  Dim fnt As New StdFont
Dim point As New point

 
  
  
  'Set up dailog box to prompt user to load a shapefile
  CommonDialog1.Filter = "ESRI Shapefiles (*.shp)|*.shp|ArcINFO Coverages (*.adf)| aat.adf;pat.adf"
  'Set cancel error so that if cancel is used then we can trap it an exit
  CommonDialog1.CancelError = True
  CommonDialog1.ShowOpen
  If Len(CommonDialog1.FileName) = 0 Then Exit Sub

   
  
      'Set up the DataConnection
      DC.Database = CurDir
      If Not DC.Connect Then Exit Sub
    
      'Get the dialog's returned filename
      FName = Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4)
      
      Set layer = New MapLayer
      Set layer.GeoDataset = DC.FindGeoDataset(FName)
      
      If FName = "LAGUNAS" Then
          layer.Symbol.Color = moBlue
          
      ElseIf FName = "ESTACIONES" Then
          layer.Symbol.Color = moGreen
        
       fnt.Name = "ESRI Cartography" 'TrueType font

      With Map1.TrackingLayer.Symbol(5)
            .Font = fnt
            .Size = 15
            
            .Style = moTrueTypeMarker
            .Color = moRed
            .CharacterIndex = 214
            VAR = 5
     End With
     
   'Colocar el geoevento en el mapa

 Y = 11.0729163311932
 X = -63.8886135110829
    point.X = X
    point.Y = Y
    Set newpunto = Map1.TrackingLayer.AddEvent(point, VAR)
      
      
 Y = 10.2065080631042
 X = -64.4397522098313
    point.X = X
    point.Y = Y
    Set newpunto = Map1.TrackingLayer.AddEvent(point, VAR)
    
    
 Y = 10.4300288062267
 X = -64.1960533911236
    point.X = X
    point.Y = Y
    Set newpunto = Map1.TrackingLayer.AddEvent(point, VAR)
    

 Y = 10.5660155426077
 X = -64.1889253202087
    point.X = X
    point.Y = Y
    Set newpunto = Map1.TrackingLayer.AddEvent(point, VAR)
    
 Y = 10.1578966184023
 X = -63.8266663319684
    point.X = X
    point.Y = Y
    Set newpunto = Map1.TrackingLayer.AddEvent(point, VAR)

 Y = 10.5506828533912
 X = -63.3221965814761
    point.X = X
    point.Y = Y
    Set newpunto = Map1.TrackingLayer.AddEvent(point, VAR)
    
 Y = 10.6006839019541
 X = -63.0702395626927
    point.X = X
    point.Y = Y
    Set newpunto = Map1.TrackingLayer.AddEvent(point, VAR)
    
 Y = 10.6734630228432
 X = -62.6518849387646
    point.X = X
    point.Y = Y
    Set newpunto = Map1.TrackingLayer.AddEvent(point, VAR)
    
 Y = 10.1185899857082
 X = -63.1124639121818
    point.X = X
    point.Y = Y
    Set newpunto = Map1.TrackingLayer.AddEvent(point, VAR)
          
      ElseIf FName = "AEROPUERTOS" Then
     '  Dim point As New point
'  Dim X, Y, VAR

    '   Dim fnt As New StdFont
       fnt.Name = "Wingdings" 'TrueType font

      With Map1.TrackingLayer.Symbol(4)
            .Font = fnt
            .Size = 18
            
            .Style = moTrueTypeMarker
            .Color = moRed
            .CharacterIndex = 81
            VAR = 4
     End With
     
  'Colocar el geoevento en el mapa
    
 X = -63.2509878791914
 Y = 10.6786973752402
    point.X = X
    point.Y = Y
    Set newpunto = Map1.TrackingLayer.AddEvent(point, VAR)
      
      
 X = -62.2969070653389
 Y = 10.5860093673229
    point.X = X
    point.Y = Y
    Set newpunto = Map1.TrackingLayer.AddEvent(point, VAR)
    
    
 X = -64.711166791575
 Y = 10.1322385500448
    point.X = X
    point.Y = Y
    Set newpunto = Map1.TrackingLayer.AddEvent(point, VAR)
    

X = -64.178268421248
Y = 10.4447975963668
    point.X = X
    point.Y = Y
    Set newpunto = Map1.TrackingLayer.AddEvent(point, VAR)
    
X = -63.9867656053483
Y = 10.9164810156892
    point.X = X
    point.Y = Y
    Set newpunto = Map1.TrackingLayer.AddEvent(point, VAR)
    
        Image1.Visible = True
      End If

    
      
      layer.Visible = True
      Map1.Layers.Add layer
      
      legend1.RemoveAll
  legend1.setMapSource Map1
  legend1.LoadLegend True
  
      Exit Sub   ' Salir para evitar el controlador.
ControlError:   ' Rutina de control de errores.
   Select Case Err.Number   ' Evalúa el número de error.
      Case 424   ' Error "Archivo ya está abierto".
         MsgBox "No existen capas para eliminar"   ' Cierra el archivo abierto.
      Case Else
      ' Puede incluir aquí otras situaciones...
            Exit Sub
   End Select
  
End Sub

Private Sub cmdcapa_Click()
' Carga los datos en el mapa
  Dim DC As New DataConnection
  DC.Database = "C:\PROYECTO\Imágenes\Región NorOriental"
  If Not DC.Connect Then End
  
  Set layer = New MapLayer
  Set layer.GeoDataset = DC.FindGeoDataset("FALLAS")
  layer.Visible = True
    layer.Symbol.Color = moBlack

  Map1.Layers.Add layer
  legend1.RemoveAll
  legend1.setMapSource Map1
  legend1.LoadLegend True
End Sub

Private Sub Command1_Click()

Dim strGText As String
Dim ptGText As mapobjects2.point
Dim X, Y

X = 3200
Y = 2200
strGText = "Isla de Margarita"
Set ptGText = Map1.ToMapPoint(X, Y)
collGtextStrings.Add strGText
collGtextPoints.Add ptGText
      
X = 5500
Y = 700
strGText = "Mar Caribe"
Set ptGText = Map1.ToMapPoint(X, Y)
collGtextStrings.Add strGText
collGtextPoints.Add ptGText

X = 2000
Y = 700
strGText = "Isla La Blanquilla"
Set ptGText = Map1.ToMapPoint(X, Y)
collGtextStrings.Add strGText
collGtextPoints.Add ptGText

Map1.TrackingLayer.Refresh True
End Sub

Private Sub Form_Load()
variable = True
 
  
 
DC.Database = "C:\PROYECTO\Imágenes\Región NorOriental\capas divididas"
Me.Map1.BackColor = RGB(75, 200, 255)   'vbBlue


Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("islas")
layer.Symbol.Color = RGB(216, 245, 141)
Map1.Layers.Add layer

Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("nva esparta")
layer.Symbol.Color = RGB(156, 143, 255)
Map1.Layers.Add layer

Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("trinidad")
layer.Sy mbol.Color = RGB(244, 155, 191)
Map1.Layers.Add layer

Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("sucre")
layer.Symbol.Color = RGB(125, 212, 143)
Map1.Layers.Add layer

Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("monagas")
layer.Symbol.Color = RGB(255, 252, 162)
Map1.Layers.Add layer

Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("anzoategui")
layer.Symbol.Color = RGB(255, 204, 162)
Map1.Layers.Add layer

Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("guarico")
layer.Symbol.Color = RGB(170, 207, 238)
Map1.Layers.Add layer

Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("miranda")
layer.Symbol.Color = RGB(255, 130, 154)
Map1.Layers.Add layer

Set layer = New MapLayer
Set layer.GeoDataset = DC.FindGeoDataset("delta")
layer.Symbol.Color = RGB(241, 255, 147)
Map1.Layers.Add layer




Set Map1.Extent = Map1.FullExtent

legend1.setMapSource Map1
  legend1.LoadLegend True

    Map1.TrackingLayer.SymbolCount = 6

End Sub


Private Sub Form_Unload(Cancel As Integer)
Erase lon
Erase lat
Erase mag
Erase prof
num = 0


End Sub



Private Sub legend1_AfterSetLayerVisible(index As Integer, isVisible As Boolean)
 Map1.Refresh
End Sub

Private Sub List1_Click()
Dim i
 'Find the GeoEvent with the selected Tag
 For i = 0 To num - 1
        If (List1.List(List1.ListIndex) = List2.List(i)) Then
 
'            If List2.ListIndex < 0 Then
'                MsgBox "Seleccione un campo de la lista"
'                Exit Sub
'            End If
  
            Set foundEvt = Map1.TrackingLayer.FindEvent(List2.List(i))
                       
            If Not foundEvt Is Nothing Then
                'If more than one event have the same tag, the first event with the
                'tag string will be displayed
                'foundEvt.SymbolIndex = 0

                Map1.FlashShape foundEvt.Shape, 3
                
            Else
                MsgBox "Evento no encontrado"
            End If
        End If
        
 Next
End Sub

Private Sub Map1_AfterLayerDraw(ByVal index As Integer, ByVal canceled As Boolean, ByVal hDC As stdole.OLE_HANDLE)
'Dim fnt As New StdFont
'fnt.Name = "Arial"
'  fnt.Size = 10
'  Set textsym.Font = fnt
'  With textsym
'    .Color = moBlack
'    .Font.Bold = True
'  End With
'  Map1.DrawText "Mar Caribe", Map1.ToMapPoint(5500, 700), textsym
'  Map1.DrawText "Isla de Margarita", Map1.ToMapPoint(3200, 2200), textsym
'  Map1.DrawText "Isla La Blanquilla", Map1.ToMapPoint(2000, 700), textsym



End Sub

Private Sub Map1_AfterTrackingLayerDraw(ByVal hDC As stdole.OLE_HANDLE)
  Dim i As Long
If collGtextStrings.Count > 0 Then
  For i = 1 To collGtextStrings.Count
    Map1.DrawText collGtextStrings(i), collGtextPoints(i), symGtext
  Next
End If
 With ScaleBar1.MapExtent
    .MinX = Map1.Extent.Left
    .MinY = Map1.Extent.Bottom
    .MaxX = Map1.Extent.Right
    .MaxY = Map1.Extent.Top
  End With
  '
  ' Set the ScaleBar's PageExtent.
  '
  With ScaleBar1.PageExtent
    .MinX = Map1.Left / Screen.TwipsPerPixelX
    .MinY = Map1.Top / Screen.TwipsPerPixelY
    .MaxX = (Map1.Left + Map1.Width) / Screen.TwipsPerPixelX

    .MaxY = (Map1.Top + Map1.Height) / Screen.TwipsPerPixelY
  End With
  '
  ' Refresh the ScaleBar after the Map has changed.
  '
  ScaleBar1.Refresh
'  ScaleBar1.

If variable = True Then

Dim strGText As String
Dim ptGText As mapobjects2.point
Dim X, Y

X = 3200
Y = 2200
strGText = "Isla de Margarita"
Set ptGText = Map1.ToMapPoint(X, Y)
collGtextStrings.Add strGText
collGtextPoints.Add ptGText
      
X = 5500
Y = 700
strGText = "Mar Caribe"
Set ptGText = Map1.ToMapPoint(X, Y)
collGtextStrings.Add strGText
collGtextPoints.Add ptGText

X = 2000
Y = 700
strGText = "Isla La Blanquilla"
Set ptGText = Map1.ToMapPoint(X, Y)
collGtextStrings.Add strGText
collGtextPoints.Add ptGText
variable = False

Map1.TrackingLayer.Refresh True
End If

  End Sub

Private Sub Map1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim rect As mapobjects2.Rectangle


If Toolbar1.Buttons(1).Value = 1 Then ' zoom
'  Map1.DrawText "Mar Caribe", Map1.ToMapPoint(5500, 700), textsym
'  Map1.DrawText "Isla de Margarita", Map1.ToMapPoint(3200, 2200), textsym
'  Map1.DrawText "Isla La Blanquilla", Map1.ToMapPoint(2000, 700), textsym
 '  Map1.ToMapPoint.shapeType
    Set rect = Map1.Extent
    If Button = 1 Then 'zoom in
        Map1.Extent = Map1.TrackRectangle
 
    Else 'zoom out
        rect.ScaleRectangle (1.5)
        Map1.Extent = rect
    End If
ElseIf Toolbar1.Buttons(2).Value = 1 Then ' pan
      Map1.MousePointer = moPan
 
    Map1.Pan
    
ElseIf Toolbar1.Buttons(3).Value = 1 Then  ' LAT LON

  
  'Get the location of the mouse-click in map units
  If Shift = 0 Then
      
    'If the form has units which are not Twips, then we should first convert
    'the X and Y coordinates to twips before passing them to the ToMapPoint method
    If frmsiismapa.ScaleMode <> vbTwips Then
      X = frmsiismapa.ScaleX(X, vbTwips, frmsiismapa.ScaleMode)

      Y = frmsiismapa.ScaleY(Y, vbTwips, frmsiismapa.ScaleMode)
    End If
    
    'Convert the twips value to Map units
    Set Loc = Map1.ToMapPoint(X, Y)
    MsgBox "Latitud:( " & Format(Loc.Y) & "), Longitud:(" + Str(Loc.X) & ")"
  
  'Get the location of the mouse-click in control units
  Else
    MsgBox "Control units: " & Str$(X) & "," & Str$(Y)
  End If
  
 ElseIf Toolbar1.Buttons(4).Value = 1 Then             ' add an event
     Dim pt As mapobjects2.point

    If Button = 1 Then 'Poner Punto
        Set pt = Map1.ToMapPoint(X, Y)
    With Map1.TrackingLayer.Symbol(0)
      .SymbolType = moPointSymbol
      .Size = 6
      .Style = moCircleMarker
    End With
    Map1.TrackingLayer.AddEvent pt, 0
    Else 'Borrar Punto
       If Map1.TrackingLayer.EventCount > 0 Then
    Map1.TrackingLayer.RemoveEvent 0
    
    End If
   End If
 ElseIf Toolbar1.Buttons(5).Value = 1 Then
 Dim r As mapobjects2.Rectangle
  Dim nEventCount As Long, i As Long
  
    Set r = Map1.TrackRectangle
    nEventCount = Map1.TrackingLayer.EventCount
    Dim testPt As New mapobjects2.point
    Dim evt As mapobjects2.GeoEvent
    Dim j
    
    For i = 0 To nEventCount - 1

      Set evt = Map1.TrackingLayer.Event(i)
      testPt.X = evt.X
      testPt.Y = evt.Y
   
If r.IsPointIn(testPt) Then
       For j = 0 To nEventCount - 2
            If (evt.Tag = List2.List(j)) Then
                posicionlista = j
                frminfo.Show
                
                
            End If
        Next
    End If
             


    Next
    
     '   Label2.Caption = evt.Tag
    Map1.Refresh
End If

    
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.index = 1 Then
  Map1.MousePointer = moZoom
  
ElseIf Button.index = 2 Then
  Map1.MousePointer = moPan
ElseIf Button.index = 3 Then
  Map1.MousePointer = moCross
  ElseIf Button.index = 4 Then
  Map1.MousePointer = moCross
   ElseIf Button.index = 5 Then
  Map1.MousePointer = moCross
ElseIf Button.index = 7 Then
  Map1.MousePointer = moArrow
  Set Map1.Extent = Map1.FullExtent

ElseIf Button.index = 6 Then
    Dim FName As String

  Map1.MousePointer = moArrow
  
  
  On Error GoTo ControlError

 

  'Set up dailog box to prompt user to load a imagefile
  CommonDialog2.Filter = "Imágenes (*.jpg)|*.jpg"
  'Set cancel error so that if cancel is used then we can trap it an exit
  CommonDialog2.CancelError = True
  CommonDialog2.ShowOpen
  If Len(CommonDialog2.FileName) = 0 Then Exit Sub

   
  
   
      FName = CommonDialog2.FileName
      
      Map1.ExportMapToJpeg FName  '"C:\PROYECTO\Imágenes\sucre\prueba.jpg"

      
      Exit Sub   ' Salir para evitar el controlador.
ControlError:   ' Rutina de control de errores.
   Select Case Err.Number   ' Evalúa el número de error.
      Case 424   ' Error "Archivo ya está abierto".
         MsgBox "No existen capas para eliminar"   ' Cierra el archivo abierto.
      Case Else
      ' Puede incluir aquí otras situaciones...
            Exit Sub
   End Select
   
   
End If
End Sub
