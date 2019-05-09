VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_buscar_archivo 
   Caption         =   "Procesar archivo SEISAN"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5985
   Begin VB.TextBox txt_events 
      BackColor       =   &H00E0E0E0&
      DataField       =   "MAX(idevent)"
      DataSource      =   "Adod_events"
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adod_events 
      Height          =   330
      Left            =   3600
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=ODBCSIISS"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ODBCSIISS"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select MAX(idevent) from events"
      Caption         =   "Events"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txt_nombre 
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmd_cerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txt_buscar 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   5295
   End
   Begin VB.CommandButton cmd_abrir 
      Caption         =   "&Abrir"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CmD 
      Left            =   5400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_procesar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Procesar..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adod_magnitud 
      Height          =   330
      Left            =   2040
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=ODBCSIISS"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ODBCSIISS"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select MAX(idevent) from events"
      Caption         =   "Magnitud"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lbl_proceso 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número del Ultimo Evento Procesado:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   360
      Width           =   2700
   End
   Begin VB.Label LblRuta 
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   390
   End
End
Attribute VB_Name = "frm_buscar_archivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Desarrollado por Lic. Francisco Alvarez
'Fecha: 08/05/2006

Dim Matlab As Object
Dim Result As String

Private Sub cmd_abrir_Click()

    ' Establecer CancelError a True
    CmD.CancelError = True
    On Error GoTo ErrHandler
    ProgressBar.Value = 0
    ' Establecer los indicadores
    CmD.Flags = cdlOFNHideReadOnly
    ' Establecer los filtros
    CmD.Filter = "Todos los archivos (*.*)|*.*|Archivos de texto" & _
    "(*.txt)|*.txt|Archivos por lotes (*.bat)|*.bat|Archivos SEISAN (*.out)|*.out"
    ' Especificar el filtro predeterminado
    CmD.FilterIndex = 4
    ' Presentar el cuadro de diálogo Abrir
    CmD.ShowOpen
    ' Presentar el nombre del archivo seleccionado
    Me.txt_buscar.Text = CmD.FileName
    Me.txt_nombre.Text = CmD.FileTitle

    'Procesar ultimo idevents
    'txt_events.Text = Adod_events.Recordset.Fields(1).Value
    
    cmd_procesar.Enabled = True
    
    Exit Sub
    
ErrHandler:
    ProgressBar.Value = 0
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub

End Sub

Private Sub cmd_abrir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd_cerrar.FontBold = False
cmd_abrir.FontBold = True
cmd_procesar.FontBold = False
'cmd_reporte.FontBold = False
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

'Private Sub Command1_Click()
'
'Dim a
''ubica la ruta en que tenga un programa puede ponerse la que se necesite
'Result = Matlab.Execute("cd C:\HOME\FAOVE\siiss\MatLab\Programas\")
''llamo mi programa y le envio una variable
''Result = Matlab.Execute("a=plot(" & Str(Val(Text1.Text)) & ")")
'Result = Matlab.Execute("seisan_lectura_out")
''tomo la variable que me devuelve el programa
''Call Matlab.GetWorkspaceData("a", "base", a)
''Label1.Caption = a
''Text2.Text = 2 * a
'
''si necesitas ejecutar las funciones de Matlab como se ve
''Result = MatLab.Execute("pause(1)")
''Result = MatLab.Execute("bode(tf([1],[1 1]))")
'
'
'End Sub

'Private Sub View_Methods()
'Dim Matlab2 As MLApp.MLApp
'
'
'Set Matlab2 = New MLApp.MLApp
'
''Para ver los métodos MATLAB Automation methods
'Matlab2.Feval
'
'
'End Sub


Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    
    TxtRuta.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo ControlError
    Dir1.Path = Drive1.Drive

Exit Sub   ' Salir para evitar el controlador.
ControlError:   ' Rutina de control de errores.
   Select Case Err.Number   ' Evalúa el número de error.
      Case 55   ' Error "Archivo ya está abierto".
         Close #1   ' Cierra el archivo abierto.
      Case Else
      ' Puede incluir aquí otras situaciones...
      ProgressBar.Value = 0
            Exit Sub
   End Select
      ' Continuar ejecución en la línea que
            ' causó el error.
    
End Sub

Private Sub File1_Click()
TxtRuta.Text = Dir1.Path & "\" & File1.FileName
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd_cerrar.FontBold = True
cmd_abrir.FontBold = False
cmd_procesar.FontBold = False
'cmd_reporte.FontBold = False
End Sub

Private Sub cmd_procesar_Click()
On Error GoTo ControlError
Dim fso As New FileSystemObject, fil As File
Dim ts As TextStream
Dim s, mag, loct, comput, sqlstr_comput, sqlstr_events, sqlstr_loct, sqlstr_mag, sqlstr_read

Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
lbl_proceso.Caption = "Conectándose a MatLab..."
ProgressBar.Value = 15

Set Matlab = CreateObject("Matlab.Application")
'al hacer matlab server invisible es más rapido pero se puede dejar visible
Matlab.Visible = 0
ProgressBar.Value = 45
Dim a
'ubica la ruta en que tenga un programa puede ponerse la que se necesite
Result = Matlab.Execute("cd C:\PROYECTO\MatLab\")
'llamo mi programa y le envio una variable

Result = Matlab.Execute("seisan_lectura_out('" & txt_buscar.Text & "'," & txt_events.Text & ")")
'Result = Matlab.Execute("e(" & txt_nombre.Text & ")")

ProgressBar.Value = 55

lbl_proceso.Caption = "Conectándose con MySQL..."

Call actualizar_conex

Set fso = CreateObject("Scripting.FileSystemObject")

'-----------------------------------------------
'Procedimiento para insertar en la tabla eventos
'-----------------------------------------------
Set ts = fso.OpenTextFile("C:\PROYECTO\MatLab\events.m", ForReading)

While ts.AtEndOfStream = False

    s = ts.ReadLine

    sqlstr_events = Trim(s)

    If Not (sqlstr_events = "") Then

        cn.Execute sqlstr_events

    End If

Wend

ts.Close
'--------------------------------------------------
'Procedimiento para insertar en la tabla magnitudes
'--------------------------------------------------
Set ts = fso.OpenTextFile("C:\PROYECTO\MatLab\magnitudes.m", ForReading)

While ts.AtEndOfStream = False

    
    mag = ts.ReadLine

    sqlstr_mag = Trim(mag)

    If Not (sqlstr_mag = "") Then

        cn.Execute sqlstr_mag

    End If

Wend

ts.Close
'-------------------------------------------------
'Procedimiento para insertar en la tabla locations
'-------------------------------------------------
Set ts = fso.OpenTextFile("C:\PROYECTO\MatLab\locations.m", ForReading)

While ts.AtEndOfStream = False

    loct = ts.ReadLine

    sqlstr_loct = Trim(loct)

    If Not (sqlstr_loct = "") Then

        cn.Execute sqlstr_loct

    End If

Wend
ts.Close
'------------------------------------------------
'Procedimiento para insertar en la tabla readings
'------------------------------------------------
Set ts = fso.OpenTextFile("C:\PROYECTO\MatLab\readings.m", ForReading)

While ts.AtEndOfStream = False

    readt = ts.ReadLine
    
    'MsgBox readt
  
    sqlstr_readt = Trim(readt)
    
    'If Not IsEmpty(sqlstr_readt) Then
    If Not (sqlstr_readt = "") Then
    
        cn.Execute sqlstr_readt
        
    End If
    
Wend
'Cierra el archivo
ts.Close

'-------------------------------------------------------
'Procedimiento para insertar en la tabla computed_values
'-------------------------------------------------------
Set ts = fso.OpenTextFile("C:\PROYECTO\MatLab\computevalues.m", ForReading)

While ts.AtEndOfStream = False

    comput = ts.ReadLine
    
    sqlstr_comput = Trim(comput)
    
    'If Not IsEmpty(sqlstr_readt) Then
    If Not (sqlstr_comput = "") Then
    
        cn.Execute sqlstr_comput
        
    End If
    
Wend
'Cierra el archivo
ts.Close
'Cierra la conexion con la base de datos
cn.Close

lbl_proceso.Caption = "Listo..."

ProgressBar.Value = 100

cmd_procesar.Enabled = False

'cmd_reporte.Enabled = True

Exit Sub   ' Salir para evitar el controlador.
ControlError:   ' Rutina de control de errores.
   Select Case Err.Number   ' Evalúa el número de error.
      Case -2147467259   ' Error "Archivo ya está abierto".
         MsgBox "Evento duplicado no se pudo insertar"   ' Cierra el archivo abierto.
         lbl_proceso.Caption = "Error, evento duplicado..."
         'cmd_reporte.Enabled = False
      Case Else
        ' Puede incluir aquí otras situaciones...
        ProgressBar.Value = 0
        'cmd_reporte.Enabled = False
        Exit Sub
   End Select
End Sub

Private Sub cmd_procesar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd_cerrar.FontBold = False
cmd_abrir.FontBold = False
'cmd_reporte.FontBold = False
cmd_procesar.FontBold = True
End Sub

Private Sub Form_Load()
Me.Height = 3645
Me.Width = 6105
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd_cerrar.FontBold = False
cmd_abrir.FontBold = False
cmd_procesar.FontBold = False
'cmd_reporte.FontBold = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ControlError
Result = Matlab.Execute("close all")
Result = Matlab.Execute("clear all")
'para bajar de memoria al PC y liberar matlab
Matlab.Quit
Exit Sub   ' Salir para evitar el controlador.
ControlError:   ' Rutina de control de errores.
   Select Case Err.Number   ' Evalúa el número de error.
      Case 55   ' Error "Archivo ya está abierto".
         Close #1   ' Cierra el archivo abierto.
      Case Else
      ' Puede incluir aquí otras situaciones...
            ProgressBar.Value = 0
            Exit Sub
   End Select
      ' Continuar ejecución en la línea que
            ' causó el error.
End Sub

Private Sub TxtRuta_Change()
    If Right(TxtRuta, 3) = "txt" Then
        CmdAbrir.Enabled = True
    Else
        CmdAbrir.Enabled = False
    End If
End Sub


