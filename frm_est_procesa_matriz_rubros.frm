VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_est_procesa_matriz_rubros 
   Caption         =   "PROCESAR MATRÍZ DE RUBROS"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   9450
   Begin VB.TextBox Text4 
      DataField       =   "Id_Obj"
      DataSource      =   "Tab_Rubros_Incidencia"
      Height          =   285
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text3 
      DataField       =   "CONCEPTO"
      DataSource      =   "TAB_TRADUCE_RUBROS"
      Height          =   285
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text2 
      DataField       =   "Renglon"
      DataSource      =   "FORMA_DE_PAGO"
      Height          =   285
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      DataField       =   "ID_INSTANCIA"
      DataSource      =   "CUM_FAC"
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSAdodcLib.Adodc CUM_FAC 
      Height          =   375
      Left            =   840
      Top             =   5400
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      Connect         =   "DSN=SIAGEP"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "SIAGEP"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM CUM_FAC WHERE ID_INSTANCIA= '-'"
      Caption         =   "CUM_FAC"
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3735
      Left            =   840
      TabIndex        =   10
      Top             =   1560
      Width           =   8295
      Begin VB.Frame Frame_rango_fechas 
         Caption         =   "Seleccione Rango de Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   6015
         Begin VB.ListBox List_archivo 
            Height          =   450
            ItemData        =   "frm_est_procesa_matriz_rubros.frx":0000
            Left            =   2880
            List            =   "frm_est_procesa_matriz_rubros.frx":000A
            TabIndex        =   3
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox txt_año 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            MaxLength       =   4
            TabIndex        =   2
            Top             =   1440
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker txt_desde 
            Height          =   375
            Left            =   600
            TabIndex        =   0
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51249155
            CurrentDate     =   38028
         End
         Begin MSComCtl2.DTPicker txt_hasta 
            Height          =   375
            Left            =   2880
            TabIndex        =   1
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51249155
            CurrentDate     =   38028
         End
         Begin VB.Label lbl_archivo 
            Caption         =   "Archivo a Procesar:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   15
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label lbl_año 
            Caption         =   "Año:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   14
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label lbl_fecha_hasta 
            Caption         =   "Fecha hasta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   13
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Lbl_fecha_desde 
            Caption         =   "Fecha desde:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   12
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   5280
         TabIndex        =   6
         Tag             =   "Cerrar matriz de rubros"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton cmd_upt_tabla 
         Caption         =   "Act. Tabla de  Incidencia"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3720
         TabIndex        =   5
         Tag             =   "Cerrar matriz de rubros"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton cmd_procesa 
         Caption         =   "&Procesar Rubros"
         Height          =   615
         Left            =   2160
         TabIndex        =   4
         Tag             =   "Cerrar matriz de rubros"
         Top             =   2640
         Width           =   1575
      End
      Begin MSComctlLib.ProgressBar PBar_matriz 
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   3360
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   960
      TabIndex        =   7
      Top             =   240
      Width           =   8295
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "DE RUBROS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   "PROCESAR MATRÍZ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   0
         Width           =   7815
      End
   End
   Begin MSAdodcLib.Adodc FORMA_DE_PAGO 
      Height          =   375
      Left            =   3960
      Top             =   5400
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      Connect         =   "DSN=SIAGEP"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "SIAGEP"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM  FORMA_DE_PAGO WHERE ID_INSTANCIA = '-'"
      Caption         =   "FORMA_DE_PAGO"
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
   Begin MSAdodcLib.Adodc TAB_TRADUCE_RUBROS 
      Height          =   375
      Left            =   6960
      Top             =   5400
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "DSN=SIAGEP"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "SIAGEP"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TAB_TRADUCE_RUBROS"
      Caption         =   "TAB_TRADUCE_RUBROS"
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
   Begin MSAdodcLib.Adodc Tab_Rubros_Incidencia 
      Height          =   375
      Left            =   6360
      Top             =   1080
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      Connect         =   "DSN=SIAGEP"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "SIAGEP"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from TAB_RUBROS_INCIDENCIA where id_obj = ''"
      Caption         =   "Tab_Rubros_Incidencia"
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
End
Attribute VB_Name = "frm_est_procesa_matriz_rubros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TAB_TRA(100, 2)

Dim ING_X_RUBROS(40, 26)

Dim MESDES As Byte, MESHAS As Byte

Dim mvBookMark

Option Explicit

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = True
    Me.cmd_procesa.FontBold = False
    Me.cmd_upt_tabla.FontBold = False
End Sub

Private Sub cmd_procesa_Click()
On Error GoTo control_de_errores
Dim C As Integer
Dim i As Byte, J As Byte
Dim MES As Byte
Rem Me.TEXT_MONTO = Me.TEXT_MONTO + CDbl(Format(Me.Monto, "###########0.00"))

Dim SW_FOUND As Boolean
Dim SW_FOUND_RUBRO As Boolean

i = 0

Dim sqlstr As String
Dim fecha_desde As String
Dim fecha_hasta As String
Dim SW_F_D_P As Boolean
Dim Rubro As String
Dim FECHA_PAGO As Date

If List_archivo.Text = "" Then
    MsgBox "Por favor, seleccione el archivo a procesar", vbInformation, "ALCALSIS"
    List_archivo.SetFocus
    Exit Sub
End If
PBar_matriz.Visible = True
PBar_matriz.Min = 0

fecha_desde = STR(Month(Me.txt_desde.Value)) + STR(Day(txt_desde.Value)) + STR(Year(txt_desde.Value))
fecha_hasta = STR(Month(txt_hasta.Value)) + STR(Day(txt_hasta.Value)) + STR(Year(txt_hasta.Value))

MESDES = Trim(Month(txt_desde.Value))
MESHAS = Trim(Month(txt_hasta.Value))

Init_Tablas

Select Case Me.List_archivo

    Case "FORMA_DE_PAGO"
    
        sqlstr = "SELECT * FROM FORMA_DE_PAGO "
'        sqlstr = sqlstr + " WHERE (FEC_PAGO>= CONVERT(DATETIME, '" + Format(fecha_desde, "dd/mm/yyyy") + "', 102))"
'        sqlstr = sqlstr + " AND (FEC_PAGO<= CONVERT(DATETIME, '" + Format(Fecha_hasta, "dd/mm/yyyy") + "', 102))"
        sqlstr = sqlstr + " WHERE FEC_PAGO >= '" & CDate(fecha_desde) & "'"
        sqlstr = sqlstr + " AND FEC_PAGO <= '" & CDate(fecha_hasta) & "'"
        sqlstr = sqlstr + " AND (STATUS = 'CA' OR STATUS IS NULL)"
        sqlstr = sqlstr + " ORDER BY FEC_PAGO, ID_RUBRO, NRO_PLANI_PAGO "

        SW_F_D_P = True
        FORMA_DE_PAGO.ConnectionString = "SIAGEP"
        FORMA_DE_PAGO.CommandType = adCmdText

        FORMA_DE_PAGO.RecordSource = sqlstr

        FORMA_DE_PAGO.Refresh

        If FORMA_DE_PAGO.Recordset.EOF Then

            MsgBox "CONJUNTO DE DATOS A PROCESAR VACIO PARA RANGO DE FECHA DADO:" + fecha_desde + " --> " + fecha_hasta
    
            Exit Sub

        End If
        PBar_matriz.Min = 0
        PBar_matriz.Max = FORMA_DE_PAGO.Recordset.RecordCount
        
        Do While FORMA_DE_PAGO.Recordset.EOF = False
            
            SW_FOUND_RUBRO = False
            
            SW_FOUND = False
        
            C = C + 1
            
            PBar_matriz.Value = C
            
            If SW_F_D_P Then
            
                Rubro = FORMA_DE_PAGO.Recordset!Id_Rubro
                FECHA_PAGO = FORMA_DE_PAGO.Recordset!Fec_pago
            
            Else
            
                Rubro = FORMA_DE_PAGO.Recordset!Concepto
                FECHA_PAGO = FORMA_DE_PAGO.Recordset!FEC_CANCEL
                
            End If
            
            For i = 1 To 100
            
            If Rubro = TAB_TRA(i, 1) Then
                
                SW_FOUND_RUBRO = True
                
                For J = 1 To 40
                
                    If TAB_TRA(i, 2) = ING_X_RUBROS(J, 1) Then
                    
                        MES = Month(FECHA_PAGO)
                        
                        MES = (MES * 2) + 1
                        
                        ING_X_RUBROS(J, MES) = ING_X_RUBROS(J, MES) + FORMA_DE_PAGO.Recordset!monto
                        
                        ING_X_RUBROS(J, MES + 1) = ING_X_RUBROS(J, MES + 1) + 1
                        
                        SW_FOUND = True
                                            
                        Exit For
                        
                    End If
               
                Next
                    
                If SW_FOUND Then
                    
                    Exit For
                        
                Else
                
                    MsgBox "CODIGO DE TRADUCCION NO EXISTE EN MATRIZ DE RUBROS:" + Rubro
                
                    Exit For
                                 
                End If
                    
            
            End If
         
                 
        Next
         
            If SW_FOUND_RUBRO = False Then
                 
                 MsgBox "Tome Nota: Rubro leido desde   F_D_P No Tiene Entrada de Traducció:" + Rubro
                 MsgBox "Proceso de lectura Continua en F_D_P."
                 
                                 
            End If
            
        '   Forms![Reporte_de_Ejecución]![Tex_Records] = C
            
            FORMA_DE_PAGO.Recordset.MoveNext
         
        Loop

    
    Case "CUM_FAC"
    
        sqlstr = "SELECT * FROM CUM_FAC "
        sqlstr = sqlstr + " WHERE FEC_CANCEL>='" + Format(fecha_desde, "dd/mm/yyyy") + "'"
        sqlstr = sqlstr + " AND   FEC_CANCEL<='" + Format(fecha_hasta, "dd/mm/yyyy") + "'"
        sqlstr = sqlstr + " AND ( ( STATUS) = 'CA' OR (STATUS) IS NULL)"
        sqlstr = sqlstr + " ORDER BY FEC_CANCEL, CONCEPTO, NRO_PLANI_PAGO" + ";"

        SW_F_D_P = False
        CUM_FAC.CommandType = adCmdText

        CUM_FAC.RecordSource = sqlstr

        CUM_FAC.Refresh
        
        If CUM_FAC.Recordset.EOF Then

            MsgBox "CONJUNTO DE DATOS A PROCESAR VACIO PARA RANGO DE FECHA DADO:" + fecha_desde + " --> " + fecha_hasta
    
            Exit Sub

        End If
        
        PBar_matriz.Min = 0
        PBar_matriz.Max = CUM_FAC.Recordset.RecordCount
        
        Do While CUM_FAC.Recordset.EOF = False
            
            SW_FOUND_RUBRO = False
            
            SW_FOUND = False
        
            C = C + 1
            
            PBar_matriz.Value = C
            
            If SW_F_D_P Then
            
                Rubro = CUM_FAC.Recordset!Id_Rubro
                FECHA_PAGO = CUM_FAC.Recordset!Fec_pago
            
            Else
            
                Rubro = CUM_FAC.Recordset!Concepto
                FECHA_PAGO = CUM_FAC.Recordset!FEC_CANCEL
                
            End If
            
            For i = 1 To 100
            
            If Rubro = TAB_TRA(i, 1) Then
                
                SW_FOUND_RUBRO = True
                
                For J = 1 To 40
                
                    If TAB_TRA(i, 2) = ING_X_RUBROS(J, 1) Then
                    
                        MES = Month(FECHA_PAGO)
                        
                        MES = (MES * 2) + 1
                        
                        ING_X_RUBROS(J, MES) = ING_X_RUBROS(J, MES) + CUM_FAC.Recordset!monto
                        
                        ING_X_RUBROS(J, MES + 1) = ING_X_RUBROS(J, MES + 1) + 1
                        
                        SW_FOUND = True
                                            
                        Exit For
                        
                    End If
               
                Next
                    
                If SW_FOUND Then
                    
                    Exit For
                        
                Else
                
                    MsgBox "CODIGO DE TRADUCCION NO EXISTE EN MATRIZ DE RUBROS:" + Rubro
                
                    Exit For
                                 
                End If
                    
            
            End If
         
                 
        Next
         
            If SW_FOUND_RUBRO = False Then
                 
                 MsgBox "Tome Nota: Rubro leido desde   F_D_P No Tiene Entrada de Traducció:" + Rubro
                 MsgBox "Proceso de lectura Continua en F_D_P."
                 
                                 
            End If
            
            CUM_FAC.Recordset.MoveNext
         
        Loop
        
        
    Case Else

        MsgBox "NOMBRE DE ARCHIVO A PROCESAR INVALIDO.VERIFIQUE.GRACIAS"
    
        Exit Sub
    

End Select


MsgBox "Records Seleccionados:" + STR(C)

Me.cmd_upt_tabla.Enabled = True
PBar_matriz.Visible = False

Exit Sub
control_de_errores:

    MsgBox " " & Err.Number & " :  " & Err.Description & "  "
    PBar_matriz.Visible = False

End Sub

Private Sub cmd_procesa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_procesa.FontBold = True
    Me.cmd_upt_tabla.FontBold = False
End Sub

Private Sub cmd_upt_tabla_Click()
Screen.MousePointer = 11
On Error GoTo control_de_errores
Dim i As Byte, J As Byte
Dim MES As Byte

Rem Me.TEXT_MONTO = Me.TEXT_MONTO + CDbl(Format(Me.Monto, "###########0.00"))

Dim rds As ADODB.Recordset
Dim sqlstr As String
Dim NAÑO As String

Set rds = New ADODB.Recordset
i = 0

Rem DoCmd.OpenQuery "UPD_TAB_RUBROS_INCIDENCIA"

'sqlstr = "" revisar
'
'sqlstr = "SELECT *  FROM Tab_Rubros_Incidencia  " revisar
'sqlstr = sqlstr + " WHERE AÑO=" + "'" + Trim((Me.AÑO)) + "'" + ";" revisar

rds.Open sqlstr, cn

If rds.EOF = True Then
    rds.Close
'    NAÑO = STR(Val(Me.AÑO) - 1)   revisar
    
    sqlstr = "SELECT *  FROM Tab_Rubros_Incidencia  "
    sqlstr = sqlstr + " WHERE AÑO=" + "'" + Trim(NAÑO) + "'" + ";"

    rds.Open sqlstr, cn
    If rds.EOF = True Then
    
        MsgBox "No Existe Referencia de Registros de Inicidencia para el Año de Proceso."
        Screen.MousePointer = 0
        Exit Sub
        
    End If
    
    Dim ARUBRO As String, ADES As String, AAÑO As String
    
'    AAÑO = Trim(Me.AÑO)  revisar
    
    sqlstr = "INSERT INTO Tab_Rubros_Incidencia  (AÑO,COD_RUBRO,DESCRIPCION) VALUES  ("
    
    Do While rds.EOF = False
    
        ARUBRO = rds!Cod_Rubro
        ADES = rds!Descripcion
        
        sqlstr = sqlstr + "'" + (AAÑO) + "'," + "'" + (ARUBRO) + "'," + "'" + (ADES) + "')"
        
        cn.Execute sqlstr
       
        rds.MoveNext
    
    Loop

End If
rds.Close

For i = 1 To 40
 
   If ING_X_RUBROS(i, 1) = Null Or ING_X_RUBROS(i, 1) = 0 Then
        
        Exit For
        
    End If
    
    For J = 3 To 25 Step 2
        
        If ING_X_RUBROS(i, J) > 0 Then
                    
'            sqlstr = "" revisar
'
'            sqlstr = "SELECT *  FROM Tab_Rubros_Incidencia  "
'            sqlstr = sqlstr + " WHERE AÑO=" + "'" + Trim((Me.AÑO)) + "'"
'            sqlstr = sqlstr + " AND   COD_RUBRO=" + "'" + (ING_X_RUBROS(I, 1)) + "';"
'
'            rds.Open sqlstr, cn, adOpenKeyset, adLockPessimistic
'
            If rds.EOF = True Then

                   MsgBox "No Existe Entrada En la Tabla RubroS Incidencia:" + ING_X_RUBROS(i, 1)
                   
            Else
     
                Select Case J
            
                    Case Is = 3
                            
                            rds!Ene_Mon = NZ(rds!Ene_Mon, 0) + ING_X_RUBROS(i, J)
                            rds!Ene_Can = NZ(rds!Ene_Can, 0) + ING_X_RUBROS(i, J + 1)
                            rds!Real = NZ(rds!Real, 0) + ING_X_RUBROS(i, J)
                            
                       rds.Update
                            
                    Case Is = 5
                        
                            rds!Feb_Mon = NZ(rds!Feb_Mon, 0) + ING_X_RUBROS(i, J)
                            rds!Feb_Can = NZ(rds!Feb_Can, 0) + ING_X_RUBROS(i, J + 1)
                            rds!Real = NZ(rds!Real, 0) + ING_X_RUBROS(i, J)
                       
                       rds.Update
                            
                     Case Is = 7
                        
                            rds!Mar_Mon = NZ(rds!Mar_Mon, 0) + ING_X_RUBROS(i, J)
                            rds!Mar_Can = NZ(rds!Mar_Can, 0) + ING_X_RUBROS(i, J + 1)
                            rds!Real = NZ(rds!Real, 0) + ING_X_RUBROS(i, J)
                       
                       rds.Update
                                   
                    Case Is = 9
                        
                            rds!Abr_Mon = NZ(rds!Abr_Mon, 0) + ING_X_RUBROS(i, J)
                            rds!Abr_Can = NZ(rds!Abr_Can, 0) + ING_X_RUBROS(i, J + 1)
                            rds!Real = NZ(rds!Real, 0) + ING_X_RUBROS(i, J)
                       
                       rds.Update
                            
                    Case Is = 11
                    
                            rds!May_Mon = NZ(rds!May_Mon, 0) + ING_X_RUBROS(i, J)
                            rds!May_Can = NZ(rds!May_Can, 0) + ING_X_RUBROS(i, J + 1)
                            rds!Real = NZ(rds!Real, 0) + ING_X_RUBROS(i, J)
                       
                       rds.Update
                    
                    Case Is = 13
                    
                            rds!Jun_Mon = NZ(rds!Jun_Mon, 0) + ING_X_RUBROS(i, J)
                            rds!Jun_Can = NZ(rds!Jun_Can, 0) + ING_X_RUBROS(i, J + 1)
                            rds!Real = NZ(rds!Real, 0) + ING_X_RUBROS(i, J)
                       
                       rds.Update
                    
                    Case Is = 15
                    
                            rds!Jul_Mon = NZ(rds!Jul_Mon, 0) + ING_X_RUBROS(i, J)
                            rds!Jul_Can = NZ(rds!Jul_Can, 0) + ING_X_RUBROS(i, J + 1)
                            rds!Real = NZ(rds!Real, 0) + ING_X_RUBROS(i, J)
                       
                       rds.Update
                    
                    Case Is = 17
                       
                            rds!Ago_Mon = NZ(rds!Ago_Mon, 0) + ING_X_RUBROS(i, J)
                            rds!Ago_Can = NZ(rds!Ago_Can, 0) + ING_X_RUBROS(i, J + 1)
                            rds!Real = NZ(rds!Real, 0) + ING_X_RUBROS(i, J)
                       
                       rds.Update
                    
                    Case Is = 19
                    
                            rds!Sep_Mon = NZ(rds!Sep_Mon, 0) + ING_X_RUBROS(i, J)
                            rds!Sep_Can = NZ(rds!Sep_Can, 0) + ING_X_RUBROS(i, J + 1)
                            rds!Real = NZ(rds!Real, 0) + ING_X_RUBROS(i, J)
                       
                       rds.Update
                    
                    Case Is = 21
                    
                            rds!Oct_Mon = NZ(rds!Oct_Mon, 0) + ING_X_RUBROS(i, J)
                            rds!Oct_Can = NZ(rds!Oct_Can, 0) + ING_X_RUBROS(i, J + 1)
                            rds!Real = NZ(rds!Real, 0) + ING_X_RUBROS(i, J)
                       
                       rds.Update
                    
                    Case Is = 23
                        
                            rds!Nov_Mon = NZ(rds!Nov_Mon, 0) + ING_X_RUBROS(i, J)
                            rds!Nov_Can = NZ(rds!Nov_Can, 0) + ING_X_RUBROS(i, J + 1)
                            rds!Real = NZ(rds!Real, 0) + ING_X_RUBROS(i, J)
                            
                       
                       rds.Update
                       
                     Case Is = 25
                    
                            rds!Dic_Mon = NZ(rds!Dic_Mon, 0) + ING_X_RUBROS(i, J)
                            rds!Dic_Can = NZ(rds!Dic_Can, 0) + ING_X_RUBROS(i, J + 1)
                            rds!Real = NZ(rds!Real, 0) + ING_X_RUBROS(i, J)
                       
                       rds.Update
                       
                End Select
     
            End If
        
          rds.Close
        
        End If
    
    Next
    
Next

    MsgBox "Fin Generación de Tabla de Incidencia Ingresos Rubros."
    Screen.MousePointer = 0
Exit Sub
control_de_errores:
    MsgBox " " & Err.Number & " :  " & Err.Description & "  "
    Screen.MousePointer = 0
End Sub

Private Sub cmd_upt_tabla_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_procesa.FontBold = False
    Me.cmd_upt_tabla.FontBold = True
End Sub

Private Sub Form_Load()
txt_desde.Value = Format(Date, "dd/mm/yyyy")
txt_hasta.Value = Format(Date, "dd/mm/yyyy")
txt_año.Text = Year(Date)
End Sub

Private Sub Form_Resize()
    Call Mover_der(Me, Frame2, 0)
    Call Mover_centrado(Me, Frame1)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_procesa.FontBold = False
    Me.cmd_upt_tabla.FontBold = False
End Sub

Private Sub Txt_año_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub
Private Sub Init_Tablas()
On Error GoTo control_de_errores

Dim sqlstr As String
Dim i As Byte, J As Byte

ING_X_RUBROS(1, 1) = "301020500"
ING_X_RUBROS(1, 2) = "Impuesto Inmuebles Urbanos"
PBar_matriz.Min = 0
ING_X_RUBROS(2, 1) = "301040500"
ING_X_RUBROS(2, 2) = "Patente de Industria y Comercio"

ING_X_RUBROS(3, 1) = "301020800"
ING_X_RUBROS(3, 2) = "Patente de Vehiculo"

ING_X_RUBROS(4, 1) = "301040700"
ING_X_RUBROS(4, 2) = "Publicidad Comercial"

ING_X_RUBROS(5, 1) = "301040800"
ING_X_RUBROS(5, 2) = "Espectáculos Públicos"

ING_X_RUBROS(6, 1) = "301040900"
ING_X_RUBROS(6, 2) = "Apuestas Lícitas"

ING_X_RUBROS(7, 1) = "301041000"
ING_X_RUBROS(7, 2) = "Deuda Morosa"

ING_X_RUBROS(8, 1) = "301100100"
ING_X_RUBROS(8, 2) = "Tasas"

ING_X_RUBROS(9, 1) = "301120200"
ING_X_RUBROS(9, 2) = "Reparos Fiscales"

ING_X_RUBROS(10, 1) = "301120800"
ING_X_RUBROS(10, 2) = "Multas y Recargos"

ING_X_RUBROS(11, 1) = "301160000"
ING_X_RUBROS(11, 2) = "Ingresos Por Aporte al Municipio"

ING_X_RUBROS(12, 1) = "301120301"
ING_X_RUBROS(12, 2) = "Reintegros de Particulares"

ING_X_RUBROS(13, 1) = "301120302"
ING_X_RUBROS(13, 2) = "Reintegros Fondos Girados Avance"

ING_X_RUBROS(14, 1) = "301120499"
ING_X_RUBROS(14, 2) = "Ingresos Varios"

ING_X_RUBROS(15, 1) = "301120500"
ING_X_RUBROS(15, 2) = "Recargos e Intereses Moratorios"

ING_X_RUBROS(16, 1) = "301120705"
ING_X_RUBROS(16, 2) = "Indemnizacion P/Incumplimiento D/Contrato"

ING_X_RUBROS(17, 1) = "301130102"
ING_X_RUBROS(17, 2) = "Ventas de Gacetas Mcpales y Formularios"

ING_X_RUBROS(18, 1) = "301130108"
ING_X_RUBROS(18, 2) = "Operaciones Diversas Municipios"

ING_X_RUBROS(19, 1) = "301140200"
ING_X_RUBROS(19, 2) = "Intereses Por Depositos"

ING_X_RUBROS(20, 1) = "301140300"
ING_X_RUBROS(20, 2) = "Intereses Por Titulos y Valores"

ING_X_RUBROS(21, 1) = "301140401"
ING_X_RUBROS(21, 2) = "Rentas Inmobiliarias"

ING_X_RUBROS(22, 1) = "301140501"
ING_X_RUBROS(22, 2) = "Arrendamientos de Ejidos"

ING_X_RUBROS(23, 1) = "302020300"
ING_X_RUBROS(23, 2) = "Reservas D/Tesoro Mcpal. No Comprometido"

ING_X_RUBROS(24, 1) = "302060200"
ING_X_RUBROS(24, 2) = "Ventas de Bienes"

ING_X_RUBROS(25, 1) = "301100150"
ING_X_RUBROS(25, 2) = "Aseo Urbano Domiciliario"

ING_X_RUBROS(26, 1) = "301100100"
ING_X_RUBROS(26, 2) = "Otras Tasas"

ING_X_RUBROS(27, 1) = "301160300"
ING_X_RUBROS(27, 2) = "Aportes"

ING_X_RUBROS(28, 1) = "301120300"
ING_X_RUBROS(28, 2) = "Reintegros"

ING_X_RUBROS(29, 1) = "301160301"
ING_X_RUBROS(29, 2) = "Aporte Especial Gobernación"

ING_X_RUBROS(30, 1) = "301160302"
ING_X_RUBROS(30, 2) = "Aporte Especial Gobierno Nacional"

ING_X_RUBROS(31, 1) = "301160100"
ING_X_RUBROS(31, 2) = "Situado Municipal"

ING_X_RUBROS(32, 1) = "301160304"
ING_X_RUBROS(32, 2) = "Aporte FIDES"

For i = 1 To 40

    For J = 3 To 26
        
        ING_X_RUBROS(i, J) = 0
    
    Next

Next

i = 0
'Set RDS = New ADODB.Recordset
'actualizar_cn
'RDS.Open "TAB_TRADUCE_RUBROS", cn
PBar_matriz.Max = TAB_TRADUCE_RUBROS.Recordset.RecordCount
Do While TAB_TRADUCE_RUBROS.Recordset.EOF = False

    i = i + 1
    PBar_matriz.Value = i
    TAB_TRA(i, 1) = TAB_TRADUCE_RUBROS.Recordset!Concepto
    TAB_TRA(i, 2) = TAB_TRADUCE_RUBROS.Recordset!CONTRADU

    Rem MsgBox Str(i) + " : " + TAB_TRA(i, 1) + "-->" + TAB_TRA(i, 2)
    
    
    TAB_TRADUCE_RUBROS.Recordset.MoveNext
    
Loop
'RDS.Close
MsgBox "Se cargaron :" + STR(i) + " , items de la Tabla de Traducción."
PBar_matriz.Value = 0
' DoCmd.OpenQuery "UPD_TAB_RUBROS_INCIDENCIA"

            
sqlstr = "SELECT *  FROM Tab_Rubros_Incidencia  "
sqlstr = sqlstr + " WHERE AÑO=" + "'" + Trim((Me.txt_año)) + "'"

Tab_Rubros_Incidencia.CommandType = adCmdText

Tab_Rubros_Incidencia.RecordSource = sqlstr

Tab_Rubros_Incidencia.Refresh

If Tab_Rubros_Incidencia.Recordset.EOF Then

    MsgBox "tabla de rubros Incidencia vacia"

    Exit Sub

End If


Do While Tab_Rubros_Incidencia.Recordset.EOF = False

    For i = MESDES To MESHAS
        
        Select Case i
            
            Case Is = 1
                        
                If NZ(Tab_Rubros_Incidencia.Recordset!Real, 0) > 0 Then
                    Tab_Rubros_Incidencia.Recordset!Real = Tab_Rubros_Incidencia.Recordset!Real - Tab_Rubros_Incidencia.Recordset!Ene_Mon
                Else
                    Tab_Rubros_Incidencia.Recordset!Real = 0
                End If
                
                Tab_Rubros_Incidencia.Recordset!Ene_Mon = 0
                Tab_Rubros_Incidencia.Recordset!Ene_Can = 0
                    
                mvBookMark = Tab_Rubros_Incidencia.Recordset.Bookmark
                Tab_Rubros_Incidencia.Recordset.Update
                Tab_Rubros_Incidencia.Recordset.Bookmark = mvBookMark
                            
            Case Is = 2
                        
                If NZ(Tab_Rubros_Incidencia.Recordset!Real, 0) > 0 Then
                    Tab_Rubros_Incidencia.Recordset!Real = Tab_Rubros_Incidencia.Recordset!Real - Tab_Rubros_Incidencia.Recordset!Feb_Mon
                Else
                    Tab_Rubros_Incidencia.Recordset!Real = 0
                End If
                
                Tab_Rubros_Incidencia.Recordset!Feb_Mon = 0
                Tab_Rubros_Incidencia.Recordset!Feb_Can = 0
                
                mvBookMark = Tab_Rubros_Incidencia.Recordset.Bookmark
                Tab_Rubros_Incidencia.Recordset.Update
                Tab_Rubros_Incidencia.Recordset.Bookmark = mvBookMark
                            
            Case Is = 3
                        
                If NZ(Tab_Rubros_Incidencia.Recordset!Real, 0) > 0 Then
                    Tab_Rubros_Incidencia.Recordset!Real = Tab_Rubros_Incidencia.Recordset!Real - Tab_Rubros_Incidencia.Recordset!Mar_Mon
                Else
                    Tab_Rubros_Incidencia.Recordset!Real = 0
                End If
                        
                Tab_Rubros_Incidencia.Recordset!Mar_Mon = 0
                Tab_Rubros_Incidencia.Recordset!Mar_Can = 0
              
                mvBookMark = Tab_Rubros_Incidencia.Recordset.Bookmark
                Tab_Rubros_Incidencia.Recordset.Update
                Tab_Rubros_Incidencia.Recordset.Bookmark = mvBookMark
                                   
            Case Is = 4
                        
                If NZ(Tab_Rubros_Incidencia.Recordset!Real, 0) > 0 Then
                    Tab_Rubros_Incidencia.Recordset!Real = Tab_Rubros_Incidencia.Recordset!Real - Tab_Rubros_Incidencia.Recordset!Abr_Mon
                Else
                    Tab_Rubros_Incidencia.Recordset!Real = 0
                End If
                
                Tab_Rubros_Incidencia.Recordset!Abr_Mon = 0
                Tab_Rubros_Incidencia.Recordset!Abr_Can = 0
     
                mvBookMark = Tab_Rubros_Incidencia.Recordset.Bookmark
                Tab_Rubros_Incidencia.Recordset.Update
                Tab_Rubros_Incidencia.Recordset.Bookmark = mvBookMark
                            
            Case Is = 5
                        
                If NZ(Tab_Rubros_Incidencia.Recordset!Real, 0) > 0 Then
                    Tab_Rubros_Incidencia.Recordset!Real = Tab_Rubros_Incidencia.Recordset!Real - Tab_Rubros_Incidencia.Recordset!May_Mon
                Else
                    Tab_Rubros_Incidencia.Recordset!Real = 0
                End If
                
                Tab_Rubros_Incidencia.Recordset!May_Mon = 0
                Tab_Rubros_Incidencia.Recordset!May_Can = 0
               
                mvBookMark = Tab_Rubros_Incidencia.Recordset.Bookmark
                Tab_Rubros_Incidencia.Recordset.Update
                Tab_Rubros_Incidencia.Recordset.Bookmark = mvBookMark
                    
            Case Is = 6
                    
                If NZ(Tab_Rubros_Incidencia.Recordset!Real, 0) > 0 Then
                    Tab_Rubros_Incidencia.Recordset!Real = Tab_Rubros_Incidencia.Recordset!Real - Tab_Rubros_Incidencia.Recordset!Jun_Mon
                Else
                    Tab_Rubros_Incidencia.Recordset!Real = 0
                End If
                
                Tab_Rubros_Incidencia.Recordset!Jun_Mon = 0
                Tab_Rubros_Incidencia.Recordset!Jun_Can = 0
                
                mvBookMark = Tab_Rubros_Incidencia.Recordset.Bookmark
                Tab_Rubros_Incidencia.Recordset.Update
                Tab_Rubros_Incidencia.Recordset.Bookmark = mvBookMark
                    
            Case Is = 7
                    
                If NZ(Tab_Rubros_Incidencia.Recordset!Real, 0) > 0 Then
                    Tab_Rubros_Incidencia.Recordset!Real = Tab_Rubros_Incidencia.Recordset!Real - Tab_Rubros_Incidencia.Recordset!Jul_Mon
                Else
                    Tab_Rubros_Incidencia.Recordset!Real = 0
                End If
                
                Tab_Rubros_Incidencia.Recordset!Jul_Mon = 0
                Tab_Rubros_Incidencia.Recordset!Jul_Can = 0
                
                
                mvBookMark = Tab_Rubros_Incidencia.Recordset.Bookmark
                Tab_Rubros_Incidencia.Recordset.Update
                Tab_Rubros_Incidencia.Recordset.Bookmark = mvBookMark
                    
            Case Is = 8
                        
                If NZ(Tab_Rubros_Incidencia.Recordset!Real, 0) > 0 Then
                    Tab_Rubros_Incidencia.Recordset!Real = Tab_Rubros_Incidencia.Recordset!Real - Tab_Rubros_Incidencia.Recordset!Ago_Mon
                Else
                    Tab_Rubros_Incidencia.Recordset!Real = 0
                End If
                
                Tab_Rubros_Incidencia.Recordset!Ago_Mon = 0
                Tab_Rubros_Incidencia.Recordset!Ago_Can = 0
                
                
                mvBookMark = Tab_Rubros_Incidencia.Recordset.Bookmark
                Tab_Rubros_Incidencia.Recordset.Update
                Tab_Rubros_Incidencia.Recordset.Bookmark = mvBookMark
                    
            Case Is = 9
                        
                 If NZ(Tab_Rubros_Incidencia.Recordset!Real, 0) > 0 Then
                    Tab_Rubros_Incidencia.Recordset!Real = Tab_Rubros_Incidencia.Recordset!Real - Tab_Rubros_Incidencia.Recordset!Sep_Mon
                Else
                    Tab_Rubros_Incidencia.Recordset!Real = 0
                End If
                
                Tab_Rubros_Incidencia.Recordset!Sep_Mon = 0
                Tab_Rubros_Incidencia.Recordset!Sep_Can = 0
                       
                mvBookMark = Tab_Rubros_Incidencia.Recordset.Bookmark
                Tab_Rubros_Incidencia.Recordset.Update
                Tab_Rubros_Incidencia.Recordset.Bookmark = mvBookMark
            Case Is = 10
                        
                If NZ(Tab_Rubros_Incidencia.Recordset!Real, 0) > 0 Then
                    Tab_Rubros_Incidencia.Recordset!Real = Tab_Rubros_Incidencia.Recordset!Real - Tab_Rubros_Incidencia.Recordset!Oct_Mon
                Else
                    Tab_Rubros_Incidencia.Recordset!Real = 0
                End If
                
                Tab_Rubros_Incidencia.Recordset!Oct_Mon = 0
                Tab_Rubros_Incidencia.Recordset!Oct_Can = 0
                mvBookMark = Tab_Rubros_Incidencia.Recordset.Bookmark
                Tab_Rubros_Incidencia.Recordset.Update
                Tab_Rubros_Incidencia.Recordset.Bookmark = mvBookMark
                    
            Case Is = 11
                        
                If NZ(Tab_Rubros_Incidencia.Recordset!Real, 0) > 0 Then
                    Tab_Rubros_Incidencia.Recordset!Real = Tab_Rubros_Incidencia.Recordset!Real - Tab_Rubros_Incidencia.Recordset!Nov_Mon
                Else
                    Tab_Rubros_Incidencia.Recordset!Real = 0
                End If
                
                Tab_Rubros_Incidencia.Recordset!Nov_Mon = 0
                Tab_Rubros_Incidencia.Recordset!Nov_Can = 0
                       
              mvBookMark = Tab_Rubros_Incidencia.Recordset.Bookmark
                Tab_Rubros_Incidencia.Recordset.Update
                Tab_Rubros_Incidencia.Recordset.Bookmark = mvBookMark
                       
            Case Is = 12
                        
                If NZ(Tab_Rubros_Incidencia.Recordset!Real, 0) > 0 Then
                    Tab_Rubros_Incidencia.Recordset!Real = Tab_Rubros_Incidencia.Recordset!Real - Tab_Rubros_Incidencia.Recordset!Dic_Mon
                Else
                    Tab_Rubros_Incidencia.Recordset!Real = 0
                End If
                
                Tab_Rubros_Incidencia.Recordset!Dic_Mon = 0
                Tab_Rubros_Incidencia.Recordset!Dic_Can = 0
                       
              mvBookMark = Tab_Rubros_Incidencia.Recordset.Bookmark
                Tab_Rubros_Incidencia.Recordset.Update
                Tab_Rubros_Incidencia.Recordset.Bookmark = mvBookMark
        
            End Select
            
    Next

    Tab_Rubros_Incidencia.Recordset.MoveNext
    
Loop

'RDS.Close

Exit Sub
control_de_errores:
    MsgBox " " & Err.Number & " :  " & Err.Description & " "

End Sub

'CONCEPTO    CONTRADU    DESCRIPCION
'301020500   301020500   IMPUESTO SOBRE INMUEBLES URBANOS
'301040301   301020500   IMPUESTO SOBRE INMUEBLES URBANOS
'301040302   301020500   MOROSIDAD AÑO ACTUAL
'301040303   301020500   INTERESES
'301040304   301100100   SOLVENCIAS   INMUEBLES URBANOS
'301040305   301040300   PRECANCELADOS
'301040306   301041000   MOROSIDAD AÑOS ANTERIORES
'301040500   301040500   IMPUESTO PATENTE DE IND.Y COMERCIO
'301020700   301040500   IMPUESTO PATENTE DE IND.Y COMERCIO
'301040503   301040500   INTERESES DE MORA Y RECARGO
'301040504   301100100   SOLVENCIAS   PATENTE DE INDUSTRIA Y COMERCIO
'301040505   301040500   PRECANCELADOS
'301040506   301041000   MOROSIDAD AÑOS ANTERIORES
'301040507   301120800   MULTAS Y RECARGOS
'301040508   301100100   LICENCIAS   DE  PATENTE DE INDUSTRIA Y COMERCIO.
'301040509   301120200   REPARO FISCAL
'301040510   301040500   DIFERENCIA COBRO(PAT/IND.Y COMERCIO)
'301040511   301040500   POOL
'301040512   301040500   CONVENIO DE PAGO
'301040513   301040500   COMPENSACION
'301040514   301040500   COMPLEMENTO TRIMESTRES
'301040520   301040500   APUESTAS LOCALES
'301020800   301020800   PATENTE DE VEHICULO.
'301040601   301020800   PATENTE DE VEHICULO.
'301040700   301040700   PUBLICIDAD
'301040701   301040700   PUBLICIDAD COMERCIAL
'301040800   301040800   ESPECTACULOS PUBLICOS
'301040900   301040900   APUESTAS  LICITAS
'301041000   301041000   DEUDA MOROSA
'301100100   301100100   TASAS
'301100102   301100100   COPIAS CERTIFICADAS DE DOCUMENTOS
'301100103   301100100   CONSTANCIAS EN GENERAL
'301100104   301100100   COPIAS CERTIFICADAS DE LICENCIAS DE PATENTE DE INDUSTRIA Y COMERCIO Y DE CONFORMIDADES DE USO
'301100105   301100100   COPIAS CERTIFICADAS DE CONSTANCIA DE DOMICILIO
'301100106   301100100   SOLVENCIAS O CONSTANCIAS DE CANCELACION DE GRAVAMENES FISCALES
'301100107   301100100   COPIAS CERTIFICADAS DE INFORMACIÓN CATASTRAL SOBRE INMUEBLES
'301100108   301100100   COPIAS HELIOGRAFICAS SIMPLES O REDUCCIÓN DE PLANOS EN LINEA AZUL
'301100109   301100100   COPIAS CERTIFICADAS DE PLANOS
'301100110   301100100   COPIAS DE FORMATOS ORIGINALES PARA PLANOS DE MENSURAS
'301100111   301100100   COPIAS HELIOGRÁFICAS SIMPLES DE PLANOS
'301100112   301100100   GACETAS MUNICIPALES CORRIENTES O DE ORDENANZAS
'301100113   301100100   AVALÚOS EN GENERAL (NOTA: SE CALCULA SOBRE EL VALOR RESULTANTE DEL AVALÚO)
'301100114   301100100   DOCUMENTOS REFERENTES A FOSAS, BÓVEDAS Y NICHOS
'301100115   301100100   CONTRATOS O TRANSACIONES SOBRE DERECHOS NO APRECIABLES
'301100116   301100100   CONTRATOS DE ARRENDAMIENTO DE BIENES E INMUEBLES MUNICIPALES
'301100117   301100100   CONTRATO DE COMPRA-VENTA DE BIENES E INMUEBLES MUNICIPALES
'301100118   301100100   CONTRATOS QUE IMPLIQUEN ENAJENACION DE OTROS BIENES INMUEBLES MUNICIPALES
'301100119   301100100   SOLICITUD DE CONFORMIDAD Y UBICACIÓN DE ZONA DE USO
'301100120   301100100   SOLICITUD PARA INTEGRACIÓN, REPARCELAMIENTO O UNION DE PARCELAS (NOTA: SE CALCULA POR Mª)
'301100121   301100100   REVISIÓN DE LA SOLICITUD SOBRE CONFORMACIÓN O RECTIFICACIÓN DE LINDEROS (NOTA: SE CALCULA POR Mª)
'301100122   301100100   APLICACIONES, MODIFICACIONES Y CAMBIOS DE USO
'301100123   301100100   MENSURAS DE TERRENOS (NOTA: SE CALCULA POR Mª)
'301100124   301100100   PLANILLAS RELATIVAS A FORMATOS DE CUALQUIER NATURALEZA
'301100126   301100100   OTRAS TASAS.
'301100137   301100100   COPIAS CERTIFICADAS DE PLANOS
'301100147   301100100   PERMISOS MUNICIPALES
'301100148   301100100   CERTIFICACIONES Y SOLVENCIAS
'301100149   301100100   MENSURA Y DESLINDE
'301100150   301100150   ASEO DOMICILIARIO
'301100158   301100100   DEUDA MOROSA POR TASAS
'301100199   301100100   OTROS  TIPOS DE TASAS MUNICIPALES
'301120200   301120200   REPAROS FISCALES
'301120301   301120301   REINTEGROS  DE  PARTICULARES
'301120302   301120302   REINTEGROS FONDOS GIRADOS EN AVANCE
'301120499   301120499   INGRESOS VARIOS
'301120500   301120500   RECARGOS E INTERESES MORATORIOS
'301120705   301120705   INDEMNIZACION P/INCUMPLIMIENTO D/CONTRATO
'301120800   301120800   MULTAS Y RECARGOS
'301130102   301120102   VENTAS DE GACETAS MCPALES Y FORMULARIOS.
'301130108   301130108   OPERACIONES DIVERSAS MUNICIPIOS
'301140200   301140200   INTERESES POR DEPOSITOS
'301140300   301140300   INTERESES POR TITULOS Y VALORES
'301140401   301140401   RENTAS INMOBILIARIAS
'301140501   301140501   ARRENDAMIENTOS DE EJIDOS
'301160000   301160000   INGRESOS POR APORTE AL MUNICIPIO
'301160100   301160100   SITUADO  MUNICIPAL
'301160301   301160301   APORTES ESPECIAL GOBERNACION
'301160302   301160302   APORTE  ESPECIAL  GOBERNO NACIONAL
'301160303   301160303   APORTE  ESPECIAL  FONTUR
'301160304   301160304   APORTE  -  FIDES
'301160305   301160305   APORTE  LEY ASIG. ECONOMICAS ESPECIALES
'301160306   301160306   APORTE   ESPECIAL - PAE
'301160307   301160307   APORTE   ESP. MINISTERIO DE    CIENCIAS Y TEC.(MCT).
'301160308   301160308   APORTE   ESP. MINISTERIO DE    INFRAESTRUCTURA.
'301160309   301160309   APORTE   ESP. MINISTERIO DEL  AMBIENTE.
'302020300   302020300   RESERVAS D/TESORO MCPAL. NO COMPROMETIDO.
'302060200   302060200  VENTAS DE BIENES

Private Sub txt_hasta_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub
