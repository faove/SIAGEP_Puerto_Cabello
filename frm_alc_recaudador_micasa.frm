VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_alc_recaudador_micasa 
   Caption         =   "Recaudación  /  S I A G E P"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6015
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   11535
      Begin VB.CommandButton CommandButton 
         Caption         =   "&Salir"
         Height          =   615
         Index           =   2
         Left            =   9960
         TabIndex        =   0
         ToolTipText     =   "Cerrar Recaudación"
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CheckBox Check 
         Caption         =   "Mostrar Todo"
         Height          =   615
         Index           =   0
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CheckBox Check 
         Caption         =   "Impresión Media Carta"
         Height          =   615
         Index           =   1
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CheckBox Check 
         Caption         =   "Impresión Tiquera"
         Height          =   615
         Index           =   2
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5280
         Width           =   1575
      End
      Begin VB.TextBox txt_Cuotas 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   4680
         Width           =   2055
      End
      Begin VB.TextBox txt_Voucher 
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   4680
         Width           =   2055
      End
      Begin VB.TextBox txt_MontoV 
         Height          =   285
         Left            =   4560
         TabIndex        =   7
         Top             =   4680
         Width           =   2055
      End
      Begin VB.TextBox txt_Monto 
         Height          =   285
         Left            =   6720
         TabIndex        =   8
         Top             =   4680
         Width           =   2055
      End
      Begin VB.TextBox txt_Saldo_Voucher 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   8880
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4680
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid Dgrid_planilla 
         Bindings        =   "frm_alc_recaudador_micasa.frx":0000
         Height          =   3735
         Left            =   0
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "Nro_Plani_Pago"
            Caption         =   "Nro_Plani_Pago"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Xnombre"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Id_Objeto"
            Caption         =   "Id_Objeto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Id_Instancia"
            Caption         =   "Id_Instancia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Cuota"
            Caption         =   "Cuota"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Rubro"
            Caption         =   "Rubro"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Monto_Origi"
            Caption         =   "Monto_Origi"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """Bs"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Fec_Pago"
            Caption         =   "Fec_Pago"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Tip_Liq"
            Caption         =   "Tip_Liq"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "status"
            Caption         =   "status"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "Id_Aso"
            Caption         =   "Id_Aso"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "descuento"
            Caption         =   "descuento"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3555,213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   824,882
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1560,189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1200,189
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1335,118
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   540,284
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1065,26
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "&Reinicio"
         Height          =   615
         Index           =   1
         Left            =   3720
         TabIndex        =   2
         ToolTipText     =   "Reinicio"
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "&Aceptar"
         Height          =   615
         Index           =   0
         Left            =   2160
         TabIndex        =   1
         ToolTipText     =   "Aceptar Liquidación"
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Label lbl_Saldo_Voucher 
         Caption         =   "Saldo del Voucher:"
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
         Left            =   8880
         TabIndex        =   19
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label lbl_Cuotas_Nro 
         Caption         =   "Items/Cuotas"
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
         Left            =   240
         TabIndex        =   18
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label lbl_Voucher 
         Caption         =   "Ingrese Nro. Voucher"
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
         Left            =   2400
         TabIndex        =   17
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label lbl_MontoV 
         Caption         =   "Monto de Voucher"
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
         Left            =   4560
         TabIndex        =   16
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label lbl_Monto 
         Caption         =   "Monto Liquidado"
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
         Left            =   6720
         TabIndex        =   15
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   11400
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   3
         Index           =   1
         X1              =   0
         X2              =   11400
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Lista de Nro. Planilla Liquidación"
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
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   8520
      TabIndex        =   9
      Top             =   240
      Width           =   3255
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   " RECAUDACIÓN"
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
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000003&
         Caption         =   "  SEMAT "
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
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc CUM_RECAUDACION 
      Height          =   375
      Left            =   360
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
      RecordSource    =   "SELECT * FROM CUM_RECAUDACION"
      Caption         =   "CUM_RECAUDACION"
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
Attribute VB_Name = "frm_alc_recaudador_micasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Gtip_liq As String
Public Aceptar As Boolean
Public Gfec_pago As Date
Public Grazon_social As String
Public Gid_obj As String
Public Gid_instancia    As String
Public Gcuota As String
Public Grubro   As String
Public Gmonto As Double
Public Grenglon As Byte
Public sqlstr As String
Public Fecha As Date
Public cuotas As ADODB.Recordset
Public TOT_MONTO As Double
Public tren_transas As String
Public rds As ADODB.Recordset
'Public BDS As Database
Public Mdescuento As Single
Public N_Voucher As String
Public control, control2, Abandono, Imp_Tiquera As Boolean


Private Sub cmd_aceptar_Click()

On Error GoTo Errores

Dim cadena As String

SCROLL 0
Screen.MousePointer = 11

Fecha = Date

cn.BeginTrans

SCROLL 10
'Llamada a crear la seleccion del reporte
'----------------------------------------
'Call SELECCION_ALC_RECAUDADOR '**************************falta*******************
'Wrk.BeginTrans


Rem Test tipo de liquidacion en proceso : 1) Liquidacion Especifica/Regular por porciones/cuotas regulares,
Rem 2) Liquidacion Generica tales : Tasas y Conceptos de Tributos tales como: Renovaciones / Licencias / Multas,
Rem Cont..2) Multas/Recargos/Mora.

If Gtip_liq = "Gen" Then
    
    SCROLL 20
    
    If PROCESA_OBJETO_GENERICO = False Then
        SCROLL 41
        Screen.MousePointer = 0
'        Me.Etiqueta42.Visible = False
        SCROLL 0
        Me.txt_Saldo_Voucher.Visible = False
        Me.lbl_Saldo_Voucher.Visible = False
        Exit Sub
    End If
 
Else
    
    SCROLL 30
    
    If Procesa_Objeto_Especifico = False Then
        SCROLL 41
        Screen.MousePointer = 0
'        Me.Etiqueta42.Visible = False
        SCROLL 0
        Me.txt_Saldo_Voucher.Visible = False
        Me.lbl_Saldo_Voucher.Visible = False
        Exit Sub
    End If
 
End If


Rem   Crea la Forma_de_Pago
ABRIR_FORMA_DE_PAGO 'PERMITE EDICIÓN OPTIMISTA

Rdsfpago.AddNew
    
    Rdsfpago!NRO_PLANI_PAGO = Gcod_planilla
    Rdsfpago!Id_Objeto = Gid_obj
    Rdsfpago!Id_Instancia = Gid_instancia
    Rdsfpago!Id_Rubro = Grubro
    Rdsfpago!monto = Format(txt_Monto, "CURRENCY")
    Rdsfpago!Fec_Pago = Format(Date, "dd/mm/yyyy")
    Rdsfpago!Id_Voucher = Me.txt_Voucher
    Rdsfpago!STATUS = "CA"
    Rdsfpago!Usuario = Usuario                 '*********************OJO****************
SCROLL 35
Rdsfpago.Update

Me.txt_MontoV = Format(Me.txt_MontoV, "CURRENCY")

If Gtip_liq = "Esp" Then

    cuotas.MoveFirst
    SCROLL 37

    While cuotas.EOF = False
        SCROLL 39
        sqlstr = "Update Alc_Obj_Liqs Set Status = 'CA',Voucher=" + "'" + (Me.txt_Voucher)
        sqlstr = sqlstr + "'" + ",monto_voucher=" + "'" + STR(Me.txt_MontoV) + "'"
        sqlstr = sqlstr + " , RUBRO = " + "'" + (cuotas!Concepto) + "', usuario_rec = '" & Usuario & "'"
        sqlstr = sqlstr + "  Where Nro_Plani_pago = " + "'" + (Me.Dgrid_planilla) + "'"
        sqlstr = sqlstr + "  And   Cuota  = " + "'" + (cuotas!CUOTA) + "'" + ";"

        cn.Execute sqlstr
        cuotas.MoveNext

    Wend

Else
        SCROLL 38
        sqlstr = "Update Alc_Obj_Liqs Set Status = 'CA',Voucher=" + "'" + (Me.txt_Voucher) + "', usuario_rec = '" & Usuario & "'"
        sqlstr = sqlstr + " , RUBRO = " + "'" + (Grubro) + "'" & ", monto_voucher=" + "'" + STR(Me.txt_MontoV) + "'"
        sqlstr = sqlstr + "  Where Nro_Plani_pago = " + "'" + (Me.Dgrid_planilla) + "'" + ";"

        cn.Execute sqlstr
End If

Rem grabar_log_operas ojo con los parms
SCROLL 40
Gcod_Transa = FGNRO_TRAN()

'Grabar_Operacion Gcod_planilla, Gcod_Transa, Gid_obj, Gid_instancia, Gitems, txt_Monto, Grubro, Goficina, Gtaquilla, Guser_id, tren_transas

Rem   TEST ULTIMA OPORTUNIDAD DE APLICAR LA OPERACION / ABORTARLA

Screen.MousePointer = 0

    If MsgBox("¿ Desea Aplicar la Recaudación ?", vbYesNo, "Confirmación Operación Recaudacion.") = vbNo Then
        
        cn.RollbackTrans
        
        MsgBox "Transacción Reversada  Por Operador.", , " Transacción de Recaudación Abortada."
'        Me.Etiqueta42.Visible = False
        SCROLL 0
        Exit Sub
    End If

Rem End Transaction Succesfull

cn.CommitTrans
'cn.Close
Rem Imprime Vista Previa Planilla de Liquidación de Recaudación

imprimir

'Dgrid_planilla = Null
txt_Cuotas = 0
txt_Monto = 0
txt_Voucher = ""
txt_MontoV = 0

'Rdsfpago.Close

'Dgrid_planilla.SetFocus

Me.CommandButton(0).Enabled = False
Me.CommandButton(1).Enabled = False

Me.lbl_Saldo_Voucher.Visible = False
Me.txt_Saldo_Voucher.Visible = False

SCROLL 41

'Call actualizar_cn("SQL Server")

Call cmd_reinicio_Click
'SCROLL 0
Exit Sub
Errores:
    If Err.Number = -2147217873 Then
        MsgBox "Factura Procesada", vbInformation, "ALCASIS"
    Else: MsgBox Err.Description
    End If
        
    Screen.MousePointer = 0
    SCROLL 0
    Dgrid_planilla.Refresh
    cn.RollbackTrans
    'cn.RollbackTrans '21/11/2002
    cn.Close
    Call actualizar_cn("SQL Server")


End Sub

Private Sub imprimir()

Dim sqlstr As String
Dim cadena As String
Dim i As Integer

'nro_copias

copias = 1

If copias <> 99 Then

    cadena = "Nro_Plani_Pago = '" + Dgrid_planilla + "'"
    
    'Si copias = 0 presenta vista preliminar

    If Gtip_liq = "Gen" Then
        If Usuario = 15 Then
            rpt_alc_recuadacion_gen.Show
'        DoCmd.OpenReport "ALC_RECAUDACION_GENERICAML", acViewPreview, , cadena, acDialog
        Else
                    If Imp_Tiquera Then
                        rpt_cr_alc_rec_tiq_gen.Show
                    Else
                        rpt_alc_recuadacion_gen.Show
                    End If
            'DoCmd.OpenReport "ALC_RECAUDACION_GENERICA", acViewPreview, , cadena, acDialog
        End If
        
    Else  ' Especifica
         
            Select Case Gid_obj
                   '**************************************************************************
                  Case "PIC"
                  If Usuario = 15 Then
                    rpt_alc_recaudacion_pic.Show
 '                 DoCmd.OpenReport "ALC_RECAUDACION_PICML", acViewPreview, , cadena, acDialog
                  Else
                    If Imp_Tiquera Then
                        rpt_cr_alc_rec_tiq_pic.Show
                    Else
                        rpt_alc_recaudacion_pic.Show
                    End If
                    'rpt_alc_recaudacion_pic_tiquera.Show
 '                 DoCmd.OpenReport "ALC_RECAUDACION_PIC", acViewPreview, , cadena, acDialog
                  End If
                  
                  
                  '**************************************************************************
                   
                   Case "INM"
                      If Usuario = 15 Then
                        rpt_alc_recaudacion_inm.Show
'                      DoCmd.OpenReport "ALC_RECAUDACION_INMML", acViewPreview, , cadena, acDialog
                      Else
                        If Imp_Tiquera Then
                            rpt_cr_alc_rec_tiq_inm.Show
                        Else
                            rpt_alc_recaudacion_inm.Show
                        End If

'                      DoCmd.OpenReport "ALC_RECAUDACION_INM", acViewPreview, , cadena, acDialog
                      End If
                   
                   '**************************************************************************
            
                    Case "PUB"
                        If Usuario = 15 Then
                            rpt_alc_recaudacion_pub.Show
'                        DoCmd.OpenReport "ALC_RECAUDACION_PUBML", acViewPreview, , cadena, acDialog
                        Else
                        If Imp_Tiquera Then
                            rpt_cr_alc_rec_tiq_pub.Show
                        Else
                            rpt_alc_recaudacion_pub.Show
                        End If
'                        DoCmd.OpenReport "ALC_RECAUDACION_PUB", acViewPreview, , cadena, acDialog
                        End If
                        
                   '**************************************************************************
            
                    Case "VEH"
                        If Usuario = 15 Then
                            rpt_alc_recaudacion_veh.Show
'                        DoCmd.OpenReport "ALC_RECAUDACION_VEHML", acViewPreview, , cadena, acDialog
                        Else
                        If Imp_Tiquera Then
                            rpt_cr_alc_rec_tiq_veh.Show
                        Else
                            rpt_alc_recaudacion_veh.Show
                        End If
'                        DoCmd.OpenReport "ALC_RECAUDACION_VEH", acViewPreview, , cadena, acDialog
                        End If
                        
                    '**************************************************************************
                    Case "APU"
                        If Imp_Tiquera Then
                            rpt_cr_alc_rec_tiq_apu.Show
                        Else
                            rpt_alc_recaudacion_apu.Show
                        End If
                    '**************************************************************************
                    Case "ADU"
                       RPT_alc_recaudacion_adu.Show

            End Select
    End If
End If

'    ELI_ALC_OBJ_LIQ
'-------------------
    
      Dgrid_planilla.Refresh
   
      'Dgrid_planilla = Null
      
      txt_Cuotas = 0
      
      txt_Monto = 0
    
End Sub

Private Sub cmd_reinicio_Click()
Me.CUM_RECAUDACION.Refresh
Me.lbl_Voucher.Caption = "Ingrese Nro. Voucher"
lbl_Saldo_Voucher.Visible = False
txt_Saldo_Voucher.Visible = False
control = True
control2 = True
N_Voucher = ""
Me.txt_Voucher = ""
Me.txt_MontoV.Locked = False
Me.txt_MontoV = ""
Me.txt_Monto = ""
Me.txt_Cuotas = ""

End Sub

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub cmd_salir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Check_Click(Index As Integer)

If Check(Index).Value = 0 Then
    Me.Check(Index).BackColor = vbButtonFace
Else
    Me.Check(Index).BackColor = vbRed
End If
    
Select Case Index
    Case 0
        With Me.CUM_RECAUDACION
        .ConnectionString = "DSN=SIAGEP"
        .CommandType = adCmdText
        If Check(Index).Value = 0 Then
            .RecordSource = "SELECT * FROM CUM_RECAUDACION WHERE usuario_liq = '" & Usuario & "'"
            'Me.Check(Index).Caption = "Mostrar Todo"
        Else
            .RecordSource = "SELECT * FROM CUM_RECAUDACION_F "
            'Me.Check(Index).Caption = "Mostrar Todo Activado"
            
        End If
        .Refresh
        End With
    Case 1
        Imp_auto = Check(Index).Value
    Case 2
        Imp_Tiquera = Check(Index).Value

End Select
End Sub

Private Sub Check_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 1
Me.Check(i).FontBold = False
Next i
For i = 0 To 2
Me.CommandButton(i).FontBold = False
Next i

Me.Check(Index).FontBold = True
Call Descripcion(Me.Check(Index).Tag)

End Sub

Private Sub CommandButton_Click(Index As Integer)
Select Case Index
    Case 0
        Call cmd_aceptar_Click
    Case 1
        Call cmd_reinicio_Click
    Case 2
        Call cmd_salir_Click
       
End Select
'If user_grupo <> "04" Then
'    Unload Me
'End If
End Sub

Private Sub CommandButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 2
Me.CommandButton(i).FontBold = False
Next i
For i = 0 To 1
Me.Check(i).FontBold = False
Next i

Me.CommandButton(Index).FontBold = True
Call Descripcion(Me.CommandButton(Index).Tag)

End Sub


Private Sub Dgrid_planilla_Click()

On Error GoTo ControlError

Dim VAR As Variant

control = True

control2 = True

Me.txt_Voucher = ""

Me.txt_MontoV.Locked = False

Me.txt_MontoV = ""

Gcod_planilla = Dgrid_planilla

'Verifica que Bookmarks tenga algun valor, para efectuar la transaccion
'----------------------------------------------------------------------

If Dgrid_planilla.SelBookmarks.Count = 0 Then

    MsgBox "Debe seleccionar una Planilla de Liquidación", vbInformation
    
    Exit Sub
    
End If

'Verifica que la seleccion realizada por el usuario tenga numero de planilla
'---------------------------------------------------------------------------

If Dgrid_planilla.Columns(0).Text = "" Then

    MsgBox "No Existe Registro de Liquidación Previa para" & Gcod_planilla
    
    Exit Sub
    
End If
'Verifica que el usuario selecione una sola fila
'-----------------------------------------------

If Dgrid_planilla.SelBookmarks.Count > 1 Then

    Dgrid_planilla.SelBookmarks.Remove (Dgrid_planilla.SelBookmarks.Count - 1)
    
    Exit Sub
    
End If
Rem  Sumariza Montos liquidados por cada rengon de la planilla

Gitems = 0

txt_Monto = 0

Dim sqlstr As String
            
sqlstr = "Select *  From Alc_Obj_liqs  "
sqlstr = sqlstr + "  Where Nro_Plani_pago = " + "'" + (Gcod_planilla) + "'" + ";"

Set rds = New ADODB.Recordset

rds.LockType = adLockReadOnly

rds.CursorType = adOpenKeyset

rds.Open sqlstr, cn
            
If rds.EOF Then

    MsgBox "No Existe Registro de Liquidación Previa para: " + Gcod_planilla
    
    Exit Sub

End If

While rds.EOF = False

    Gitems = Gitems + 1
    
    txt_Monto = txt_Monto + rds!Monto_Origi
       
    txt_Monto = Format(txt_Monto, "CURRENCY")
       
    rds.MoveNext
    
Wend

    rds.MoveFirst
    
    Mdescuento = 0
    
    Gmonto = rds!Monto_Origi
    
    txt_Cuotas = rds!CUOTA
    
    Gtip_liq = rds!Tip_Liq
    
    Gid_obj = rds!Id_Objeto
    
    Gid_instancia = rds!Id_Instancia
    
    Grubro = rds!Rubro
    
    Mdescuento = NZ(rds!descuento, 0)
    
    'Hab/Desabilita los Botones Pertinentes
    '--------------------------------------
    Me.CommandButton(0).Enabled = False
    
    Me.CommandButton(1).Enabled = True
 
    'Hace Visible campos auxiliares de la transaccion
    '------------------------------------------------
    lbl_Cuotas_Nro.Visible = True
    
    Me.txt_Cuotas.Visible = True
    
    Lbl_monto.Visible = True
    
    Me.txt_Monto.Visible = True
    
    Me.txt_Voucher.Enabled = True
    
    Me.txt_Voucher.SetFocus
    
    'Call txt_MontoV_Click

Exit Sub
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Error en la selección", vbOKOnly, "ALCASIS"

    End Select
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    lbl_Saldo_Voucher.Visible = False
    txt_Saldo_Voucher.Visible = False
    
With Me.CUM_RECAUDACION
.ConnectionString = "DSN=SIAGEP"
.CommandType = adCmdText
.RecordSource = "SELECT * FROM CUM_RECAUDACION WHERE usuario_liq = '" & Usuario & "'"
.Refresh
End With
Me.Check(1).Value = 1
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 2
Me.CommandButton(i).FontBold = False
Next i
For i = 0 To 1
Me.Check(i).FontBold = False
Next i

Call Descripcion("")

End Sub

Private Sub txt_MontoV_GotFocus()
If Not control Then
    Call VExistV
    Me.lbl_Voucher.Caption = "Ingrese Nro. Voucher"
    control2 = Not control
    'ChequeoV ' llamada a confirmar si existe voucher
    Exit Sub
End If

'If control = True And control2 = True Then
If control Then
    N_Voucher = Me.txt_Voucher
    control = Not control
    'control2 = False
    Me.lbl_Voucher.Caption = "Reingrese N° Voucher"
    Me.txt_Voucher = ""
    Me.txt_Voucher.SetFocus
End If

End Sub
Private Sub VExistV()
'------------------------------------CHEQUEA SI EXISTE EL VOUCHER----------------
    Dim p_voucher As ADODB.Recordset
    Dim sqlstr As String
    Dim cargos, M_Voucher As Double


    sqlstr = "Select monto_origi,monto_voucher,nro_plani_pago From  alc_obj_liqs "
    sqlstr = sqlstr & "  Where (voucher='" & Me.txt_Voucher & "') and (status <> 'AN');"
    
    Set p_voucher = New ADODB.Recordset
    p_voucher.CursorType = adOpenForwardOnly
    p_voucher.LockType = adLockReadOnly
    p_voucher.Open sqlstr, cn
    
With p_voucher


If Not .EOF Then
    cargos = 0
    .MoveFirst
    
    If IsNull(!monto_voucher) Then
        M_Voucher = 0
        control2 = True
    Else
        M_Voucher = Format(!monto_voucher, "0.00")
    End If
    
    Do While Not .EOF
        cargos = cargos + !Monto_Origi
        .MoveNext
    Loop
    
    'AGREGAR SUMA + LO QUE SE VA A PAGAR
    '-----------------------------------
    If CStr(cargos + Me.txt_Monto) > M_Voucher Then
        'Saldo = M_Voucher - Suma
        'If Not control2 Then
        
            MsgBox "El Voucher N° " & Me.txt_Voucher & " con saldo de: " & Format((M_Voucher - cargos), "0.00") & " Bs." & Chr(13) & _
            "no puede cancelar la planilla: " & Gcod_planilla
            Me.txt_Voucher = ""
            Me.txt_Voucher.SetFocus
            control = True
'        Else
'            MsgBox "El Voucher N° " & Me.txt_Voucher & " con saldo de: " & Format((M_Voucher - cargos), "0.00") & " Bs." & Chr(13) & _
'            "no puede cancelar la planilla: " & Gcod_planilla
'            Me.txt_Voucher = ""
        
        'End If
    Else
        Me.txt_MontoV = M_Voucher
        
        Me.lbl_Saldo_Voucher.Visible = True
        Me.txt_Saldo_Voucher.Visible = True
        
        Me.txt_Saldo_Voucher = Format(M_Voucher - cargos, "0.00")
        Me.txt_Saldo_Voucher.Locked = True
        Me.txt_MontoV.Locked = True
        If control = False Then CommandButton(0).Enabled = True
        Me.CommandButton(0).SetFocus
    End If
    
'    Exit Sub
End If
End With
End Sub
'Private Sub txt_MontoV_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 9 Or KeyCode = 13 Then
'        Dim respuesta As String
'        aceptar = True
'        respuesta = MsgBox("¿Está de acuerdo con el Monto Suministrado: " & Me.txt_MontoV.Text & " Bs. ?", vbYesNo + vbInformation, "ALCASIS")
'        If respuesta <> vbYes Then
'            aceptar = False
'            Me.txt_MontoV = ""
'            Me.txt_MontoV.SetFocus
'            Exit Sub
'        End If
'    End If
'End Sub
Private Sub txt_MontoV_KeyPress(KeyAscii As Integer)

If KeyAscii = 46 Then KeyAscii = 44

If KeyAscii = 9 Or KeyAscii = 13 Then
    
'    If Not aceptar Then Exit Sub
    
    
      Dim respuesta As String
        respuesta = MsgBox("¿Está de acuerdo con el Monto Suministrado? Bs. " & Me.txt_MontoV, vbYesNo + vbInformation, "ALCASIS")
        If respuesta <> vbYes Then
            Exit Sub
        End If
    
    Dim NumMonto1, NumMonto2 As Double
    
    NumMonto1 = CDbl(Me.txt_Monto)
'    NumMonto2 = Format(Me.txt_MontoV, "0,00")
    If CDbl(Me.txt_MontoV) < NumMonto1 Then
        MsgBox "Error en el monto del Voucher (monto inferior)", _
        vbCritical + vbOKOnly
        Me.txt_MontoV = ""
        'Me.txt_Voucher = ""
        'Me.txt_Voucher.SetFocus
        'control = True
        'control2 = True
'        Exit Sub
    Else
        If CDbl(Me.txt_MontoV) > NumMonto1 Then
            MsgBox "Monto a favor"
'            Exit Sub
        End If
    End If

    
    Gvoucher = Me.txt_Voucher

'    Com_Aceptar.Enabled = True
        
    Me.txt_MontoV.Enabled = True
    
    CommandButton(0).Enabled = True
    CommandButton(0).SetFocus
    
End If
End Sub

Private Sub txt_Voucher_KeyPress(KeyAscii As Integer)
    
If KeyAscii = 13 Then SendKeys "{tab}"

    
End Sub

'Al salir del boton aceptar
'**********************************************ojo****************
'Screen.MousePointer = 0
'Me.lblqueta42.Visible = False
'Form_ALC_RECAUDADOR_MICASA.Dgrid_planilla.Requery
'**********************************************ojo****************


Private Function Procesa_Objeto_Especifico() As Boolean

Rem Selecciona modalidad de Recaudacion segun objeto de tributación
sqlstr = "Select *  From Alc_Obj_liqs  "
sqlstr = sqlstr + "  Where Nro_Plani_pago = " + "'" + (Gcod_planilla) + "'" + ";"
Set rds = New ADODB.Recordset
rds.LockType = adLockReadOnly
rds.CursorType = adOpenKeyset
rds.Open sqlstr, cn

If rds.EOF Then

    MsgBox "No Existe Registro de Liquidación Previa para: " + Gcod_planilla
    
    Exit Function
  

End If

Select Case Gid_obj
'****************************************************************************
'PIC
       Case "PIC"
        
        'rds.MoveFirst
        
        While rds.EOF = False
        
           Grubro = rds!Rubro
           Gmonto = rds!Monto_Origi
                    
           If Mdescuento > 0 And Grubro = "301020700" Then
        
                Grubro = "301040505"
           
            End If
    
            If rds!CUOTA = "200205" Then
            
                Grubro = "301040508"
                
            Else
            
                If rds!CUOTA = "200207" Then
            
                    Grubro = "301040507"
                
                
                End If
            End If
        
            sqlstr = "Update Cum_Fac Set "
            sqlstr = sqlstr + " Cum_Fac.Status='CA', Cum_Fac.Fec_Cancel = getdate()"
            sqlstr = sqlstr + ", Cum_Fac.Monto=" + "'" + (STR(Gmonto)) + "'"
            sqlstr = sqlstr + ", Cum_Fac.Concepto=" + "'" + (Grubro) + "', usuario_rec = '" & Usuario & "'"
            sqlstr = sqlstr + "  Where Cum_Fac.Id_Obj=" + "'" + (Gid_obj) + "'"
            sqlstr = sqlstr + "  And  Cum_Fac.Id_Instancia = " + "'" + (Gid_instancia) + "'"
            sqlstr = sqlstr + "  And  Cum_Fac.Nro_Plani_pago=" + "'" + (Gcod_planilla) + "'"
            sqlstr = sqlstr + "  And  Cum_Fac.Cuota=" + "'" + (rds!CUOTA) + "'" + ";"
          
            Dim cadena As String
            cn.Execute sqlstr, cadena
            
            If cadena = 0 Then

                MsgBox "Cuotas/Facturas  No Se Actualizaron. Seleccione de Nuevo: " + STR(cadena)
            
                'Wrk.Rollback
                cn.RollbackTrans
                
                Procesa_Objeto_Especifico = False
            
                Exit Function
            
            End If

            rds.MoveNext
               
        Wend
        
        sqlstr = "SELECT * FROM CUM_FAC "
        sqlstr = sqlstr + "  Where Cum_Fac.Id_Obj=" + "'" + (Gid_obj) + "'"
        sqlstr = sqlstr + "And   Cum_Fac.Id_Instancia = " + "'" + (Gid_instancia) + "'"
        sqlstr = sqlstr + "  And Cum_Fac.Nro_Plani_pago=" + "'" + (Gcod_planilla) + "'" + ";"
        
        Set cuotas = New ADODB.Recordset
        cuotas.Open sqlstr, cn
        cuotas.MoveFirst
        
        Do While cuotas.EOF = False
        
           tren_transas = tren_transas + ";" + cuotas!CUOTA
            
            cuotas.MoveNext
   
        Loop 'PIC
'****************************************************************************
'INMUEBLES
       Case "INM"
       
       rds.MoveFirst
        
        While rds.EOF = False
        
           Grubro = rds!Rubro
           Gmonto = rds!Monto_Origi
           
           If Mdescuento > 0 And Grubro = "301040301" Then
        
                Grubro = "301040305"
                
            End If
 
                   
            sqlstr = "Update Cum_Fac Set "
            sqlstr = sqlstr + "Cum_Fac.Status='CA', Cum_Fac.Fec_Cancel = getdate()"
            sqlstr = sqlstr + ", Cum_Fac.Monto=" + "'" + (STR(Gmonto)) + "'"
            sqlstr = sqlstr + ", Cum_Fac.Concepto=" + "'" + (rds!Rubro) + "', usuario_rec = '" & Usuario & "'"
            sqlstr = sqlstr + "  Where Cum_Fac.Id_Obj='INM' And  Cum_Fac.Id_Instancia = " + "'" + (Gid_instancia) + "'"
            sqlstr = sqlstr + "  And Cum_Fac.Nro_Plani_pago=" + "'" + (Gcod_planilla) + "'"
            sqlstr = sqlstr + "  And Cum_Fac.Cuota=" + "'" + (rds!CUOTA) + "'" + ";"
          
            cn.Execute sqlstr, cadena
            
            If cadena = 0 Then

                MsgBox "Cuotas/Facturas  No Se Actualizaron. Seleccione de Nuevo: " + cadena
            
                cn.RollbackTrans
            
                Procesa_Objeto_Especifico = False
            
                Exit Function
            
            End If

            rds.MoveNext
               
        Wend
        
        sqlstr = "SELECT * FROM CUM_FAC "
        sqlstr = sqlstr + "  Where Cum_Fac.Id_Obj='INM'  And   Cum_Fac.Id_Instancia = " + "'" + (Gid_instancia) + "'"
        sqlstr = sqlstr + "  And Cum_Fac.Nro_Plani_pago=" + "'" + (Gcod_planilla) + "'" + ";"
        
        Set cuotas = New ADODB.Recordset
        cuotas.Open sqlstr, cn
        cuotas.MoveFirst
       
  Rem      Gid_rubro = cuotas!Rubro

        tren_transas = cuotas!CUOTA
        
        cuotas.MoveNext
        
        Do While cuotas.EOF = False
        
            tren_transas = tren_transas + ";" + cuotas!CUOTA
            
            cuotas.MoveNext
   
        Loop 'INMUEBLES
        
'****************************************************************************
'PUBLICIDAD
Case "PUB"
       
       rds.MoveFirst
        
        While rds.EOF = False
        
           Grubro = rds!Rubro
           Gmonto = rds!Monto_Origi
           
           If Mdescuento > 0 And Grubro = "301040301" Then
        
                Grubro = "301040305"
                
            End If
 
                   
            sqlstr = "Update Cum_Fac Set "
            sqlstr = sqlstr + "Cum_Fac.Status='CA', Cum_Fac.Fec_Cancel = getdate()"
            sqlstr = sqlstr + ", Cum_Fac.Monto=" + "'" + (STR(Gmonto)) + "'"
            sqlstr = sqlstr + ", Cum_Fac.Concepto=" + "'" + (rds!Rubro) + "', usuario_rec = '" & Usuario & "'"
            sqlstr = sqlstr + "  Where Cum_Fac.Id_Obj='PUB' And  Cum_Fac.Id_Instancia = " + "'" + (Gid_instancia) + "'"
            sqlstr = sqlstr + "  And Cum_Fac.Nro_Plani_pago=" + "'" + (Gcod_planilla) + "'"
            sqlstr = sqlstr + "  And Cum_Fac.Cuota=" + "'" + (rds!CUOTA) + "'" + " and cum_fac.id_aso = '" & rds!id_aso & "';"
          
            cn.Execute sqlstr, cadena
            
            If cadena = 0 Then

                MsgBox "Cuotas/Facturas  No Se Actualizaron. Seleccione de Nuevo: " + cadena
            
                cn.RollbackTrans
            
                Procesa_Objeto_Especifico = False
            
                Exit Function
            
            End If

            rds.MoveNext
               
        Wend
        
        sqlstr = "SELECT * FROM CUM_FAC "
        sqlstr = sqlstr + "  Where Cum_Fac.Id_Obj='PUB'  And   Cum_Fac.Id_Instancia = " + "'" + (Gid_instancia) + "'"
        sqlstr = sqlstr + "  And Cum_Fac.Nro_Plani_pago=" + "'" + (Gcod_planilla) + "'" + ";"
        
        Set cuotas = New ADODB.Recordset
        cuotas.Open sqlstr, cn
        cuotas.MoveFirst
       
  Rem      Gid_rubro = cuotas!Rubro

        tren_transas = cuotas!CUOTA
        
        cuotas.MoveNext
        
        Do While cuotas.EOF = False
        
            tren_transas = tren_transas + ";" + cuotas!CUOTA
            
            cuotas.MoveNext
   
        Loop 'PUBLICIDAD
        
'****************************************************************************
'VEHICULOS
    Case "VEH"
       
       rds.MoveFirst
        
        While rds.EOF = False
        
           Grubro = rds!Rubro
           Gmonto = rds!Monto_Origi
           
            sqlstr = "Update Cum_Fac Set "
            sqlstr = sqlstr + "Cum_Fac.Status='CA', Cum_Fac.Fec_Cancel = getdate()"
            sqlstr = sqlstr + ", Cum_Fac.Monto=" + "'" + (STR(Gmonto)) + "'"
            sqlstr = sqlstr + ", Cum_Fac.Concepto=" + "'" + (rds!Rubro) + "', usuario_rec = '" & Usuario & "'"
            sqlstr = sqlstr + "  Where Cum_Fac.Id_Obj='VEH' And  Cum_Fac.Id_Instancia = " + "'" + (Gid_instancia) + "'"
            sqlstr = sqlstr + "  And Cum_Fac.Nro_Plani_pago=" + "'" + (Gcod_planilla) + "'"
            sqlstr = sqlstr + "  And Cum_Fac.Cuota=" + "'" + (rds!CUOTA) + "'" + ";"
          
            cn.Execute sqlstr, cadena
            
            If cadena = 0 Then

                MsgBox "Cuotas/Facturas  No Se Actualizaron. Seleccione de Nuevo: " + cadena
            
                cn.RollbackTrans
            
                Procesa_Objeto_Especifico = False
            
                Exit Function
            
            End If

            rds.MoveNext
               
        Wend
        
        sqlstr = "SELECT * FROM CUM_FAC "
        sqlstr = sqlstr + "  Where Cum_Fac.Id_Obj='VEH'  And   Cum_Fac.Id_Instancia = " + "'" + (Gid_instancia) + "'"
        sqlstr = sqlstr + "  And Cum_Fac.Nro_Plani_pago=" + "'" + (Gcod_planilla) + "'" + ";"
        
        Set cuotas = New ADODB.Recordset
        cuotas.Open sqlstr, cn
        cuotas.MoveFirst
       
  Rem      Gid_rubro = cuotas!Rubro

        tren_transas = cuotas!CUOTA
        
        cuotas.MoveNext
        
        Do While cuotas.EOF = False
        
            tren_transas = tren_transas + ";" + cuotas!CUOTA
            
            cuotas.MoveNext
   
        Loop 'VEHICULOS
'****************************************************************************
'ADUANA

     Case "ADU"
       
       rds.MoveFirst
        
        While rds.EOF = False
        
           Grubro = rds!Rubro
           Gmonto = rds!Monto_Origi
           
            sqlstr = "Update Cum_Fac Set "
            sqlstr = sqlstr + "Cum_Fac.Status='CA', Cum_Fac.Fec_Cancel = getdate()"
            sqlstr = sqlstr + ", Cum_Fac.Monto=" + "'" + (STR(Gmonto)) + "'"
            sqlstr = sqlstr + ", Cum_Fac.Concepto=" + "'" + (rds!Rubro) + "', usuario_rec = '" & Usuario & "'"
            sqlstr = sqlstr + "  Where Cum_Fac.Id_Obj='ADU' And  Cum_Fac.Id_Instancia = " + "'" + (Gid_instancia) + "'"
            sqlstr = sqlstr + "  And Cum_Fac.Nro_Plani_pago=" + "'" + (Gcod_planilla) + "'"
            sqlstr = sqlstr + "  And Cum_Fac.Cuota=" + "'" + (rds!CUOTA) + "'" + ";"
          
            cn.Execute sqlstr, cadena
            
            If cadena = 0 Then

                MsgBox "Cuotas/Facturas  No Se Actualizaron. Seleccione de Nuevo: " + cadena
            
                cn.RollbackTrans
            
                Procesa_Objeto_Especifico = False
            
                Exit Function
            
            End If

            rds.MoveNext
               
        Wend
        
        sqlstr = "SELECT * FROM CUM_FAC "
        sqlstr = sqlstr + "  Where Cum_Fac.Id_Obj='ADU'  And   Cum_Fac.Id_Instancia = " + "'" + (Gid_instancia) + "'"
        sqlstr = sqlstr + "  And Cum_Fac.Nro_Plani_pago=" + "'" + (Gcod_planilla) + "'" + ";"
        
        Set cuotas = New ADODB.Recordset
        cuotas.Open sqlstr, cn
        cuotas.MoveFirst
       
  Rem      Gid_rubro = cuotas!Rubro

        tren_transas = cuotas!CUOTA
        
        cuotas.MoveNext
        
        Do While cuotas.EOF = False
        
            tren_transas = tren_transas + ";" + cuotas!CUOTA
            
            cuotas.MoveNext
   
        Loop 'ADUANA
'****************************************************************************
'APUESTAS LICITAS

     Case "APU"
       
       rds.MoveFirst
        
        While rds.EOF = False
        
           Grubro = rds!Rubro
           Gmonto = rds!Monto_Origi
           
            sqlstr = "Update Cum_Fac Set "
            sqlstr = sqlstr + "Cum_Fac.Status='CA', Cum_Fac.Fec_Cancel = getdate()"
            sqlstr = sqlstr + ", Cum_Fac.Monto=" + "'" + (STR(Gmonto)) + "'"
            sqlstr = sqlstr + ", Cum_Fac.Concepto=" + "'" + (rds!Rubro) + "', usuario_rec = '" & Usuario & "'"
            sqlstr = sqlstr + "  Where Cum_Fac.Id_Obj='APU' And  Cum_Fac.Id_Instancia = " + "'" + (Gid_instancia) + "'"
            sqlstr = sqlstr + "  And Cum_Fac.Nro_Plani_pago=" + "'" + (Gcod_planilla) + "'"
            sqlstr = sqlstr + "  And Cum_Fac.Cuota=" + "'" + (rds!CUOTA) + "'" + ";"
          
            cn.Execute sqlstr, cadena
            
            If cadena = 0 Then

                MsgBox "Cuotas/Facturas  No Se Actualizaron. Seleccione de Nuevo: " + cadena
            
                cn.RollbackTrans
            
                Procesa_Objeto_Especifico = False
            
                Exit Function
            
            End If

            rds.MoveNext
               
        Wend
        
        sqlstr = "SELECT * FROM CUM_FAC "
        sqlstr = sqlstr + "  Where Cum_Fac.Id_Obj='APU'  And   Cum_Fac.Id_Instancia = " + "'" + (Gid_instancia) + "'"
        sqlstr = sqlstr + "  And Cum_Fac.Nro_Plani_pago=" + "'" + (Gcod_planilla) + "'" + ";"
        
        Set cuotas = New ADODB.Recordset
        cuotas.Open sqlstr, cn
        cuotas.MoveFirst
       
  Rem      Gid_rubro = cuotas!Rubro

        tren_transas = cuotas!CUOTA
        
        cuotas.MoveNext
        
        Do While cuotas.EOF = False
        
            tren_transas = tren_transas + ";" + cuotas!CUOTA
            
            cuotas.MoveNext
   
        Loop 'APUESTAS LICITAS

'*****************************************************************************
 Case Else
    
            MsgBox "Objeto/Genero en Proceso No Identificado: " + Gid_obj + " .Nro_Planilla: " + Gcod_planilla
        
            'Wrk.Rollback
            cn.RollbackTrans
        
            Procesa_Objeto_Especifico = False
        
            MsgBox "!!! Operación Abortada.Llame al Administrador del Sistema.Gracias."
        
        
            Exit Function
    
End Select

Procesa_Objeto_Especifico = True

End Function


Private Function PROCESA_OBJETO_GENERICO() As Boolean

Rem Crea el record para el Objeto_Genérico_Recauda : A Quien se le Recauda.
Rem El Contribuyente Generico objeto de la Recuadacion Generica.

Dim sqlstr As String

sqlstr = "Insert Into Objetos_Genericos_Recauda "
sqlstr = sqlstr + " (id_objeto,id_instancia,Id_contri,Nombre01) "
sqlstr = sqlstr + " Values ("
sqlstr = sqlstr + "'" + (Gid_obj) + "'," + "'" + (Gid_instancia) + "'," + "'" + (Gid_Contri) + "'," + "'" + (Grazon_social) + "')" + ";"

cn.Execute sqlstr
'bds.Execute (sqlstr)

Rem  Genera Cuotas/Porciones Genericas : Que y Cuanto se le Recauda.

Dim renglones As Integer
Dim i As Byte
Dim pos As Byte
Dim scuota As String
Dim smonto As Double
Dim porciones As String
Dim token As String
Dim Fecha As Date

ABRIR_CUM_FAC ' ABRE CUM_FAC PARA EDICIÓN PESIMISTA


'Set rdsfac = bds.OpenRecordset("Cum_Fac", dbOpenDynaset, dbSeeChanges, dbPessimistic)

Rem SQL REQUIERE LA FECHA EN ESTE FORMATO :mm-dd-aa

'FECHA = Str(Month(Gfec_pago)) + "/" + Str(Day(Gfec_pago)) + "/" + Str(Year(Gfec_pago))
Fecha = Format(Date, "mm,dd,yy")
porciones = Gid_instancia   ' Colección : <porción,monto>+.
 
If Grenglon > 1 Then
  
    renglones = Grenglon

    For i = 1 To renglones
    
        If Len(porciones) = 0 Then
    
            MsgBox "Porciones/Cuotas Genericas agotadas."
        
            Exit For
        
        End If
    
        pos = InStr(1, porciones, ",", 1)    ' Ubica la  (,)  separadora de cuota y monto
    
        scuota = Trim(Mid(porciones, 1, pos - 1))   ' Extrae la cuota : 200101
       
        porciones = Mid(porciones, pos + 1)  ' Compacta porciones: desde comienza el monto
        
        pos = InStr(1, porciones, ";", 1)    ' Ubica el (;) 15000;2001-02,15000.99;
    
        token = Mid(porciones, 1, pos - 1)   ' Extrae el monto
    
        smonto = CDbl(token)                 ' Convierte a doble numero
    
        porciones = Mid(porciones, pos + 1)  ' Getnetx token desde el (;)
        
        rdsfac.AddNew
                
                rdsfac!usuario_rec = Usuario
         
                rdsfac!ID_OBJ = Gid_obj
                
                rdsfac!Id_Instancia = Gid_instancia
                
                rdsfac!CUOTA = Grenglon
                
                'rdsfac!NRO_PLANI_PAGO = Tex_Planilla ' ****************falta***************
                
                rdsfac!Concepto = Grubro
                
                rdsfac!monto = Gmonto
                
                rdsfac!FEC_EMI = Date
                
                rdsfac!FEC_VIG = Date
                
                rdsfac!FEC_CANCEL = Date
                
                rdsfac!STATUS = "CA"
                
                rdsfac!voucher = Me.txt_Voucher
                
                rdsfac!cod_recauda = "99"
                
        rdsfac.Update
         

    Next i
  
Else
         
         rdsfac.AddNew
         
                rdsfac!usuario_rec = Usuario
         
                rdsfac!ID_OBJ = Gid_obj
                
                rdsfac!Id_Instancia = Mid(Gid_instancia, 1, 12)
                
                'rdsfac!Cuota = Me.Tex_Planilla '**falta***************
                
                'rdsfac!NRO_PLANI_PAGO = Me.Tex_Planilla ' *************falta***********
                
                rdsfac!Concepto = Grubro
                
                rdsfac!monto = Gmonto
                
                rdsfac!AÑO = Trim(Year(Date))
                
                rdsfac!FEC_EMI = Date
                
                rdsfac!FEC_VIG = Date
                
                rdsfac!FEC_CANCEL = Date
                
                rdsfac!STATUS = "CA"
                
                rdsfac!cod_recauda = "99"
                rdsfac!Select = 0
                
         rdsfac.Update
         
        
End If


rdsfac.Close

'Com_Reinicio.Enabled = False ' **************falta*****************

PROCESA_OBJETO_GENERICO = True


End Function

Private Sub txt_Voucher_Validate(Cancel As Boolean)
Dim i As Integer
Abandono = False
For i = 1 To 2
If Me.CommandButton(i).FontBold Then
    Cancel = False
    Exit Sub
End If
Next i

        If Len(txt_Voucher) = 0 Or txt_Voucher = "" Or IsNull(txt_Voucher) Then
            MsgBox "Debe Suministrar el Nro. del Voucher/Planilla de Déposito.Gracias."
            Cancel = True
            Abandono = True
        End If

        If Not control Then
            Me.lbl_Voucher.Caption = "Ingrese Nro. Voucher"
            control = Not control
            Cancel = True
        End If

        If control And Not Abandono Then

            If N_Voucher = "" Then
            
                N_Voucher = Me.txt_Voucher
                Me.lbl_Voucher.Caption = "Reingrese N° Voucher"
                Me.txt_Voucher = ""
                Cancel = True 'Me.txt_Voucher.SetFocus
                
             Else
                If N_Voucher <> Me.txt_Voucher.Text Then
                    MsgBox "Error en el Nro. del Voucher ", vbCritical
                    N_Voucher = ""
                    Me.lbl_Voucher.Caption = "Ingrese Nro. Voucher"
                    Me.txt_Voucher.Text = ""
                    Cancel = True 'Me.txt_Voucher.SetFocus
                Else
                    Me.lbl_Voucher.Caption = "Ingrese Nro. Voucher"
                    Cancel = False 'Me.txt_MontoV.SetFocus
                    control = Not control
                    'chequeo_V  -----AQUI LLAMADA A VERIFICAR SI VOUCHER EXISTE----
                End If
            End If
        End If

End Sub

