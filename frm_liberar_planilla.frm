VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_liberar_planilla 
   Caption         =   "Liberar una Factura"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8925
   Icon            =   "frm_liberar_planilla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   8925
   Begin MSAdodcLib.Adodc obj 
      Height          =   495
      Left            =   5040
      Top             =   3360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
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
      RecordSource    =   "TAB_ID_OBJ"
      Caption         =   "obj"
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
   Begin VB.CommandButton cmd_eliminar 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   600
      TabIndex        =   7
      ToolTipText     =   "Elimina una Factura"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Introduzca el Número de planilla de Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   8415
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         DataField       =   "ULT"
         DataSource      =   "alc_obj_liq"
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txt_liquidado 
         DataField       =   "Nro_Plani_Pago"
         DataSource      =   "liquidado"
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txt_id_pedidos 
         DataField       =   "NRO_PLANI_PAGO"
         DataSource      =   "cum_fac"
         Height          =   285
         Left            =   360
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo DCombo_idobj 
         Bindings        =   "frm_liberar_planilla.frx":08CA
         Height          =   360
         Left            =   4200
         TabIndex        =   1
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "Descripcion"
         BoundColumn     =   "Id_Obj"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txt_planilla 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   3735
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "Cerrar"
         Height          =   735
         Left            =   6600
         TabIndex        =   3
         ToolTipText     =   "Cerrar la Pantalla"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmd_liberar 
         Caption         =   "Liberar"
         Enabled         =   0   'False
         Height          =   735
         Left            =   5040
         TabIndex        =   2
         ToolTipText     =   "Libera una Factura"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "La ultima planilla de pago es:"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label lbl_informa_liq 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Lbl_informacion 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   2775
      End
   End
   Begin MSAdodcLib.Adodc cum_fac 
      Height          =   375
      Left            =   480
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "CUM_FAC"
      Caption         =   "cum_fac"
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
   Begin MSAdodcLib.Adodc liquidado 
      Height          =   375
      Left            =   2760
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "ALC_OBJ_LIQS"
      Caption         =   "liquidado"
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
   Begin MSAdodcLib.Adodc alc_obj_liq 
      Height          =   375
      Left            =   5040
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "select MAX(nro_plani_pago) AS ULT from ALC_OBJ_LIQS "
      Caption         =   "alc_obj_liq"
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
Attribute VB_Name = "frm_liberar_planilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = True
Me.cmd_liberar.FontBold = False

End Sub

Private Sub cmd_eliminar_Click()
On Error GoTo ControlError
Dim Cod
Cod = InputBox("Suministre la clave para eliminar la planilla", "ALCASIS C.A.")

If Cod = "123" Then
    
    pedidos.Recordset.MoveFirst
    
    strquery = "id_pedido = '" & Me.txt_planilla.Text & "'"

    pedidos.Recordset.Find strquery
    
    If pedidos.Recordset.EOF Then
    
            MsgBox "Nºde Planilla suministrada no encontrada, por favor verifique ", vbInformation, "ALCASIS C.A."
            
            Me.Lbl_informacion.Caption = "Planilla No Encontrada"
            
            Exit Sub
                    
    Else
            
            pedidos.Recordset.Delete
            
            Me.Lbl_informacion.Caption = "Planilla Eliminada"
            
            'Liberando en liquidacion
            liquidado.Recordset.MoveFirst
            
            strquery = "id_pedido = '" & Me.txt_planilla.Text & "'"
        
            liquidado.Recordset.Find strquery
            
            If liquidado.Recordset.EOF Then
            
                    MsgBox "Nºde Planilla suministrada no encontrada, por favor verifique ", vbInformation, "ALCASIS C.A."
                    
                    Me.Lbl_informacion.Caption = "Planilla No Encontrada"
                    
                    Exit Sub
                            
            Else
                liquidado.Recordset.Delete
                lbl_informa_liq.Caption = "Planilla borrada de liquidación"
            End If
            
            liquidado.Recordset.Close

    End If
    
    pedidos.Recordset.Close
End If
Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS C.A.")
        
    End Select

End Sub

Private Sub cmd_liberar_Click()
On Error GoTo ControlError

Dim NRO_PAT
Dim nohayplanilla As Boolean
Dim strquery

If DCombo_idobj.Text = "" Then

    MsgBox "Seleccione el rubro (ADU,INM,VEH,PIC,PUB)", vbInformation
    Exit Sub
    
End If


    nohayplanilla = True
    
    cum_fac.Recordset.MoveFirst
    
    strquery = "NRO_PLANI_PAGO = '" & Me.txt_planilla.Text & "'"
    
    cum_fac.Recordset.Find strquery
    
    While Not cum_fac.Recordset.EOF
    
        NRO_PAT = cum_fac.Recordset!Id_Instancia
    
        If cum_fac.Recordset!STATUS = "VI" Then
        
            If cum_fac.Recordset!ID_OBJ = Me.DCombo_idobj.BoundText Then
                If Not IsNull(cum_fac.Recordset!NRO_PLANI_PAGO) Then
                    cum_fac.Recordset!STATUS = "VI"
                     
                    cum_fac.Recordset!FEC_CANCEL = Null
                    
                    cum_fac.Recordset!NRO_PLANI_PAGO = Null
                    
                    cum_fac.Recordset.Update
                    
                    Me.Lbl_informacion.Caption = "Planilla Liberada"
                    
                    cmd_liberar.Enabled = False
                    Me.cmd_cerrar.SetFocus
                    nohayplanilla = False
                    
                End If
            End If
        Else
            If nohayplanilla = True Then
                MsgBox "Verifique el rubro (INM,VEH,PIC,PUB,ADU), si las cuotas estan pagadas no se pueden liberar", vbInformation
            End If
        End If
        cum_fac.Recordset.MoveNext
    Wend
    
     liquidado.Recordset.MoveFirst
                
    liquidado.Recordset.Find strquery
    
    While Not liquidado.Recordset.EOF
        liquidado.Recordset.Delete
        liquidado.Recordset.MoveNext
    Wend
    
'    cum_fac.Refresh
    
'    cum_fac.ConnectionString = "DSN=SIAGEP"
'
'    cum_fac.CommandType = adCmdText
'
'    strquery = "SELECT * From cum_fac WHERE (Id_Instancia='" & NRO_PAT & "' and STATUS = 'VI' and NRO_PLANI_PAGO = '" & Me.txt_planilla.Text & "' and id_obj='" & DCombo_idobj.BoundText & "')"
'
'    cum_fac.RecordSource = strquery
'
'    While Not cum_fac.Recordset.EOF
'            nohayplanilla = False
'
'            cum_fac.Recordset!STATUS = "VI"
'
'            cum_fac.Recordset!FEC_CANCEL = Null
'
'            cum_fac.Recordset!NRO_PLANI_PAGO = Null
''            cum_fac.Recordset!Id_Instancia = Null
'            cum_fac.Recordset.Update
'
'            cum_fac.Recordset.MoveNext
'
'            Me.Lbl_informacion.Caption = "Planilla Liberada"
'
'    Wend
    
    cum_fac.Recordset.Close
    
'    If nohayplanilla = True Then
'
'        MsgBox "Verifique el rubro (INM,VEH,PIC,PUB), si las cuotas estan pagadas no se pueden liberar", vbInformation
'    End If
'    liquidado.ConnectionString = "DSN=SIAGEP"
'
'    liquidado.CommandType = adCmdText
'
'    strquery = "SELECT * From ALC_OBJ_LIQS WHERE (status <> 'CA' and Nro_Plani_Pago = '" & Me.txt_planilla.Text & "' and Id_Objeto='" & DCombo_idobj.BoundText & "')"
'
'    liquidado.RecordSource = strquery
'
'    While Not liquidado.Recordset.EOF
'
'                liquidado.Recordset.Delete
'                liquidado.Recordset.MoveNext
'                lbl_informa_liq.Caption = "Planilla borrada de liquidación"
'
'    Wend
'
'    liquidado.Recordset.Close


    
Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS C.A.")
        Case 3704
            cum_fac.Refresh
    End Select
End Sub

Private Sub cmd_liberar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_liberar.FontBold = True
End Sub

Private Sub DCombo_idobj_Click(area As Integer)
cmd_liberar.Enabled = True
End Sub

Private Sub Form_Load()
Me.Height = 2700
Me.Width = 6700
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_liberar.FontBold = False

End Sub

Private Sub Form_Resize()
frm_liberar_planilla.Width = 9045
frm_liberar_planilla.Height = 3145
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_liberar.FontBold = False

End Sub

Private Sub txt_planilla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub
