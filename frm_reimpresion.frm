VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_reimpresion 
   Caption         =   "Reimpresión"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command 
      Caption         =   "Cerrar"
      Height          =   615
      Index           =   2
      Left            =   4440
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   3120
      TabIndex        =   7
      Top             =   1440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   51380225
      CurrentDate     =   38061
   End
   Begin VB.TextBox TextBox 
      DataField       =   "NRO_PAT"
      DataSource      =   "Establecimientos"
      Height          =   315
      Index           =   0
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.ListBox Lista_tipo_factura 
      Height          =   1425
      ItemData        =   "frm_reimpresion.frx":0000
      Left            =   240
      List            =   "frm_reimpresion.frx":0016
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command 
      Caption         =   "Buscar Nº de Planilla"
      Height          =   615
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Caption         =   "Vista Previa"
      Height          =   615
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CheckBox Check 
      Caption         =   "Impresión Tiquera"
      Height          =   615
      Index           =   2
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "Fecha"
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
      Index           =   2
      Left            =   3120
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label 
      Caption         =   "Tipo de Factura"
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
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label 
      Caption         =   "Nº de Planilla"
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
      Index           =   0
      Left            =   3120
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frm_reimpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim existe, Imp_Tiquera As Boolean

Private Sub Check_Click(Index As Integer)

If Check(Index).Value = 0 Then
    Me.Check(Index).BackColor = vbButtonFace
Else
    Me.Check(Index).BackColor = vbRed
End If
    
Select Case Index
    Case 2
        Imp_Tiquera = Check(Index).Value

End Select

End Sub

Private Sub Command_Click(Index As Integer)
Select Case Index
    Case 0
        Imp_auto = False
        Call Vista_P
    Case 1
        frm_buscar_n_voucher.Show 1, Me
    Case 2
        Unload Me
    
End Select
End Sub

Private Sub Vista_P()
On Error GoTo control_de_errores

If IsNull(Me.TextBox(0)) Then
    MsgBox "Nº de Planilla no puede estar vacío", vbInformation, "ALCASIS"
    Exit Sub
End If
Dim cadena, filtro, copias As String
Dim i As Integer
Dim Fecha, sqlstr, fecha_reimp As String
Fecha = Format(DTPicker1.Object, "long date")
fecha_reimp = Date & " " & Time
cadena = Lista_tipo_factura.ListIndex
Gcod_planilla = Me.TextBox(0)
filtro = "Nro_Plani_Pago = '" & Me.TextBox(0) & "'"
nro_de_planilla = filtro
sqlstr = "insert into control_de_reimpresión (NRO_PLANI_PAGO,USUARIO,FECHA_REIMPRESIÓN) VALUES('" & Me.TextBox(0) & "','" & user_name & "','" & fecha_reimp & "')"
    Select Case cadena
            
           Case "0"
               verificar_veh
               If existe Then
                    Unload Me
                        If Imp_Tiquera Then
                            rpt_cr_alc_rec_tiq_veh.Show
                        Else
                            rpt_alc_recaudacion_veh.Show
                        End If
                    cn.Execute sqlstr
               Else
                    MsgBox "Nro. de Planilla no existe en Específica de Vehículos", vbInformation, "ALCASIS"
                    Exit Sub
               End If
           Case "1"
               verificar_inm
               If existe Then
                    Unload Me
                        If Imp_Tiquera Then
                            rpt_cr_alc_rec_tiq_inm.Show
                        Else
                            rpt_alc_recaudacion_inm.Show
                        End If
                    cn.Execute sqlstr
               Else
                    MsgBox "Nro. de Planilla no existe en Específica de Inmuebles", vbInformation, "ALCASIS"
                    Exit Sub
               End If
           
           Case "2"
               verificar_pic
               If existe Then
                    Unload Me
                    If Imp_Tiquera Then
                        rpt_cr_alc_rec_tiq_pic.Show
                    Else
                        rpt_alc_recaudacion_pic.Show
                    End If
                    cn.Execute sqlstr
               Else
                    MsgBox "Nro. de Planilla no existe en Específica de PIC", vbInformation, "ALCASIS"
                    Exit Sub
               End If
               
           Case "3"
              verificar_pub
               If existe Then
                    Unload Me
                        If Imp_Tiquera Then
                            rpt_cr_alc_rec_tiq_pub.Show
                        Else
                            rpt_alc_recaudacion_pub.Show
                        End If
                    cn.Execute sqlstr
               Else
                    MsgBox "Nro. de Planilla no existe en Específica de Publicidad", vbInformation, "ALCASIS"
                    Exit Sub
               End If
           
           Case "4"
              verificar_gen
               If existe Then
                   Unload Me
                    If Imp_Tiquera Then
                        rpt_cr_alc_rec_tiq_gen.Show
                    Else
                        rpt_alc_recuadacion_gen.Show
                    End If
                    cn.Execute sqlstr
               Else
                    MsgBox "Nro. de Planilla no existe en Genérica", vbInformation, "ALCASIS"
                    Exit Sub
               End If
           Case "5"
               verificar_apu
               If existe Then
                   Unload Me
                        If Imp_Tiquera Then
                            rpt_cr_alc_rec_tiq_apu.Show
                        Else
                            rpt_alc_recaudacion_apu.Show
                        End If
                    cn.Execute sqlstr
               Else
                    MsgBox "Nro. de Planilla no existe en Genérica", vbInformation, "ALCASIS"
                    Exit Sub
               End If
    Case Else
        MsgBox "Debe seleccionar el tipo de factura", vbInformation + vbOKOnly, "ALCASIS"
    End Select
Exit Sub
control_de_errores:
    MsgBox Err.Description, vbInformation, "ALCASIS"

End Sub

Private Sub verificar_veh()
Dim rst As ADODB.Recordset
Dim sqlstr As String
Set rst = New ADODB.Recordset
sqlstr = "SELECT ALC_OBJ_LIQS.Nro_Plani_Pago FROM ALC_OBJ_LIQS INNER JOIN " _
       & "CUM_FAC ON ALC_OBJ_LIQS.Id_Instancia = CUM_FAC.ID_INSTANCIA AND " _
       & "ALC_OBJ_LIQS.Id_Objeto = CUM_FAC.ID_OBJ AND ALC_OBJ_LIQS.Cuota = " _
       & "CUM_FAC.CUOTA INNER JOIN TAB_TASAS ON dbo.CUM_FAC.CONCEPTO = " _
       & "TAB_TASAS.CONCEPTO INNER JOIN VEHICULOS ON CUM_FAC.ID_INSTANCIA = " _
       & "VEHICULOS.PLACA WHERE CUM_FAC.ID_OBJ = 'VEH' AND " _
       & "ALC_OBJ_LIQS.Nro_Plani_Pago = '" & Gcod_planilla & "'"
'actualizar_cn
rst.Open sqlstr, cn
If rst.EOF Then
    existe = False
Else
    existe = True
End If
rst.Close
End Sub


Private Sub verificar_inm()
Dim rst As ADODB.Recordset
Dim sqlstr As String
Set rst = New ADODB.Recordset
sqlstr = "SELECT ALC_OBJ_LIQS.Nro_Plani_Pago FROM ALC_OBJ_LIQS INNER JOIN " _
       & "CUM_FAC ON ALC_OBJ_LIQS.Id_Instancia = CUM_FAC.ID_INSTANCIA AND " _
       & "ALC_OBJ_LIQS.Id_Objeto = CUM_FAC.ID_OBJ AND ALC_OBJ_LIQS.Cuota = " _
       & "CUM_FAC.CUOTA INNER JOIN INMUEBLES ON CUM_FAC.ID_INSTANCIA = " _
       & "INMUEBLES.COD_CATA INNER JOIN TAB_TASAS ON CUM_FAC.CONCEPTO = " _
       & "TAB_TASAS.CONCEPTO WHERE dbo.CUM_FAC.ID_OBJ = 'INM' AND " _
       & "ALC_OBJ_LIQS.Nro_Plani_Pago = '" & Gcod_planilla & "'"

rst.Open sqlstr, cn
If rst.EOF Then
    existe = False
Else
    existe = True
End If
rst.Close
End Sub

Private Sub verificar_pic()
Dim rst As ADODB.Recordset
Dim sqlstr As String
Set rst = New ADODB.Recordset
sqlstr = "SELECT ALC_OBJ_LIQS.Nro_Plani_Pago FROM ALC_OBJ_LIQS INNER JOIN " _
       & "CUM_ESTABLECIMIENTOS INNER JOIN CUM_FAC ON CUM_ESTABLECIMIENTOS.NRO_PAT " _
       & "=CUM_FAC.ID_INSTANCIA ON ALC_OBJ_LIQS.Id_Instancia = dbo.CUM_FAC.ID_INSTANCIA " _
       & "AND ALC_OBJ_LIQS.Id_Objeto = CUM_FAC.ID_OBJ AND ALC_OBJ_LIQS.Cuota = " _
       & "CUM_FAC.CUOTA INNER JOIN TAB_TASAS ON CUM_FAC.CONCEPTO = TAB_TASAS.CONCEPTO " _
       & "WHERE CUM_FAC.STATUS = 'CA' AND CUM_FAC.ID_OBJ = 'PIC' AND " _
       & "ALC_OBJ_LIQS.Nro_Plani_Pago = '" & Gcod_planilla & "'"
'Call actualizar_cn
rst.Open sqlstr, cn
If rst.EOF Then
    existe = False
Else
    existe = True
End If
rst.Close
End Sub

Private Sub verificar_apu()
Dim rst As ADODB.Recordset
Dim sqlstr As String
Set rst = New ADODB.Recordset
sqlstr = "SELECT ALC_OBJ_LIQS.Nro_Plani_Pago FROM ALC_OBJ_LIQS INNER JOIN " _
       & "CUM_ESTABLECIMIENTOS INNER JOIN CUM_FAC ON CUM_ESTABLECIMIENTOS.NRO_PAT " _
       & "=CUM_FAC.ID_INSTANCIA ON ALC_OBJ_LIQS.Id_Instancia = dbo.CUM_FAC.ID_INSTANCIA " _
       & "AND ALC_OBJ_LIQS.Id_Objeto = CUM_FAC.ID_OBJ AND ALC_OBJ_LIQS.Cuota = " _
       & "CUM_FAC.CUOTA INNER JOIN TAB_TASAS ON CUM_FAC.CONCEPTO = TAB_TASAS.CONCEPTO " _
       & "WHERE CUM_FAC.STATUS = 'CA' AND CUM_FAC.ID_OBJ = 'APU' AND " _
       & "ALC_OBJ_LIQS.Nro_Plani_Pago = '" & Gcod_planilla & "'"
'actualizar_cn
'rst.Open sqlstr, cn
If rst.EOF Then
    existe = False
Else
    existe = True
End If
rst.Close
End Sub

Private Sub verificar_pub()
Dim rst As ADODB.Recordset
Dim sqlstr As String
Set rst = New ADODB.Recordset
sqlstr = "SELECT ALC_OBJ_LIQS.Nro_Plani_Pago FROM CUM_ESTABLECIMIENTOS INNER JOIN " _
       & "CUM_FAC ON CUM_ESTABLECIMIENTOS.NRO_PAT = CUM_FAC.ID_INSTANCIA INNER JOIN " _
       & "TAB_TASAS ON CUM_FAC.CONCEPTO = TAB_TASAS.CONCEPTO INNER JOIN CUM_PUBLICIDADES " _
       & "ON CUM_FAC.ID_ASO = CUM_PUBLICIDADES.ID_PUB INNER JOIN ALC_OBJ_LIQS " _
       & "ON CUM_FAC.ID_ASO = ALC_OBJ_LIQS.Id_Aso AND CUM_FAC.NRO_PLANI_PAGO =" _
       & "Alc_Obj_Liqs.NRO_PLANI_PAGO And CUM_FAC.CUOTA = Alc_Obj_Liqs.CUOTA " _
       & "WHERE CUM_FAC.STATUS = 'CA' AND CUM_FAC.ID_OBJ = 'PUB' AND ALC_OBJ_LIQS.Nro_Plani_Pago = '" & Gcod_planilla & "'"
'actualizar_cn
'rst.Open sqlstr, cn
If rst.EOF Then
    existe = False
Else
    existe = True
End If
rst.Close
End Sub

Private Sub verificar_gen()
Dim rst As ADODB.Recordset
Dim sqlstr As String
Set rst = New ADODB.Recordset
sqlstr = "SELECT Nro_Plani_Pago FROM Alc_Obj_Liqs WHERE " _
       & "Nro_Plani_Pago = '" & Gcod_planilla & "'"
'actualizar_cn
rst.Open sqlstr, cn
If rst.EOF Then
    existe = False
Else
    existe = True
End If
rst.Close
End Sub


Private Sub Form_Load()
Me.DTPicker1.Value = Date
Me.Check(2).Value = 1
End Sub

Private Sub Lista_tipo_factura_Click()
Me.TextBox(0).SetFocus
End Sub

Private Sub TextBox_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub
