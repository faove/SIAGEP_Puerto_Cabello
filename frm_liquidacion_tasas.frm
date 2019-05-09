VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_liquidacion_tasas 
   Caption         =   "Tasas y Otros Tributos"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2400
      Top             =   6960
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
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
      RecordSource    =   "Get_Liq_tasas_Sfrm"
      Caption         =   "Adodc2"
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
   Begin VB.TextBox Text1 
      DataField       =   "Nro_Plani_Pago"
      DataSource      =   "Adodc3"
      Height          =   285
      Left            =   2640
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   6600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   120
      Top             =   6600
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Connect         =   "DSN=SIAGEP"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "SIAGEP"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   " SELECT * FROM ALC_OBJ_LIQS WHERE ALC_OBJ_LIQS.Nro_Plani_Pago = '0'"
      Caption         =   "Adodc3"
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
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   3240
      TabIndex        =   12
      Top             =   360
      Width           =   8295
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "TASAS Y OTROS TRIBUTOS"
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
         Left            =   2760
         TabIndex        =   13
         Top             =   0
         Width           =   5535
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Edición"
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
         TabIndex        =   14
         Top             =   360
         Width           =   6855
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5895
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   11175
      Begin VB.TextBox Id_Objeto 
         DataField       =   "Id_Objeto"
         DataSource      =   "Adodc3"
         Height          =   285
         Left            =   2400
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   5040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   975
         Left            =   5760
         TabIndex        =   25
         Top             =   4920
         Width           =   5415
         Begin VB.CommandButton Command 
            Caption         =   "Cerrar"
            Height          =   615
            Index           =   3
            Left            =   4080
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command 
            Caption         =   "Cancelar"
            Height          =   615
            Index           =   4
            Left            =   2760
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command 
            Caption         =   "Aceptar Liquidación"
            Height          =   615
            Index           =   5
            Left            =   1440
            TabIndex        =   8
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command 
            Caption         =   "Próxima"
            Enabled         =   0   'False
            Height          =   615
            Index           =   6
            Left            =   120
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frm_liquidacion_tasas.frx":0000
         Height          =   4095
         Left            =   0
         TabIndex        =   0
         Top             =   720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7223
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
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
         ColumnCount     =   1
         BeginProperty Column00 
            DataField       =   "DESCRIPCION"
            Caption         =   "                                           DESCRIPCION"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   5520,189
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame3 
         Caption         =   "Items a Liquidar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   6120
         TabIndex        =   18
         Top             =   2640
         Width           =   5055
         Begin VB.CommandButton Command 
            Caption         =   "Aceptar Items"
            Height          =   615
            Index           =   0
            Left            =   3240
            TabIndex        =   7
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox TextBox 
            DataField       =   "NRO_PAT"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Establecimientos"
            Height          =   315
            Index           =   6
            Left            =   3360
            MaxLength       =   12
            TabIndex        =   6
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox TextBox 
            DataField       =   "NRO_PAT"
            DataSource      =   "Establecimientos"
            Height          =   315
            Index           =   5
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   5
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox TextBox 
            DataField       =   "NRO_PAT"
            DataSource      =   "Establecimientos"
            Height          =   315
            Index           =   4
            Left            =   120
            MaxLength       =   10
            TabIndex        =   4
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton Command 
            Caption         =   "Siguiente"
            Height          =   615
            Index           =   1
            Left            =   1920
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton Command 
            Caption         =   "Reinicio"
            Height          =   615
            Index           =   2
            Left            =   600
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label 
            Caption         =   "Cantidad / Monto"
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
            Index           =   6
            Left            =   3360
            TabIndex        =   21
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label 
            Caption         =   "Porciones Pago"
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
            Index           =   5
            Left            =   1800
            TabIndex        =   20
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label 
            Caption         =   "Cantidad / Monto"
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
            Index           =   4
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.TextBox TextBox 
         DataField       =   "NRO_PAT"
         DataSource      =   "Establecimientos"
         Height          =   1155
         Index           =   3
         Left            =   6120
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1320
         Width           =   5055
      End
      Begin VB.TextBox TextBox 
         DataField       =   "NRO_PAT"
         DataSource      =   "Establecimientos"
         Height          =   675
         Index           =   2
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   5055
      End
      Begin VB.TextBox TextBox 
         DataField       =   "NRO_PAT"
         DataSource      =   "Establecimientos"
         Height          =   315
         Index           =   1
         Left            =   1920
         MaxLength       =   35
         TabIndex        =   2
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox TextBox 
         DataField       =   "NRO_PAT"
         DataSource      =   "Establecimientos"
         Height          =   315
         Index           =   0
         Left            =   0
         MaxLength       =   14
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label 
         Caption         =   "Descripción"
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
         Index           =   3
         Left            =   6120
         TabIndex        =   17
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label Label 
         Caption         =   "Concepto"
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
         Left            =   6120
         TabIndex        =   16
         Top             =   0
         Width           =   4095
      End
      Begin VB.Label Label 
         Caption         =   "Razón Social / Denominación Comercial"
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
         Left            =   1920
         TabIndex        =   11
         Top             =   0
         Width           =   3735
      End
      Begin VB.Label Label 
         Caption         =   "Contribuyente"
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
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frm_liquidacion_tasas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Items(12, 12)
Dim Tex_Tot_Monto As Single
Dim CONTADOR As Byte, J As Byte
Dim Tex_Items As String
Dim BS1 As Single, BS2 As Single
Dim Nro_P As String
Dim RUB As String



Private Sub Command_Click(Index As Integer)
Select Case Index
    Case 0
        Call Aceptar_Item
    Case 2
        Call Reinicio
    Case 3
        Unload Me
    Case 5
        Call Aceptar
    Case 6
        Call Proxima

End Select
End Sub

Private Sub DataGrid1_Click()
Call Reinicio
With Me.Adodc2.Recordset
Me.TextBox(2).Text = !Concepto & " " & !Descripcion
Me.TextBox(3).Text = !ID_OBJ & " " & !Concepto & " " & !Descripcion
Gid_obj = !ID_OBJ
End With
RUB = Me.Adodc2.Recordset!Concepto

Me.TextBox(0).SetFocus
End Sub

Private Sub Form_Load()

   
   
    Me.TextBox(0).Text = FGID_CONTRI()
    
'    Nro_P = FGNRO_LIQ()
'    Me.Label2.Caption = "Nro. de Planilla: " & Nro_P
    
    
    J = 1
    
    CONTADOR = 0
    
    Tex_Cantidad = 0
    
    Tex_Item = ""
    
    Me.TextBox(5) = 0
    Tex_Items = ""
    

End Sub
Private Sub Aceptar_Item()
    
    Dim i As Byte
    Dim Valor As String
    Dim CUOTA As String
    Dim num As Integer
    
    If (Len(Me.TextBox(1).Text) = 0) Or (Me.TextBox(1).Text = "") Then

        MsgBox "Debe Suministrar Razón Social / Nombre Contribuyente.Gracias.", vbCritical, "Control de Operaciones de Liquidación"
        
    
        Me.TextBox(1).SetFocus
    
        Exit Sub
    
    Exit Sub
    
        
    End If
    
    
    If Me.TextBox(4) = 0 Then
        
        MsgBox "Cantidad de Items Procesados Es Cero(0)."
        
        Exit Sub
        
    
    End If
    
    Me.Command(5).Enabled = True
    
    Me.Command(5).SetFocus
        
    If Me.TextBox(4).Text = "1" Then
    
        Tex_Items = Me.TextBox(5)
        
        Exit Sub
    
    End If
        
        
    Tex_Items = ""
    
    num = CInt(Me.TextBox(4).Text)
        
    For i = 1 To num - 1
        
        If IsNull(Items(i, 1)) Then
        
            Exit For
                        
        End If
        
        CUOTA = Items(i, 1)
        
        Valor = STR(Items(i, 2))
        
        Tex_Items = Tex_Items + CUOTA + " : " + Valor + " ; "
        
    Next
End Sub
Private Sub Reinicio()
        
        CONTADOR = 0
        Me.TextBox(0) = ""
        Me.TextBox(1) = ""
        Me.TextBox(5) = ""
        Me.TextBox(6) = ""
        J = 1
        Tex_Tot_Monto = 0
        Me.TextBox(4) = "1"
        Me.TextBox(2).Text = ""
    
        'Me.Command(5).Enabled = True

End Sub
Private Sub Aceptar()
Dim num As Integer
Dim Alc_Obj_Liqs  As ADODB.Recordset

SCROLL 0

If Me.TextBox(4) = 0 Then

    MsgBox "No se suministraron Items/Cuotas/Porciones."
    
    MsgBox "Contador:" + STR(CONTADOR) + " .Tex_Tot_Monto:" + STR(Tex_Tot_Monto)
    
    Exit Sub
   
End If
num = CInt(Me.TextBox(4))
SCROLL 10

If num > 1 Then

    Me.TextBox(4) = num - 1
    
End If
Nro_P = FGNRO_LIQ()
Me.Adodc3.Recordset.AddNew

        Me.Adodc3.Recordset!usuario_liq = Usuario
        
        Me.Adodc3.Recordset!NRO_PLANI_PAGO = Nro_P
        
       
        Me.Adodc3.Recordset!Renglon = CONTADOR
           SCROLL 20

        Me.Adodc3.Recordset!CUOTA = Me.TextBox(5)
        
        Me.Adodc3.Recordset!Id_Objeto = Gid_obj
            
        Me.Adodc3.Recordset!Id_Instancia = Me.TextBox(0)
        
        Me.Adodc3.Recordset!Xinstancia = Tex_Items
        
        Me.Adodc3.Recordset!Id_Contri = Me.TextBox(0)

        CONCEPTO_SUB = RUB

        
        If Gid_obj = "PIC" And CONCEPTO_SUB = "301040508" Then
                
                
                Me.Adodc3.Recordset!CUOTA = STR(Year(Date)) + "15"
                
        Else
                
                Me.Adodc3.Recordset!CUOTA = Nro_P
        
        End If
       SCROLL 35
        
        
        Me.Adodc3.Recordset!Monto_Origi = CDbl(Me.TextBox(4)) * CDbl(Me.TextBox(6))
        
        Me.Adodc3.Recordset!Rubro = CONCEPTO_SUB
        
        Me.Adodc3.Recordset!Xid_Contri = Me.TextBox(0)
        
        Me.Adodc3.Recordset!Xnombre = Mid(Me.TextBox(1), 1, 65)
        
        Me.Adodc3.Recordset!Xdescripcion = Me.TextBox(3)
        
        Me.Adodc3.Recordset!Fec_pago = Date
        
        Me.Adodc3.Recordset!Tip_Liq = "Gen"
        
                
Me.Adodc3.Recordset.Update
SCROLL 41
    For i = 0 To 6
        If i <> 3 And i <> 6 Then
            Me.Command(i).Enabled = False
        ElseIf i = 6 Then
            Me.Command(i).Enabled = True
        End If
    Next i
    
    Dim respuesta As String

    respuesta = MsgBox("¿Desea ir a Recaudación?", vbYesNo + vbDefaultButton2, "ALCASIS")

    If respuesta = vbYes Then
        Unload Me
        frm_alc_recaudador_micasa.Show
    Else
        Me.Command(6).SetFocus
    End If

End Sub

Private Sub Form_Resize()
    Call Mover_der(Me, Frame2, 0)
    Call Mover_centrado(Me, Frame1)
End Sub

Private Sub TextBox_GotFocus(Index As Integer)
Dim cade As Long

If Index = 3 Then
    Me.TextBox(3).SelStart = 0
    cade = Len(Me.TextBox(3).Text)
    Me.TextBox(3).SelLength = cade
End If

End Sub

Private Sub TextBox_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If Index = 6 Then
        If KeyAscii = 46 Then KeyAscii = 44
    End If
    
End Sub

Private Sub Proxima()
    Nro_P = FGNRO_LIQ()
    Me.Label2.Caption = "Nro. de Planilla: " & Nro_P
        
    J = 1
    
    CONTADOR = 0
    
    Tex_Cantidad = 0
    
    Tex_Item = ""
    
    For i = 0 To 6
        Me.TextBox(i).Text = ""
    Next i
    
    Me.TextBox(5) = 0
    Tex_Items = ""
    
    For i = 0 To 6
        If i <> 3 And i <> 6 Then
            Me.Command(i).Enabled = True
        ElseIf i = 6 Then
            Me.Command(i).Enabled = False
        End If
    Next i

    Me.DataGrid1.SetFocus
End Sub
