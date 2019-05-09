VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_rev_cuadre_mov 
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command 
      Caption         =   "Cerrar"
      Height          =   615
      Index           =   2
      Left            =   7560
      TabIndex        =   25
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   240
      TabIndex        =   22
      Top             =   0
      Width           =   8775
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Revisión y Cuadre  de Movimientos  Diarios"
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
         Left            =   480
         TabIndex        =   23
         Top             =   360
         Width           =   8295
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "PROCESOS"
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
         Left            =   6600
         TabIndex        =   24
         Top             =   0
         Width           =   2175
      End
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Top             =   2160
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   8535
      Begin VB.TextBox TextBox 
         DataField       =   "Nro_Plani_Pago"
         DataSource      =   "FORMA_DE_PAGO"
         Height          =   315
         Index           =   5
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox TextBox 
         DataField       =   "Nro_Plani_Pago"
         DataSource      =   "FORMA_DE_PAGO"
         Height          =   315
         Index           =   4
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox TextBox 
         DataField       =   "Monto"
         DataSource      =   "FORMA_DE_PAGO"
         Height          =   315
         Index           =   3
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox TextBox 
         DataField       =   "Id_Rubro"
         DataSource      =   "FORMA_DE_PAGO"
         Height          =   315
         Index           =   2
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1800
         Width           =   1455
      End
      Begin VB.OptionButton Option 
         Height          =   315
         Index           =   1
         Left            =   5280
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton Option 
         Height          =   315
         Index           =   0
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TextBox 
         DataField       =   "Id_Voucher"
         DataSource      =   "FORMA_DE_PAGO"
         Height          =   315
         Index           =   1
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Busq 
         Height          =   315
         Left            =   5760
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox TextBox 
         DataField       =   "Nro_Plani_Pago"
         DataSource      =   "FORMA_DE_PAGO"
         Height          =   315
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command 
         Caption         =   "Buscar"
         Height          =   375
         Index           =   1
         Left            =   7200
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Fec_pago"
         DataSource      =   "FORMA_DE_PAGO"
         Height          =   315
         Left            =   0
         TabIndex        =   10
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16711681
         CurrentDate     =   38061
      End
      Begin MSAdodcLib.Adodc FORMA_DE_PAGO 
         Height          =   375
         Left            =   5880
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
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
         RecordSource    =   "FORMA_DE_PAGO"
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
      Begin VB.Label Label 
         Caption         =   "ID Instancia"
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
         Index           =   8
         Left            =   6120
         TabIndex        =   21
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label 
         Caption         =   "ID objeto"
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
         Index           =   7
         Left            =   4440
         TabIndex        =   19
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "Monto"
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
         Left            =   1800
         TabIndex        =   17
         Top             =   1560
         Width           =   1695
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
         Index           =   5
         Left            =   0
         TabIndex        =   15
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label 
         Caption         =   "Busqueda"
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
         Left            =   5760
         TabIndex        =   13
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label 
         Caption         =   "Estatus"
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
         Left            =   1680
         TabIndex        =   9
         Top             =   720
         Width           =   1695
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
         Left            =   0
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label 
         Caption         =   "Nº Voucher"
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
         Left            =   2880
         TabIndex        =   6
         Top             =   0
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
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frm_rev_cuadre_mov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Criterio As String

Private Sub Command_Click(Index As Integer)
    Select Case Index
        Case 1
            Call Buscar
        Case 2
            Cerrar = True
            Unload Me
    End Select
End Sub

Private Sub Buscar()
    
    If Me.Busq.Text = "" Then
        MsgBox "Introduzca un criterio de búsqueda", vbInformation + vbOKOnly, "ALCASIS"
        Exit Sub
    End If
    Dim strquery
    
    FORMA_DE_PAGO.Recordset.MoveFirst
       
    strquery = Criterio & Busq.Text

    FORMA_DE_PAGO.Recordset.Find strquery
    
    If FORMA_DE_PAGO.Recordset.EOF Then
    
        MsgBox "No encontrado", vbOKOnly, "ALCASIS"
'        dcmb_Busqueda.Text = ""
        Busq.SetFocus
    End If

End Sub

Private Sub Option_Click(Index As Integer)
Select Case Index
    Case 0
        Criterio = "Nro_Plani_Pago = "
    Case 1
        Criterio = "Id_Voucher = "
End Select
End Sub
