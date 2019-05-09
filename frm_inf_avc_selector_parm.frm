VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_inf_avc_selector_parm 
   Caption         =   "Gestión  de  Cobranza y Recaudación"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   11490
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_Veh_Comi_Bs 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   -120
      Locked          =   -1  'True
      TabIndex        =   80
      Top             =   6720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txt_Pic_Comi_Bs 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   -120
      Locked          =   -1  'True
      TabIndex        =   79
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txt_Pub_Comi_Bs 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   -120
      Locked          =   -1  'True
      TabIndex        =   78
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txt_Inm_Comi_Bs 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   -120
      Locked          =   -1  'True
      TabIndex        =   77
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txt_Otros_Comi_Bs 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   -120
      Locked          =   -1  'True
      TabIndex        =   76
      Top             =   8160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   74
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "Id_Objeto"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      DataSource      =   "ALC_OBJ_AVC"
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   73
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "% EFT_ K"
      Height          =   6735
      Left            =   360
      TabIndex        =   44
      Top             =   1440
      Width           =   10815
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   8880
         TabIndex        =   9
         Tag             =   "Cerrar Informe de Recaudación"
         Top             =   6000
         Width           =   1575
      End
      Begin VB.CommandButton cmd_recaudadores 
         Caption         =   "A&dm de Recaudadores"
         Height          =   615
         Left            =   7320
         TabIndex        =   8
         Tag             =   "Visualizar Informe de Recaudación para su posterior impresión"
         Top             =   6000
         Width           =   1575
      End
      Begin VB.TextBox txt_total_cuota 
         Height          =   285
         Left            =   1320
         TabIndex        =   75
         Top             =   5040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmd_imprimir_detallado 
         Caption         =   "I&mprimir Detallado"
         Enabled         =   0   'False
         Height          =   615
         Left            =   5760
         TabIndex        =   7
         Tag             =   "Visualizar Informe de Recaudación para su posterior impresión"
         Top             =   6000
         Width           =   1575
      End
      Begin VB.CommandButton cmd_imprimir_sumario 
         Caption         =   "&Imprimir Sumario"
         Enabled         =   0   'False
         Height          =   615
         Left            =   4200
         TabIndex        =   6
         Tag             =   "Visualizar Informe de Recaudación para su posterior impresión"
         Top             =   6000
         Width           =   1575
      End
      Begin VB.TextBox txt_total_cancel 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   5520
         Width           =   975
      End
      Begin VB.TextBox txt_total_comision 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox txt_total_cantidad 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   5520
         Width           =   975
      End
      Begin VB.TextBox txt_otros_cancel 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txt_otros_comision 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox txt_otros_cantidad 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txt_veh_cancel 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txt_pub_cancel 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txt_pic_cancel 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txt_inm_cancel 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txt_veh_comision 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txt_pub_comision 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txt_pic_comision 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox txt_inm_comision 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txt_veh_cantidad 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txt_pub_cantidad 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txt_pic_cantidad 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txt_inm_cantidad 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txt_vigentes_eft_bs 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txt_cuotas_eft_bs 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txt_vencidos_eft_bs 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txt_vigentes_eft_k 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txt_cuotas_eft_k 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txt_vencidos_eft_k 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txt_vigentes_monto 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox txt_cuotas_monto 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txt_asig_monto 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txt_vencidos_monto 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txt_vigentes_avc 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txt_cuotas_avc 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox txt_asig_avc 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txt_vencidos_avc 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmd_aceptar 
         Caption         =   "&Aceptar"
         Height          =   615
         Left            =   2640
         TabIndex        =   5
         Tag             =   "Visualizar Informe de Recaudación para su posterior impresión"
         Top             =   6000
         Width           =   1575
      End
      Begin MSDataListLib.DataList DList_tributo 
         Bindings        =   "frm_inf_avc_selector_parm.frx":0000
         Height          =   1815
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   3201
         _Version        =   393216
         ListField       =   "Descripcion"
         BoundColumn     =   "Id_Obj"
      End
      Begin MSDataListLib.DataList Dlist_status 
         Bindings        =   "frm_inf_avc_selector_parm.frx":0019
         Height          =   1815
         Left            =   3120
         TabIndex        =   1
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   3201
         _Version        =   393216
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "STATUS"
      End
      Begin MSDataListLib.DataList DList_recaudador 
         Bindings        =   "frm_inf_avc_selector_parm.frx":0038
         Height          =   1815
         Left            =   5880
         TabIndex        =   2
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   3201
         _Version        =   393216
         ListField       =   "Nombre"
         BoundColumn     =   "Id_Recaudador"
      End
      Begin MSComCtl2.DTPicker txt_desde 
         Height          =   375
         Left            =   8640
         TabIndex        =   3
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   55443459
         CurrentDate     =   38028
      End
      Begin MSComCtl2.DTPicker txt_hasta 
         Height          =   375
         Left            =   8640
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   55443459
         CurrentDate     =   38028
      End
      Begin VB.Label lbl_total 
         Caption         =   "Total:"
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
         Left            =   6360
         TabIndex        =   72
         Top             =   5520
         Width           =   615
      End
      Begin VB.Label lbl_otros 
         Caption         =   "Otros:"
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
         Left            =   6360
         TabIndex        =   71
         Top             =   5040
         Width           =   615
      End
      Begin VB.Label lbl_veh 
         Caption         =   "Vehs:"
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
         Left            =   6360
         TabIndex        =   70
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label lbl_inm 
         Caption         =   "Inms:"
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
         Left            =   6360
         TabIndex        =   69
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label lbl_comision 
         Alignment       =   2  'Center
         Caption         =   "Comisión"
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
         Left            =   8160
         TabIndex        =   68
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label lbl_cancel_bs 
         Caption         =   "% Cancel_ Bs."
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
         Left            =   9240
         TabIndex        =   67
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label lbl_pic 
         Caption         =   "Pics:"
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
         Left            =   6360
         TabIndex        =   66
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label lbl_cant 
         Alignment       =   2  'Center
         Caption         =   "Cantidad"
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
         Left            =   6960
         TabIndex        =   65
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lbl_tribut 
         Caption         =   "Tributo"
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
         Left            =   6360
         TabIndex        =   64
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label lbl_pub 
         Caption         =   "Pubs:"
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
         Left            =   6360
         TabIndex        =   63
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label lbl_eft_bs 
         Caption         =   "% EFT_ Bs."
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
         Left            =   4800
         TabIndex        =   62
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lbl_nomina 
         Caption         =   "Nómina de  Comisiones por Cuotas Canceladas"
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
         Left            =   6360
         TabIndex        =   61
         Top             =   2280
         Width           =   4095
      End
      Begin VB.Label lbl_cuotas 
         Caption         =   "Cuotas Canceladas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   60
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label lbl_Monto 
         Alignment       =   2  'Center
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
         Left            =   2640
         TabIndex        =   59
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lbl_eft_k 
         Caption         =   "% EFT_ K"
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
         Left            =   3840
         TabIndex        =   58
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lbl_estadistica 
         Caption         =   "Estadística  Desempeño Por  Avisos de Cobros y Cuotas Canceladas"
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
         TabIndex        =   57
         Top             =   2280
         Width           =   6015
      End
      Begin VB.Label lbl_vencidos 
         Caption         =   "Vencidos"
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
         TabIndex        =   56
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label lbl_cant_avc 
         Alignment       =   2  'Center
         Caption         =   "Cantidad Avisos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   55
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lbl_asignados 
         Caption         =   "Asignados"
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
         TabIndex        =   54
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label lbl_vigentes 
         Caption         =   "Vigentes"
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
         TabIndex        =   53
         Top             =   4680
         Width           =   855
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
         Left            =   8640
         TabIndex        =   52
         Top             =   120
         Width           =   1335
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
         Left            =   8640
         TabIndex        =   51
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lbl_status 
         Caption         =   "Status"
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
         Left            =   3120
         TabIndex        =   50
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label lbl_tributo 
         Caption         =   "Tributos a Procesar:"
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
         TabIndex        =   49
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label lbl_recaudadores 
         Caption         =   "Recaudadores:"
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
         Left            =   5880
         TabIndex        =   48
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   1800
      TabIndex        =   45
      Top             =   360
      Width           =   8295
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   "Gestión  de  Cobranza "
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
         TabIndex        =   47
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "  y Recaudación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   2640
         TabIndex        =   46
         Top             =   360
         Width           =   5655
      End
   End
   Begin MSAdodcLib.Adodc TAB_RECAUDA 
      Height          =   375
      Left            =   8640
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      RecordSource    =   "SELECT * FROM Tab_Recaudador WHERE (status = 1) ORDER BY Id_Recaudador DESC, Nombre DESC"
      Caption         =   "TAB_RECAUDA"
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
   Begin MSAdodcLib.Adodc TAB_ID_OBJ 
      Height          =   375
      Left            =   6120
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      RecordSource    =   "SELECT Id_Obj, Descripcion FROM TAB_ID_OBJ ORDER BY Id_Obj"
      Caption         =   "TAB_ID_OBJ"
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
   Begin MSAdodcLib.Adodc TABLA_STATUS_AVC 
      Height          =   375
      Left            =   3600
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      RecordSource    =   "SELECT STATUS, DESCRIPCION FROM TABLA_STATUS_AVC ORDER BY STATUS"
      Caption         =   "TABLA_STATUS_AVC"
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
   Begin MSAdodcLib.Adodc ALC_OBJ_AVC 
      Height          =   375
      Left            =   1080
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      RecordSource    =   "select * from ALC_OBJ_AVC WHERE Id_Objeto=''"
      Caption         =   "ALC_OBJ_AVC"
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
Attribute VB_Name = "frm_inf_avc_selector_parm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MAT(4, 2)
Dim K_pics, K_Inms, K_Pubs, K_Vehs, K_Otros As Integer
                    
Private Sub cmd_aceptar_Click()

Dim Clave As String


If Me.DList_recaudador.Text = "" Then
    MsgBox "Por favor, suministre el Recaudador", vbCritical, "ALCALSIS"
    Exit Sub
End If

'If Me.Dlist_status.Text = "" Then
'    MsgBox "Por favor, suministre el Status", vbCritical, "ALCALSIS"
'    Exit Sub
'End If
'
'If Me.DList_tributo.Text = "" Then
'    MsgBox "Por favor, suministre el Tipo de Tributo", vbCritical, "ALCALSIS"
'    Exit Sub
'End If

Init_Contadores

'Me.Com_Print_Sumario.Enabled = False
'Me.Com_Print_Detallado.Enabled = False

If Me.Dlist_status.BoundText = "*" Then
    
    
'    sqlstr = "SELECT Nro_Plani_AVC, Cuota, Renglon, Id_Objeto, Id_Instancia, Monto_Origi, Rubro, Fec_AVC, descuento, Id_Aso, autonumerico, Cod_Recauda, Status,Id_Objeto"
    sqlstr = "SELECT * "
    
    sqlstr = sqlstr + " FROM ALC_OBJ_AVC  "
    
    sqlstr = sqlstr + "  WHERE (Fec_AVC<= CONVERT(DATETIME, '" + Format(Me.txt_desde.Value, "mm/dd/yyyy") + "', 102))" 'Fec_AVC>=" + "'" + Format(STR(Me.txt_desde.Value), "dd/mm/YYYY") + "'"
    
    sqlstr = sqlstr + "  And (Fec_AVC<= CONVERT(DATETIME, '" + Format(Me.txt_hasta.Value, "mm/dd/yyyy") + "', 102))"         'Fec_AVC<=" + "'" + Format(STR(Me.txt_hasta.Value), "dd/mm/YYYY") + "'"
    
    sqlstr = sqlstr + " AND (Cod_Recauda)='" + Me.DList_recaudador.BoundText + "' AND ALC_OBJ_AVC.Status In('VI','CA','VE',NULL)"
    
    If Me.DList_tributo.BoundText <> "*" Then
                
        sqlstr = sqlstr + "  And Id_Objeto=" + "'" + (Me.DList_tributo.BoundText) + "'"
        
    End If
    
    sqlstr = sqlstr + "   ORDER BY Nro_Plani_AVC"
    
Else

       
'    sqlstr = "SELECT Nro_Plani_AVC, Cuota, Renglon, Id_Objeto, Id_Instancia, Monto_Origi, Rubro, Fec_AVC, descuento, d_Aso, autonumerico, Cod_Recauda, Status,Id_Objeto"
    sqlstr = "SELECT * "
    
    sqlstr = sqlstr + " FROM ALC_OBJ_AVC  "
    
    sqlstr = sqlstr + "WHERE (Fec_AVC>=" + "'" + Format(STR(Me.txt_desde.Value), "dd/mm/yyyy") + "'"
    
    sqlstr = sqlstr + "  And Fec_AVC<=" + "'" + Format(STR(Me.txt_hasta.Value), "dd/mm/yyyy") + "')" + " AND Cod_Recauda=" + "'" + (Me.DList_recaudador.BoundText) + "'" + " AND ALC_OBJ_AVC.Status In(" + "'" + (Me.Dlist_status.BoundText) + "')"
    
    If Me.DList_tributo.BoundText <> "*" Then
                
        sqlstr = sqlstr + "  And Id_Objeto=" + "'" + (Me.DList_tributo.BoundText) + "'"
        
    End If
    
    sqlstr = sqlstr + "   ORDER BY Id_Objeto,Status,Nro_Plani_AVC"
    
End If

Me.ALC_OBJ_AVC.ConnectionString = "SIAGEP"
Me.ALC_OBJ_AVC.CommandType = adCmdText
'MsgBox sqlstr
Me.ALC_OBJ_AVC.RecordSource = sqlstr
Me.ALC_OBJ_AVC.Refresh

    
If Me.ALC_OBJ_AVC.Recordset.EOF = True Then
    
    
        MsgBox "Conjunto de datos solicitado está vacio.Verifique o Intente con Otro.Gracias.", vbCritical, "ALCASIS"
        
        Exit Sub
        
End If
    
Me.ALC_OBJ_AVC.Recordset.MoveFirst


Clave = Me.ALC_OBJ_AVC.Recordset!nro_plani_avc

MAT(1, 1) = MAT(1, 1) + 1  ' Cuenta y Acumula por Aviso de Cobro Asignados. No! por Cuotas


Rem ACUMULA_X_STATUS

Do While Me.ALC_OBJ_AVC.Recordset.EOF = False
       
      If Clave <> Me.ALC_OBJ_AVC.Recordset!nro_plani_avc Then  ' Cambio de Nro_Plani_AVC
            
            MAT(1, 1) = MAT(1, 1) + 1
            
            Clave = Me.ALC_OBJ_AVC.Recordset!nro_plani_avc
            
            
            ACUMULA_X_STATUS_DIFER
        
       Else
       
             ACUMULA_X_STATUS_IGUALES
       
       End If
       
       Me.txt_total_cuota = NZ(Val(Me.txt_total_cuota), 0) + 1
       
       
       MAT(1, 2) = MAT(1, 2) + Me.ALC_OBJ_AVC.Recordset!Monto_Origi
       
       Me.txt_asig_avc.Text = MAT(1, 1)
       Me.txt_asig_monto.Text = MAT(1, 2)
       
       
     Me.ALC_OBJ_AVC.Recordset.MoveNext
    
Loop

Me.cmd_imprimir_detallado.Enabled = True
Me.cmd_imprimir_sumario.Enabled = True

  
Me.txt_vencidos_eft_k.Text = Format((Val(txt_vencidos_avc.Text) * 100) / NZ(Val(Me.txt_asig_avc.Text), 0), "00.000000")
Me.txt_cuotas_eft_k.Text = Format((Val(txt_cuotas_avc.Text) * 100) / NZ(Val(Me.txt_asig_avc.Text), 0), "00.000000")
Me.txt_vigentes_eft_k.Text = Format((Val(txt_vigentes_avc.Text) * 100) / NZ(Val(Me.txt_asig_avc.Text), 0), "00.000000")

Me.txt_vencidos_eft_bs.Text = Format((Val(txt_vencidos_monto.Text) * 100) / Val(Me.txt_asig_monto.Text), "00.000000")
Me.txt_cuotas_eft_bs.Text = Format((Val(txt_cuotas_monto.Text) * 100) / Val(Me.txt_asig_monto.Text), "00.000000")
Me.txt_vigentes_eft_bs.Text = Format((Val(txt_vigentes_monto.Text) * 100) / Val(Me.txt_asig_monto.Text), "00.000000")

txt_pic_comision.Text = Val(Me.txt_pic_cantidad.Text) * Val(Me.txt_Pic_Comi_Bs.Text)
txt_inm_comision.Text = Val(Me.txt_inm_cantidad.Text) * Val(Me.txt_Inm_Comi_Bs.Text)
txt_pub_comision.Text = Val(Me.txt_pub_cantidad.Text) * Val(Me.txt_Pub_Comi_Bs.Text)
txt_veh_comision.Text = Val(Me.txt_veh_cantidad.Text) * Val(Me.txt_Veh_Comi_Bs.Text)
txt_otros_comision.Text = Val(Me.txt_otros_cantidad.Text) * Val(Me.txt_Otros_Comi_Bs.Text)

If Me.txt_cuotas_monto.Text <> "" Then

    Me.txt_pic_cancel.Text = Format((Val(txt_pic_comision.Text) * 100) / Val(txt_cuotas_monto.Text), "00.000000")
    Me.txt_inm_cancel.Text = Format((Val(txt_inm_comision.Text) * 100) / Val(txt_cuotas_monto.Text), "00.000000")
    Me.txt_pub_cancel.Text = Format((Val(txt_pub_comision.Text) * 100) / Val(txt_cuotas_monto.Text), "00.000000")
    Me.txt_veh_cancel.Text = Format((Val(txt_veh_comision.Text) * 100) / Val(txt_cuotas_monto.Text), "00.000000")
    Me.txt_otros_cancel.Text = Format((Val(Me.txt_otros_comision.Text) * 100) / Val(txt_cuotas_monto.Text), "00.000000")
    Me.txt_total_cancel.Text = Format((Val(Me.txt_total_comision.Text) * 100) / Val(txt_cuotas_monto.Text), "00.000000")
    
Else
    MsgBox "El monto de las cuotas canceladaas es null, por favor verifique la seleccion suministrada, gracias", vbCritical, "ALCALSIS"
    Me.cmd_imprimir_detallado.Enabled = False
    Me.cmd_imprimir_sumario.Enabled = False
End If
txt_total_cantidad.Text = Val(Me.txt_pic_cantidad.Text) + Val(Me.txt_inm_cantidad.Text) + Val(Me.txt_pub_cantidad.Text) + Val(Me.txt_veh_cantidad.Text) + Val(Me.txt_otros_cantidad.Text)
txt_total_comision.Text = Val(Me.txt_pic_comision.Text) + Val(Me.txt_inm_comision.Text) + Val(Me.txt_pub_comision.Text) + Val(Me.txt_veh_comision.Text) + Val(Me.txt_otros_comision.Text)
End Sub

Private Sub cmd_aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_aceptar.FontBold = True
    Me.cmd_imprimir_detallado.FontBold = False
    Me.cmd_imprimir_sumario.FontBold = False
    Me.cmd_recaudadores.FontBold = False
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_recaudadores.FontBold = False
Me.cmd_cerrar.FontBold = True
    Me.cmd_aceptar.FontBold = False
    Me.cmd_imprimir_detallado.FontBold = False
    Me.cmd_imprimir_sumario.FontBold = False
End Sub

Private Sub cmd_imprimir_detallado_Click()
MsgBox "Disculpe, esta en desarrollo"
End Sub

Private Sub cmd_imprimir_detallado_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_recaudadores.FontBold = False
Me.cmd_cerrar.FontBold = False
    Me.cmd_aceptar.FontBold = False
    Me.cmd_imprimir_detallado.FontBold = True
    Me.cmd_imprimir_sumario.FontBold = False
End Sub

Private Sub cmd_imprimir_sumario_Click()
MsgBox "Disculpe, esta en desarrollo"
End Sub

Private Sub cmd_imprimir_sumario_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_aceptar.FontBold = False
    Me.cmd_imprimir_detallado.FontBold = False
    Me.cmd_imprimir_sumario.FontBold = True
    Me.cmd_recaudadores.FontBold = False
End Sub

Private Sub cmd_recaudadores_Click()
    frm_inf_tab_recaudador_mantenimiento.Show
End Sub

Private Sub cmd_recaudadores_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_aceptar.FontBold = False
    Me.cmd_imprimir_detallado.FontBold = False
    Me.cmd_imprimir_sumario.FontBold = False
    Me.cmd_recaudadores.FontBold = True
End Sub

Private Sub DList_recaudador_Click()
Dim sqlstr As String

    TAB_RECAUDA.Recordset.MoveFirst
    
    If Not Me.DList_recaudador.BoundText = "" Then
    
        sqlstr = "Id_Recaudador  = '" + Me.DList_recaudador.BoundText + "'"
        TAB_RECAUDA.Recordset.Find sqlstr
        Me.txt_Pic_Comi_Bs = TAB_RECAUDA.Recordset!Pic_Comi_Bs
        Me.txt_Inm_Comi_Bs = TAB_RECAUDA.Recordset!Inm_Comi_Bs
        Me.txt_Pub_Comi_Bs = TAB_RECAUDA.Recordset!Pub_Comi_Bs
        Me.txt_Veh_Comi_Bs = TAB_RECAUDA.Recordset!Veh_Comi_Bs
        Me.txt_Otros_Comi_Bs = TAB_RECAUDA.Recordset!Otros_Comi_Bs
        
    End If
End Sub

Private Sub DList_recaudador_GotFocus()
    Me.lbl_recaudadores.ForeColor = vbRed
End Sub

Private Sub DList_recaudador_LostFocus()
    Me.lbl_recaudadores.ForeColor = vbWindowText
End Sub

Private Sub Dlist_status_GotFocus()
    Me.lbl_status.ForeColor = vbRed
End Sub

Private Sub Dlist_status_LostFocus()
    Me.lbl_status.ForeColor = vbWindowText
End Sub

Private Sub DList_tributo_GotFocus()
    Me.lbl_tributo.ForeColor = vbRed
    
End Sub

Private Sub DList_tributo_LostFocus()
    Me.lbl_tributo.ForeColor = vbWindowText
End Sub

Private Sub Form_Load()
    Me.txt_desde.Value = Format(Date, "dd/mm/yyyy")
    Me.txt_hasta.Value = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub Form_Resize()
    Call Mover_der(Me, Frame2, 0)
    Call Mover_centrado(Me, Frame1)
End Sub

Private Sub ACUMULA_X_STATUS_IGUALES()

Select Case Me.ALC_OBJ_AVC.Recordset!STATUS
       
            Case "VE"
                
                
                MAT(2, 2) = MAT(2, 2) + Me.ALC_OBJ_AVC.Recordset!Monto_Origi
                
                Me.txt_vencidos_avc.Text = MAT(2, 1)
                Me.txt_vencidos_monto.Text = MAT(2, 2)
            
            
            Case "CA"
            
                MAT(3, 1) = MAT(3, 1) + 1
                MAT(3, 2) = MAT(3, 2) + Me.ALC_OBJ_AVC.Recordset!Monto_Origi
            
            
                Me.txt_cuotas_avc.Text = MAT(3, 1)
                Me.txt_cuotas_monto.Text = MAT(3, 2)
                
                Computa_Comi
                
            
            Case "VI"
            
                
                MAT(4, 2) = MAT(4, 2) + Me.ALC_OBJ_AVC.Recordset!Monto_Origi
            
                Me.txt_vigentes_avc.Text = MAT(4, 1)
                Me.txt_vigentes_monto.Text = MAT(4, 2)
            
            Case Else
       
       End Select
       

End Sub
Private Sub ACUMULA_X_STATUS_DIFER()

Select Case Me.ALC_OBJ_AVC.Recordset!STATUS
       
            Case "VE"
                
                
                MAT(2, 1) = MAT(2, 1) + 1
                
                MAT(2, 2) = MAT(2, 2) + Me.ALC_OBJ_AVC.Recordset!Monto_Origi
                
                Me.txt_vencidos_avc.Text = MAT(2, 1)
                Me.txt_vencidos_monto.Text = MAT(2, 2)
            
            
            Case "CA"
            
                MAT(3, 1) = MAT(3, 1) + 1
                
                MAT(3, 2) = MAT(3, 2) + Me.ALC_OBJ_AVC.Recordset!Monto_Origi
            
            
                Me.txt_cuotas_avc.Text = MAT(3, 1)
                Me.txt_cuotas_monto.Text = MAT(3, 2)
                
                Computa_Comi
                
            
            Case "VI"
            
                MAT(4, 1) = MAT(4, 1) + 1
                
                MAT(4, 2) = MAT(4, 2) + Me.ALC_OBJ_AVC.Recordset!Monto_Origi
            
                Me.txt_vigentes_avc.Text = MAT(4, 1)
                Me.txt_vigentes_monto.Text = MAT(4, 2)
            
            Case Else
       
       End Select
       

End Sub

Private Sub Init_Contadores()

For i = 1 To 4

    For J = 1 To 2
    
        MAT(i, J) = 0
        
    Next J
    
Next i
   
   
'Me.Me.txt_asig_avc.Text = 0
'Me.Me.txt_asig_monto.Text = 0
'
'Me.Ven_Can = 0
'Me.Me.txt_vencidos_monto.Text = 0
'
'Me.Vig_Can = 0
'Me.txt_vigentes_monto.Text = 0
'
'Me.txt_cuotas_avc = 0
'Me.txt_cuotas_monto.Text = 0
'
'Me.K_Pics = 0
'
'Me.K_Inms = 0
'
'Me.K_Pubs = 0
'
'Me.K_Vehs = 0
'
'Me.K_Otros = 0

End Sub

Private Sub Computa_Comi()


           Select Case Me.ALC_OBJ_AVC.Recordset!Id_Objeto
       
              Case "PIC"
              
                    K_pics = K_pics + 1
                    Me.txt_pic_cantidad.Text = K_pics
                    
              Case "INM"
                    K_Inms = K_Inms + 1
                    Me.txt_inm_cantidad.Text = K_Inms
              Case "PUB"
              
                    K_Pubs = K_Pubs + 1
                    Me.txt_pub_cantidad.Text = K_Pubs
              Case "VEH"
              
                    K_Vehs = K_VEH + 1
                    Me.txt_veh_cantidad.Text = K_Vehs
              Case Else
              
                    K_Otros = K_Otros + 1
                    Me.txt_otros_cantidad.Text = K_Otros
       End Select
     



End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_aceptar.FontBold = False
    Me.cmd_imprimir_detallado.FontBold = False
    Me.cmd_imprimir_sumario.FontBold = False
    Me.cmd_recaudadores.FontBold = False
End Sub

Private Sub txt_desde_GotFocus()
    Me.Lbl_fecha_desde.ForeColor = vbRed
End Sub

Private Sub txt_desde_LostFocus()
    Me.Lbl_fecha_desde.ForeColor = vbWindowText
End Sub

Private Sub txt_hasta_GotFocus()
Me.lbl_fecha_hasta.ForeColor = vbRed
End Sub

Private Sub txt_hasta_LostFocus()
   Me.lbl_fecha_hasta.ForeColor = vbWindowText
End Sub

