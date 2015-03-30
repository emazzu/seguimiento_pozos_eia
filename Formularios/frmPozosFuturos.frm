VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form frmPozosFuturos 
   Caption         =   "ABM - Pozos"
   ClientHeight    =   11010
   ClientLeft      =   -1380
   ClientTop       =   -2355
   ClientWidth     =   17160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPozosFuturos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   17160
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab 
      Height          =   11625
      Left            =   405
      TabIndex        =   2
      Top             =   180
      Width           =   17085
      _ExtentX        =   30136
      _ExtentY        =   20505
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Pozos"
      TabPicture(0)   =   "frmPozosFuturos.frx":2CFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdImportarDatosVintage"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdExportarPlanilla"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FrameTree"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Ficha de Pozos"
      TabPicture(1)   =   "frmPozosFuturos.frx":2D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdGuardar"
      Tab(1).Control(1)=   "fraGrupo6"
      Tab(1).Control(2)=   "fraGrupo4"
      Tab(1).Control(3)=   "fraGrupo3"
      Tab(1).Control(4)=   "fraGrupo2"
      Tab(1).Control(5)=   "fraGrupo1"
      Tab(1).Control(6)=   "fraGrupo5"
      Tab(1).ControlCount=   7
      Begin VB.Frame Frame3 
         Height          =   7980
         Left            =   4590
         TabIndex        =   166
         Top             =   1470
         Width           =   10515
         Begin FPSpreadADO.fpSpread spdCab 
            Height          =   4515
            Left            =   90
            TabIndex        =   167
            Top             =   180
            Width           =   10000
            _Version        =   393216
            _ExtentX        =   17639
            _ExtentY        =   7964
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpreadDesigner  =   "frmPozosFuturos.frx":2D32
         End
         Begin FPSpreadADO.fpSpread spdAUX 
            Height          =   915
            Left            =   3990
            TabIndex        =   168
            Top             =   4740
            Visible         =   0   'False
            Width           =   5055
            _Version        =   393216
            _ExtentX        =   8916
            _ExtentY        =   1614
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Protect         =   0   'False
            SpreadDesigner  =   "frmPozosFuturos.frx":2F06
         End
         Begin MSComDlg.CommonDialog comDestino 
            Left            =   1830
            Top             =   7050
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComctlLib.ImageList iml 
            Left            =   1170
            Top             =   6960
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   16777215
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   16777215
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPozosFuturos.frx":30E1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPozosFuturos.frx":3433
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPozosFuturos.frx":3785
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FrameTree 
         Height          =   7980
         Left            =   90
         TabIndex        =   164
         Top             =   1470
         Width           =   4095
         Begin MSComctlLib.TreeView Tree 
            Height          =   7620
            Left            =   60
            TabIndex        =   165
            Top             =   150
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   13441
            _Version        =   393217
            Indentation     =   529
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
            Appearance      =   1
         End
      End
      Begin VB.Frame Frame8 
         Height          =   1095
         Left            =   90
         TabIndex        =   161
         Top             =   360
         Width           =   885
         Begin VB.CommandButton cmdTree 
            Height          =   885
            Left            =   30
            Picture         =   "frmPozosFuturos.frx":3D97
            Style           =   1  'Graphical
            TabIndex        =   162
            Top             =   150
            Width           =   795
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1095
         Left            =   13440
         TabIndex        =   152
         Top             =   360
         Width           =   2430
         Begin VB.CommandButton cmdDMA 
            Height          =   915
            Left            =   1620
            Picture         =   "frmPozosFuturos.frx":46D2
            Style           =   1  'Graphical
            TabIndex        =   155
            ToolTipText     =   "Exportar formato DMA"
            Top             =   135
            Width           =   735
         End
         Begin VB.CommandButton cmdPER 
            Height          =   915
            Left            =   855
            Picture         =   "frmPozosFuturos.frx":70F7
            Style           =   1  'Graphical
            TabIndex        =   154
            ToolTipText     =   "Exportar formato Personalizado"
            Top             =   135
            Width           =   735
         End
         Begin VB.CommandButton cmdSTD 
            Height          =   915
            Left            =   45
            Picture         =   "frmPozosFuturos.frx":9B1C
            Style           =   1  'Graphical
            TabIndex        =   153
            ToolTipText     =   "Exportar formato Estandar"
            Top             =   135
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   1020
         TabIndex        =   143
         Top             =   360
         Width           =   4395
         Begin VB.TextBox txtDato 
            Height          =   345
            Left            =   750
            TabIndex        =   1
            Top             =   660
            Width           =   1815
         End
         Begin VB.CommandButton cmdBuscar1 
            Caption         =   "Filtra por &columna"
            Height          =   300
            Left            =   2580
            TabIndex        =   158
            Top             =   690
            Width           =   1710
         End
         Begin VB.ComboBox cmb 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmPozosFuturos.frx":C33A
            Left            =   750
            List            =   "frmPozosFuturos.frx":C33C
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   225
            Width           =   3555
         End
         Begin VB.Label lbl 
            Caption         =   "Campo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   144
            Top             =   285
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   5460
         TabIndex        =   142
         Top             =   360
         Width           =   7935
         Begin VB.CommandButton cmdFiltrar 
            Caption         =   "Filtrar por &ubicación"
            Height          =   300
            Left            =   5880
            TabIndex        =   160
            Top             =   720
            Width           =   1980
         End
         Begin FPSpreadADO.fpSpread spdFiltro 
            Height          =   525
            Left            =   90
            TabIndex        =   159
            Top             =   150
            Width           =   7935
            _Version        =   393216
            _ExtentX        =   13996
            _ExtentY        =   926
            _StockProps     =   64
            AutoSize        =   -1  'True
            DAutoCellTypes  =   0   'False
            DAutoSizeCols   =   1
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   20
            MaxRows         =   1
            Protect         =   0   'False
            ScrollBars      =   0
            SpreadDesigner  =   "frmPozosFuturos.frx":C33E
            UserResize      =   1
            Appearance      =   1
         End
      End
      Begin VB.CommandButton cmdGuardar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -58530
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPozosFuturos.frx":C555
         Style           =   1  'Graphical
         TabIndex        =   135
         ToolTipText     =   "Guardar"
         Top             =   8520
         UseMaskColor    =   -1  'True
         Width           =   510
      End
      Begin VB.CommandButton cmdExportarPlanilla 
         Height          =   285
         Left            =   15720
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPozosFuturos.frx":C901
         Style           =   1  'Graphical
         TabIndex        =   134
         ToolTipText     =   "Exportar planilla a Excel"
         Top             =   9765
         UseMaskColor    =   -1  'True
         Width           =   285
      End
      Begin VB.CommandButton cmdImportarDatosVintage 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   16080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPozosFuturos.frx":CC43
         Style           =   1  'Graphical
         TabIndex        =   133
         ToolTipText     =   "Importar Datos desde Vintage Data"
         Top             =   9765
         UseMaskColor    =   -1  'True
         Width           =   285
      End
      Begin VB.Frame fraGrupo6 
         Caption         =   "Calculos. Tiempo Entre..."
         Height          =   2835
         Left            =   -63570
         TabIndex        =   132
         Top             =   8670
         Width           =   5565
         Begin VB.TextBox txtTiempoEntreEntregaEIAaDMAyAprobacion 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2910
            MaxLength       =   9
            TabIndex        =   126
            Top             =   2445
            Width           =   2580
         End
         Begin VB.TextBox txtTiempoEntrePrimerMonografiaYPedidoEIA 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            MaxLength       =   9
            TabIndex        =   124
            Top             =   2430
            Width           =   2760
         End
         Begin VB.TextBox txtTiempoEntreRecepcionETIAYpresentacionAnteDMA 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   105
            MaxLength       =   9
            TabIndex        =   116
            Top             =   1200
            Width           =   2760
         End
         Begin VB.TextBox txtTiempoEntrePrimerMonografiayRecepcionETIA 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2910
            MaxLength       =   9
            TabIndex        =   114
            Top             =   600
            Width           =   2580
         End
         Begin VB.TextBox txtTiempoEntreVisitaYAprobacionFinaldeDMA 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   105
            MaxLength       =   9
            TabIndex        =   120
            Top             =   1800
            Width           =   2760
         End
         Begin VB.TextBox txtTiempoEntrePresentacionAnteDMAyVisita 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2910
            MaxLength       =   9
            TabIndex        =   118
            Top             =   1200
            Width           =   2580
         End
         Begin VB.TextBox txtTiempoEntrePresentacionDePozoYAprobacionFinalDMA 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2910
            MaxLength       =   9
            TabIndex        =   122
            Top             =   1800
            Width           =   2580
         End
         Begin VB.TextBox txtTiempoEntrePedidoETIAyRecepcionETIA 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   105
            MaxLength       =   9
            TabIndex        =   112
            Top             =   585
            Width           =   2760
         End
         Begin VB.Label lblTiempoEntreEntregaEIAaDMAyAprobacion 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Entrega EIA a DMA y Aprob"
            Height          =   195
            Left            =   2910
            TabIndex        =   125
            Top             =   2220
            Width           =   2280
         End
         Begin VB.Label lblTiempoEntrePrimerMonografiaYPedidoEIA 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "1° Monografia y Pedido EIA"
            Height          =   195
            Left            =   135
            TabIndex        =   123
            Top             =   2205
            Width           =   2295
         End
         Begin VB.Label lblTiempoEntreRecepcionETIAYpresentacionAnteDMA 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Recep. EIA y Pres. ante DMA"
            Height          =   195
            Left            =   150
            TabIndex        =   115
            Top             =   990
            Width           =   2385
         End
         Begin VB.Label lblTiempoEntrePrimerMonografiayRecepcionETIA 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "1° Monografia y recep. EIA"
            Height          =   195
            Left            =   2910
            TabIndex        =   113
            Top             =   390
            Width           =   2250
         End
         Begin VB.Label lblTiempoEntreVisitaYAprobacionFinaldeDMA 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Visita y Aprob. Final de DMA"
            Height          =   195
            Left            =   150
            TabIndex        =   119
            Top             =   1605
            Width           =   2340
         End
         Begin VB.Label lblTiempoEntrePresentacionAnteDMAyVisita 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Present. Ante DMA y Visita"
            Height          =   195
            Left            =   2910
            TabIndex        =   117
            Top             =   1005
            Width           =   2250
         End
         Begin VB.Label lblTiempoEntrePresentacionDePozoYAprobacionFinalDMA 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Pres. y Aprob. Final DMA"
            Height          =   195
            Left            =   2925
            TabIndex        =   121
            Top             =   1590
            Width           =   2040
         End
         Begin VB.Label lblTiempoEntrePedidoETIAyRecepcionETIA 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Pedido EIA y recepcion EIA"
            Height          =   195
            Left            =   150
            TabIndex        =   111
            Top             =   360
            Width           =   2250
         End
      End
      Begin VB.Frame fraGrupo4 
         Caption         =   "Aprobacion"
         Height          =   8265
         Left            =   -69420
         TabIndex        =   130
         Top             =   3240
         Width           =   5775
         Begin VB.Frame Frame6 
            Caption         =   "Site Visit Comments"
            Height          =   1965
            Left            =   45
            TabIndex        =   150
            Top             =   5460
            Width           =   5670
            Begin FPSpreadADO.fpSpread spdDet2 
               Height          =   1650
               Left            =   45
               TabIndex        =   157
               Top             =   225
               Width           =   5550
               _Version        =   393216
               _ExtentX        =   9790
               _ExtentY        =   2910
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SpreadDesigner  =   "frmPozosFuturos.frx":CF55
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Consultant Recomendation"
            Height          =   1740
            Left            =   45
            TabIndex        =   148
            Top             =   3630
            Width           =   5670
            Begin VB.TextBox txtConsultantRecomendation 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1395
               Left            =   45
               MaxLength       =   32500
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   149
               Top             =   225
               Width           =   5550
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Land Owner Comments"
            Height          =   1950
            Left            =   45
            TabIndex        =   147
            Top             =   1620
            Width           =   5670
            Begin FPSpreadADO.fpSpread spdDet1 
               Height          =   1635
               Left            =   60
               TabIndex        =   156
               Top             =   225
               Width           =   5520
               _Version        =   393216
               _ExtentX        =   9737
               _ExtentY        =   2884
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SpreadDesigner  =   "frmPozosFuturos.frx":D129
            End
         End
         Begin VB.TextBox txtConsult 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2340
            MaxLength       =   255
            TabIndex        =   66
            Top             =   225
            Width           =   3315
         End
         Begin VB.TextBox txtType 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2340
            MaxLength       =   255
            TabIndex        =   68
            Top             =   510
            Width           =   3315
         End
         Begin VB.TextBox txtEstado 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2385
            MaxLength       =   255
            TabIndex        =   76
            Top             =   7860
            Width           =   3285
         End
         Begin MSComCtl2.DTPicker dtpDMAFinalPermit 
            Height          =   375
            Left            =   2385
            TabIndex        =   74
            Top             =   7470
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin MSComCtl2.DTPicker dtpFechaPedidoETIA 
            Height          =   375
            Left            =   2340
            TabIndex        =   70
            Top             =   825
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin MSComCtl2.DTPicker dtpFechaEsperadaETIA 
            Height          =   375
            Left            =   2340
            TabIndex        =   72
            Top             =   1200
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin VB.Label lblFechaEsperadaETIA 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Esperada EIA"
            Height          =   195
            Left            =   525
            TabIndex        =   71
            Top             =   1290
            Width           =   1650
         End
         Begin VB.Label lblConsult 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Consult"
            Height          =   195
            Left            =   1545
            TabIndex        =   65
            Top             =   270
            Width           =   630
         End
         Begin VB.Label lblType 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            Height          =   195
            Left            =   1755
            TabIndex        =   67
            Top             =   555
            Width           =   420
         End
         Begin VB.Label lblFechaPedidoETIA 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Pedido EIA"
            Height          =   195
            Left            =   735
            TabIndex        =   69
            Top             =   915
            Width           =   1440
         End
         Begin VB.Label lblDMAFinalPermit 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "DMA Final Permit"
            Height          =   195
            Left            =   780
            TabIndex        =   73
            Top             =   7560
            Width           =   1440
         End
         Begin VB.Label lblEstado 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Estado"
            Height          =   195
            Left            =   1650
            TabIndex        =   75
            Top             =   7920
            Width           =   570
         End
      End
      Begin VB.Frame fraGrupo3 
         Caption         =   "Rigs Scheduler"
         Height          =   2895
         Left            =   -69420
         TabIndex        =   129
         Top             =   360
         Width           =   5775
         Begin VB.TextBox txtTD 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2355
            MaxLength       =   9
            TabIndex        =   50
            Top             =   210
            Width           =   3315
         End
         Begin VB.TextBox txtTotDays 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2355
            MaxLength       =   9
            TabIndex        =   52
            Top             =   495
            Width           =   3315
         End
         Begin VB.TextBox txtRemDays 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2355
            MaxLength       =   9
            TabIndex        =   54
            Top             =   780
            Width           =   3315
         End
         Begin VB.TextBox txtStatus 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2355
            MaxLength       =   255
            TabIndex        =   56
            Top             =   1065
            Width           =   3315
         End
         Begin VB.TextBox txtLandOwner 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2355
            MaxLength       =   255
            TabIndex        =   62
            Top             =   2085
            Width           =   3345
         End
         Begin MSComCtl2.DTPicker dtpStartDate 
            Height          =   375
            Left            =   2355
            TabIndex        =   58
            Top             =   1350
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin MSComCtl2.DTPicker dtpEndDate 
            Height          =   375
            Left            =   2355
            TabIndex        =   60
            Top             =   1725
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin MSComCtl2.DTPicker dtpLandOwnerPermitDate 
            Height          =   375
            Left            =   2355
            TabIndex        =   64
            Top             =   2370
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin VB.Label lblTD 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "TD"
            Height          =   195
            Left            =   1965
            TabIndex        =   49
            Top             =   255
            Width           =   225
         End
         Begin VB.Label lblTotDays 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Tot. Days"
            Height          =   195
            Left            =   1395
            TabIndex        =   51
            Top             =   540
            Width           =   795
         End
         Begin VB.Label lblRemDays 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Rem. Days"
            Height          =   195
            Left            =   1290
            TabIndex        =   53
            Top             =   825
            Width           =   900
         End
         Begin VB.Label lblStatus 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            Height          =   195
            Left            =   1635
            TabIndex        =   55
            Top             =   1110
            Width           =   555
         End
         Begin VB.Label lblStartDate 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date"
            Height          =   195
            Left            =   1305
            TabIndex        =   57
            Top             =   1440
            Width           =   885
         End
         Begin VB.Label lblEndDate 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "End Date"
            Height          =   195
            Left            =   1440
            TabIndex        =   59
            Top             =   1815
            Width           =   750
         End
         Begin VB.Label lblLandOwner 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Land Owner"
            Height          =   195
            Left            =   1200
            TabIndex        =   61
            Top             =   2130
            Width           =   990
         End
         Begin VB.Label lblLandOwnerPermitDate 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Land Owner Permit Date"
            Height          =   195
            Left            =   135
            TabIndex        =   63
            Top             =   2460
            Width           =   2055
         End
      End
      Begin VB.Frame fraGrupo2 
         Caption         =   "Plan de 8 y 13"
         Height          =   4215
         Left            =   -74865
         TabIndex        =   128
         Top             =   7290
         Width           =   5355
         Begin VB.TextBox txtFieldManifold 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1710
            MaxLength       =   50
            TabIndex        =   45
            Top             =   2460
            Width           =   3555
         End
         Begin VB.TextBox txtBatteryAssigned 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1710
            MaxLength       =   50
            TabIndex        =   46
            Top             =   2865
            Width           =   3555
         End
         Begin VB.TextBox txtFirstProd 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1710
            MaxLength       =   255
            TabIndex        =   136
            TabStop         =   0   'False
            Top             =   3780
            Width           =   3555
         End
         Begin VB.CheckBox chkPrognosis 
            Alignment       =   1  'Right Justify
            Caption         =   "Plan de 8"
            Height          =   255
            Left            =   1920
            TabIndex        =   36
            Top             =   540
            Width           =   1095
         End
         Begin VB.CheckBox chkAcuifero 
            Alignment       =   1  'Right Justify
            Caption         =   "Acuífero"
            Height          =   255
            Left            =   4110
            TabIndex        =   35
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkMonografia 
            Alignment       =   1  'Right Justify
            Caption         =   "Monografia"
            Height          =   255
            Left            =   1680
            TabIndex        =   34
            Top             =   255
            Width           =   1335
         End
         Begin VB.CheckBox chkPrograma 
            Alignment       =   1  'Right Justify
            Caption         =   "Plan de 13"
            Height          =   255
            Left            =   4005
            TabIndex        =   37
            Top             =   540
            Width           =   1215
         End
         Begin VB.TextBox txtDocumentoAPreparar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1710
            MaxLength       =   255
            TabIndex        =   43
            Top             =   2055
            Width           =   3555
         End
         Begin MSComCtl2.DTPicker dtpFechaSolicitud 
            Height          =   375
            Left            =   1710
            TabIndex        =   39
            Top             =   990
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin MSComCtl2.DTPicker dtpFechaPrioridad 
            Height          =   375
            Left            =   1710
            TabIndex        =   41
            Top             =   1515
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin MSComCtl2.DTPicker dtpFechaEntregaDictamenTecnico 
            Height          =   375
            Left            =   1710
            TabIndex        =   47
            Top             =   3270
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin VB.Label lblFieldManifold 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Field Manifold"
            Height          =   195
            Left            =   90
            TabIndex        =   138
            Top             =   2505
            Width           =   1545
         End
         Begin VB.Label lblBatteryAssigned 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Battery Assigned"
            Height          =   195
            Left            =   90
            TabIndex        =   137
            Top             =   2895
            Width           =   1545
         End
         Begin VB.Label lblFechaSolicitud 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Solicitud"
            Height          =   195
            Left            =   90
            TabIndex        =   38
            Top             =   1065
            Width           =   1545
         End
         Begin VB.Label lblFechaPrioridad 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Prioridad"
            Height          =   195
            Left            =   90
            TabIndex        =   40
            Top             =   1620
            Width           =   1545
         End
         Begin VB.Label lblDocumentoAPreparar 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Doc. a Preparar"
            Height          =   195
            Left            =   90
            TabIndex        =   42
            Top             =   2100
            Width           =   1545
         End
         Begin VB.Label lblFechaEntregaDictamenTecnico 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Entrega Dictamen"
            Height          =   195
            Left            =   90
            TabIndex        =   44
            Top             =   3360
            Width           =   1545
         End
         Begin VB.Label lblFirstProd 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "First Prod."
            Height          =   195
            Left            =   90
            TabIndex        =   48
            Top             =   3840
            Width           =   1545
         End
      End
      Begin VB.Frame fraGrupo1 
         Height          =   6840
         Left            =   -74865
         TabIndex        =   127
         Top             =   360
         Width           =   5355
         Begin VB.CommandButton cmdAnterior 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1350
            Picture         =   "frmPozosFuturos.frx":D2FD
            Style           =   1  'Graphical
            TabIndex        =   163
            Top             =   180
            Width           =   375
         End
         Begin VB.TextBox txtIDpozo 
            Enabled         =   0   'False
            Height          =   405
            Left            =   120
            TabIndex        =   151
            Text            =   "IDpozo"
            Top             =   1860
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox txtArea 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1710
            MaxLength       =   255
            TabIndex        =   146
            Top             =   1305
            Width           =   3525
         End
         Begin VB.TextBox txtWellInformed 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1710
            MaxLength       =   255
            TabIndex        =   145
            Top             =   4215
            Width           =   3540
         End
         Begin VB.TextBox txtEquipo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1710
            MaxLength       =   255
            TabIndex        =   9
            Top             =   990
            Width           =   3525
         End
         Begin VB.CommandButton cmdSiguiente 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4830
            Picture         =   "frmPozosFuturos.frx":D687
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   165
            Width           =   375
         End
         Begin VB.CheckBox chkDefinitiva 
            Alignment       =   1  'Right Justify
            Caption         =   "Definitiva"
            Height          =   255
            Left            =   1770
            TabIndex        =   25
            Top             =   5010
            Width           =   1215
         End
         Begin VB.ComboBox cmbUbicacion 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmPozosFuturos.frx":DA11
            Left            =   1710
            List            =   "frmPozosFuturos.frx":DA13
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   630
            Width           =   3525
         End
         Begin VB.TextBox txtWellID 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   1740
            MaxLength       =   255
            TabIndex        =   4
            Top             =   165
            Width           =   3075
         End
         Begin VB.TextBox txtX_PDC 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1725
            MaxLength       =   10
            TabIndex        =   27
            Top             =   5355
            Width           =   3525
         End
         Begin VB.TextBox txtY_PDC 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1725
            MaxLength       =   10
            TabIndex        =   29
            Top             =   5715
            Width           =   3525
         End
         Begin VB.TextBox txtX_Pos94 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1725
            MaxLength       =   10
            TabIndex        =   31
            Top             =   6060
            Width           =   3525
         End
         Begin VB.TextBox txtY_Pos94 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1725
            MaxLength       =   10
            TabIndex        =   33
            Top             =   6405
            Width           =   3525
         End
         Begin VB.TextBox txtInformedBy 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1710
            MaxLength       =   255
            TabIndex        =   24
            Top             =   4590
            Width           =   3540
         End
         Begin VB.TextBox txtProspect 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1725
            MaxLength       =   255
            TabIndex        =   15
            Top             =   2535
            Width           =   3525
         End
         Begin VB.OptionButton optExploratorio 
            Caption         =   "Exploratorio"
            Height          =   255
            Left            =   1725
            TabIndex        =   12
            Top             =   1965
            Width           =   1335
         End
         Begin VB.OptionButton optAvanzada 
            Caption         =   "Avanzada"
            Height          =   255
            Left            =   1725
            TabIndex        =   13
            Top             =   2205
            Width           =   1215
         End
         Begin VB.OptionButton optDesarrollo 
            Caption         =   "Desarrollo"
            Height          =   255
            Left            =   1725
            TabIndex        =   11
            Top             =   1680
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.TextBox txtMonografias 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1725
            MaxLength       =   9
            TabIndex        =   19
            Top             =   3405
            Width           =   3525
         End
         Begin MSComCtl2.DTPicker dtpPrimerMonografia 
            Height          =   375
            Left            =   1740
            TabIndex        =   17
            Top             =   2940
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   41234
         End
         Begin MSComCtl2.DTPicker dtpUltimaMonografia 
            Height          =   375
            Left            =   1725
            TabIndex        =   21
            Top             =   3765
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin VB.Label lblEquipo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Equipo"
            Height          =   195
            Left            =   90
            TabIndex        =   8
            Top             =   990
            Width           =   1575
         End
         Begin VB.Label lblWellInformed 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Well Informed"
            Height          =   195
            Left            =   90
            TabIndex        =   22
            Top             =   4260
            Width           =   1575
         End
         Begin VB.Label lblUbicacion 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Ubicacion"
            Height          =   195
            Left            =   90
            TabIndex        =   6
            Top             =   645
            Width           =   1575
         End
         Begin VB.Label lblWellID 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Well Id."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   90
            TabIndex        =   3
            Top             =   315
            Width           =   1575
         End
         Begin VB.Label lblX_PDC 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "X PDC"
            Height          =   195
            Left            =   90
            TabIndex        =   26
            Top             =   5400
            Width           =   1575
         End
         Begin VB.Label lblY_PDC 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Y PDC"
            Height          =   195
            Left            =   90
            TabIndex        =   28
            Top             =   5760
            Width           =   1575
         End
         Begin VB.Label lblX_Pos94 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "X Pos94"
            Height          =   195
            Left            =   90
            TabIndex        =   30
            Top             =   6105
            Width           =   1575
         End
         Begin VB.Label lblY_Pos94 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Y Pos94"
            Height          =   195
            Left            =   90
            TabIndex        =   32
            Top             =   6450
            Width           =   1575
         End
         Begin VB.Label lblInformedBy 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Informed By"
            Height          =   195
            Left            =   90
            TabIndex        =   23
            Top             =   4665
            Width           =   1575
         End
         Begin VB.Label lblProspect 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Prospect"
            Height          =   195
            Left            =   90
            TabIndex        =   14
            Top             =   2580
            Width           =   1575
         End
         Begin VB.Label lblPrimerMonografia 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Primer Monografia"
            Height          =   195
            Left            =   90
            TabIndex        =   16
            Top             =   3030
            Width           =   1575
         End
         Begin VB.Label lblUltimaMonografia 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Ultima Monografia"
            Height          =   195
            Left            =   90
            TabIndex        =   20
            Top             =   3855
            Width           =   1575
         End
         Begin VB.Label lblMonografias 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Monografias"
            Height          =   195
            Left            =   90
            TabIndex        =   18
            Top             =   3450
            Width           =   1575
         End
         Begin VB.Label lblYacimiento 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Area"
            Height          =   195
            Left            =   90
            TabIndex        =   10
            Top             =   1320
            Width           =   405
         End
      End
      Begin MSComDlg.CommonDialog cmd 
         Left            =   16080
         Top             =   9120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame fraGrupo5 
         Caption         =   "Estado"
         Height          =   8235
         Left            =   -63570
         TabIndex        =   131
         Top             =   360
         Width           =   5565
         Begin VB.TextBox txtInformeEvaluacionArqueologico 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2955
            MaxLength       =   255
            TabIndex        =   110
            Top             =   7800
            Width           =   2520
         End
         Begin VB.TextBox txtTasaContralor 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2970
            MaxLength       =   9
            TabIndex        =   99
            Top             =   4785
            Width           =   2520
         End
         Begin VB.TextBox txtTasaAdministrativa 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2970
            MaxLength       =   9
            TabIndex        =   97
            Top             =   4440
            Width           =   2520
         End
         Begin VB.TextBox txtTechnicalReport 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2970
            MaxLength       =   255
            TabIndex        =   95
            Top             =   4065
            Width           =   2520
         End
         Begin VB.TextBox txtAdenda 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2970
            MaxLength       =   255
            TabIndex        =   103
            Top             =   5535
            Width           =   2520
         End
         Begin VB.TextBox txtEstudio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2970
            MaxLength       =   9
            TabIndex        =   101
            Top             =   5160
            Width           =   2520
         End
         Begin VB.CheckBox chkEIAPresentado 
            Alignment       =   1  'Right Justify
            Caption         =   "EIA Presentado"
            Height          =   255
            Left            =   2940
            TabIndex        =   83
            Top             =   1545
            Width           =   1695
         End
         Begin VB.TextBox txtIDManifiesto 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2970
            MaxLength       =   255
            TabIndex        =   78
            Top             =   240
            Width           =   2520
         End
         Begin MSComCtl2.DTPicker dtpFechaManifiesto 
            Height          =   375
            Left            =   2970
            TabIndex        =   80
            Top             =   600
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin MSComCtl2.DTPicker dtpFechaEntregaEiaXConsultoraAOxy 
            Height          =   375
            Left            =   2970
            TabIndex        =   82
            Top             =   1035
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin MSComCtl2.DTPicker dtpFechaPresentacionDMA 
            Height          =   375
            Left            =   2970
            TabIndex        =   87
            Top             =   2355
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin MSComCtl2.DTPicker dtpFechaEnvioACS 
            Height          =   375
            Left            =   2970
            TabIndex        =   85
            Top             =   1920
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin MSComCtl2.DTPicker dtpFechaPresentacionSMA 
            Height          =   375
            Left            =   2970
            TabIndex        =   89
            Top             =   2790
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin MSComCtl2.DTPicker dtpPagoTasaAdministrativa 
            Height          =   375
            Left            =   2970
            TabIndex        =   91
            Top             =   3225
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin MSComCtl2.DTPicker dtpFechaInfoComplementaria 
            Height          =   375
            Left            =   2970
            TabIndex        =   93
            Top             =   3660
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   41233
         End
         Begin MSComCtl2.DTPicker dtpInformeAvanceObra50PorCiento 
            Height          =   375
            Left            =   2955
            TabIndex        =   106
            Top             =   6900
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   21168129
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker dtpInformeAvanceObra100PorCiento 
            Height          =   375
            Left            =   2955
            TabIndex        =   108
            Top             =   7365
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin MSComCtl2.DTPicker dtpFECHAINICIODIA 
            Height          =   375
            Left            =   2970
            TabIndex        =   140
            Top             =   5910
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin MSComCtl2.DTPicker dtpFECHAFINDIA 
            Height          =   375
            Left            =   2970
            TabIndex        =   141
            Top             =   6375
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   21168129
            CurrentDate     =   39804
         End
         Begin VB.Label lblFECHAFINDIA 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Fin DIA"
            Height          =   195
            Left            =   90
            TabIndex        =   139
            Top             =   6495
            Width           =   2955
         End
         Begin VB.Label lblInformeEvaluacionArqueologico 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Informe Evaluacion Arqueologico"
            Height          =   195
            Left            =   90
            TabIndex        =   109
            Top             =   7845
            Width           =   2955
         End
         Begin VB.Label lblInformeAvanceObra100PorCiento 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Informe Avance de Obra 100%"
            Height          =   195
            Left            =   90
            TabIndex        =   107
            Top             =   7455
            Width           =   2955
         End
         Begin VB.Label lblInformeAvanceObra50PorCiento 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Informe Avance de Obra 50%"
            Height          =   195
            Left            =   90
            TabIndex        =   105
            Top             =   6990
            Width           =   2955
         End
         Begin VB.Label lblTasaContralor 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Tasa Contralor ($)"
            Height          =   195
            Left            =   90
            TabIndex        =   98
            Top             =   4800
            Width           =   2955
         End
         Begin VB.Label lblTasaAdministrativa 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Tasa Administrativa ($)"
            Height          =   195
            Left            =   90
            TabIndex        =   96
            Top             =   4485
            Width           =   2955
         End
         Begin VB.Label lblTechnicalReport 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Technical Report"
            Height          =   195
            Left            =   90
            TabIndex        =   94
            Top             =   4110
            Width           =   2955
         End
         Begin VB.Label lblAdenda 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Adenda ($)"
            Height          =   195
            Left            =   90
            TabIndex        =   102
            Top             =   5580
            Width           =   2955
         End
         Begin VB.Label lblEstudio 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Estudio ($)"
            Height          =   195
            Left            =   90
            TabIndex        =   100
            Top             =   5205
            Width           =   2955
         End
         Begin VB.Label lblFECHAINICIODIA 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Inicio DIA"
            Height          =   195
            Left            =   90
            TabIndex        =   104
            Top             =   6030
            Width           =   2955
         End
         Begin VB.Label lblFechaEnvioACS 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Envio a CS"
            Height          =   195
            Left            =   90
            TabIndex        =   84
            Top             =   2010
            Width           =   2955
         End
         Begin VB.Label lblFechaPresentacionDMA 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Presentacion DMA"
            Height          =   195
            Left            =   90
            TabIndex        =   86
            Top             =   2445
            Width           =   2955
         End
         Begin VB.Label lblFechaPresentacionSMA 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Presentacion SMA"
            Height          =   195
            Left            =   90
            TabIndex        =   88
            Top             =   2880
            Width           =   2955
         End
         Begin VB.Label lblPagoTasaAdministrativa 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Pago Tasa Administrat."
            Height          =   195
            Left            =   90
            TabIndex        =   90
            Top             =   3315
            Width           =   2955
         End
         Begin VB.Label lblFechaInfoComplementaria 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Info Complementaria"
            Height          =   195
            Left            =   90
            TabIndex        =   92
            Top             =   3750
            Width           =   2955
         End
         Begin VB.Label lblFechaManifiesto 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Manifiesto"
            Height          =   195
            Left            =   90
            TabIndex        =   79
            Top             =   690
            Width           =   2955
         End
         Begin VB.Label lblFechaEntregaEiaXConsultoraAOxy 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Entrega x Consult. a OXY"
            Height          =   195
            Left            =   90
            TabIndex        =   81
            Top             =   1125
            Width           =   2955
         End
         Begin VB.Label lblIDManifiesto 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "ID Manifiesto"
            Height          =   195
            Left            =   90
            TabIndex        =   77
            Top             =   285
            Width           =   2950
         End
      End
   End
   Begin VB.Menu Grilla 
      Caption         =   "Grilla Header"
      Visible         =   0   'False
      Begin VB.Menu Nueva_fila_detalle 
         Caption         =   "Nueva Fila"
      End
      Begin VB.Menu Editar_fila_detalle 
         Caption         =   "Editar Fila"
      End
      Begin VB.Menu Eliminar_fila_detalle 
         Caption         =   "Eliminar Fila"
      End
   End
   Begin VB.Menu VariosPozos 
      Caption         =   "Grilla Header Actualizo varios Pozos"
      Visible         =   0   'False
      Begin VB.Menu Editar_pozos_seleccionados 
         Caption         =   "Seleccionar Pozos para su Modificación"
      End
   End
End
Attribute VB_Name = "frmPozosFuturos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''
'ABM de pozos'
''''''''''''''
Option Explicit


'Para guardar el filtro por Ubicacion, debe ser global
'para que se pueda concatenar con el filtro x columna
Dim strWhereUbicacion As String


'Para guardar fila y columna antes que pierda
'el poco la grilla luego poder recuperarlos
Dim lngFila, lngColumna As Long


'Para controlar si paso por evento active la primera vez
Dim lngPrimeraVez As Long


'Para controlar si paso por evento active la primera vez
Dim blnPintaFilas As Boolean


'Para definir limites cuando se seleccionan muchos pozos y se
'los quiere editar a todos juntos en la solapa Ficha de Pozo
Dim lngFilaDesde, lngFilaHasta, lngFilaActual As Long


Dim strSQL As String 'Query para buscar los contratistas
Dim arrCampos(59) As String 'Campos de los filtros en el mismo orden que aparecen en el combo
Dim IDModificacion As Long 'ID del registro que se esta modificando


'Constantes de los menu de las listas de comments
Const mCommentNuevo As Long = 0
Const mCommentEditar As Long = 1
Const mCommentEliminar As Long = 2



Private Enum enmMenuGrilla
  mEditar
End Enum

Dim blnTree As Boolean

Dim ColumnaSeleccionada As Long 'Columna que se elige al mostrar el menu para editar multiples filas a la vez

Private Enum enmSolapas 'Solapas del control SSTab
  enmSolapaPozos
  enmSolapaFichaPozos
End Enum


'cambiar apariencia a grilla
Public Property Get dsiCambiaApariencia(spd As fpSpread) As Boolean
  
  spd.UnitType = UnitTypeTwips                  'trabajar en twips

  spd.Appearance = AppearanceFlat               'apariencia 3D
  spd.BorderStyle = BorderStyleNone             'tipo de borde: sin borde

  spd.ColHeadersAutoText = DispBlank            'titulos de columnas en blanco
  spd.ColHeadersShow = True                     'muestra encabezado columnas
  spd.RowHeadersShow = True                     'muestra encabezado de filas

  spd.CursorStyle = CursorStyleArrow            'stilo cursor
  spd.CursorType = CursorTypeDefault            'tipo cursor

'  spd.AutoSize = True                           'automaticamente ajusta ancho grilla
  spd.DAutoSizeCols = DAutoSizeColsMax          'tipo de ajuste 2: al dato mas ancho

  spd.UserColAction = UserColActionDefault      'cuando hace click en header pinta columna o fila
  spd.FontSize = 9                              'tamaño letra
  spd.RowHeight(0) = 450                        'altura fila de titulos
  spd.MoveActiveOnFocus = False                 '
  spd.Protect = False                           'exporta a excel sin proteccion

  spd.BackColorStyle = BackColorStyleUnderGrid  'estilo
  spd.GridShowHoriz = True                      'muestra grilla horizontal
  spd.GridShowVert = True                       'muestra grilla vertical
  spd.GridColor = RGB(200, 200, 200)            'color muy suave
  spd.NoBorder = True                           'sin borde fin zona de datos

  spd.ScrollBars = ScrollBarsBoth               'ambas barras de desplazamiento
  spd.ScrollBarExtMode = False                  'cuando sean necesarias
  spd.VScrollSpecial = False                     'barra especial

  spd.SetOddEvenRowColor RGB(245, 245, 245), RGB(60, 60, 60), RGB(245, 245, 245), RGB(60, 60, 60)
  spd.SelBackColor = RGB(204, 230, 255)         'fondo del area seleccionada
  spd.GrayAreaBackColor = RGB(245, 245, 245)

  spd.VirtualMode = False                           'ajusta rows al tamaño del recordset
  'spd.VirtualRows = 300                            'rows a leer del virtual buffer
  'spd.VirtualScrollBuffer = True                   'scroll vertical lee de tantas rows del buffer

  'setea para mostrar tooltip en las celdas donde no se ve toda la info
  spd.TextTip = TextTipFixedFocusOnly
  spd.TextTipDelay = 250

  spd.TypeDateFormat = TypeDateFormatDDMMYY     'formato fecha
  spd.TypeDateCentury = True                    '4 digitos para el año
  
  spdCab.EditMode = True                        'grilla editable
  spdCab.EditModeReplace = True                 'No hay que eliminar el valor y volver a escribir, reemplaza
                                                
  
End Property


'cambiar apariencia a grilla
Public Property Get dsiCambiaAparienciaDet(spd As fpSpread) As Boolean
  
  spd.UnitType = UnitTypeTwips                  'trabajar en twips
  
  spd.Appearance = AppearanceFlat               'apariencia 3D
  spd.BorderStyle = BorderStyleNone             'tipo de borde: sin borde
  
  spd.ColHeadersAutoText = DispBlank            'titulos de columnas en blanco
  spd.ColHeadersShow = True                     'muestra encabezado columnas
  spd.RowHeadersShow = False                     'muestra encabezado de filas
  
  spd.CursorStyle = CursorStyleArrow            'stilo cursor
  spd.CursorType = CursorTypeDefault            'tipo cursor
  
  spd.AutoSize = False                          'automaticamente ajusta ancho grilla
  spd.DAutoSizeCols = DAutoSizeColsMax          'tipo de ajuste 2: al dato mas ancho

  spd.UserColAction = UserColActionDefault      'cuando hace click en header pinta columna o fila
  spd.FontSize = 9                              'tamaño letra
  spd.RowHeight(0) = 300                        'altura fila de titulos
  spd.MoveActiveOnFocus = False                 '
  spd.Protect = False                           'exporta a excel sin proteccion
  
  spd.BackColorStyle = BackColorStyleUnderGrid  'estilo
  spd.GridShowHoriz = True                      'muestra grilla horizontal
  spd.GridShowVert = True                       'muestra grilla vertical
  spd.GridColor = RGB(200, 200, 200)            'color muy suave
  spd.NoBorder = True                           'sin borde fin zona de datos
  
  spd.ScrollBars = ScrollBarsVertical           'ambas barras de desplazamiento
  spd.ScrollBarExtMode = True                   'cuando sean necesarias
  spd.VScrollSpecial = False                     'barra especial
      
  spd.SetOddEvenRowColor RGB(245, 245, 245), RGB(60, 60, 60), RGB(245, 245, 245), RGB(60, 60, 60)
  spd.SelBackColor = RGB(204, 230, 255)         'fondo del area seleccionada
  spd.GrayAreaBackColor = RGB(245, 245, 245)
  
  spd.VirtualMode = False                        ' ajusta rows al tamaño del recordset
  'spd.VirtualRows = 300                            ' rows a leer del virtual buffer
  'spd.VirtualScrollBuffer = True                   ' scroll vertical lee de tantas rows del buffer
  
  'setea para mostrar tooltip en las celdas donde no se ve toda la info
  spd.TextTip = TextTipFixedFocusOnly
  spd.TextTipDelay = 250
  
End Property


Private Sub GET_datos(strWhere As String, Optional blnFillLista As Boolean)

  Dim rs As ADODB.Recordset
  Dim intI As Integer
  Dim strT  As String
  Dim blnB  As Boolean
  Dim lngL, lngIni, lngFin  As Long
  Dim varDato As Variant
  
  
 'puntero mouse reloj
  Screen.MousePointer = vbHourglass
  
 
  strT = "select p.* " & _
         "from EIApozosDatos_vw p"
  
  
  'CHECK si existe where
  If strWhere <> "" Then
    
    'ADD order by a Query
    strT = strT & " where " & strWhere
    
  End If
  
  'ADD order by a Query
  strT = strT & " order by OrdenEnGrilla"
  
  'EXEC query
  Set rs = SQLexec(strT)
  
  
  'CHECK error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    End
  End If
  
  'DISABLE  para que la grilla trabaje en background Lo que haces
  '         es que no se vean los refresh cuando se realizan cambios
  Me.spdCab.Redraw = False
  
  'CLEAR vacio grilla para no tener problemas con celdas con colores
  Me.spdCab.MaxRows = 0
  
  'BINDING grilla
  Set Me.spdCab.DataSource = rs
  
  'SET limite en grilla
  Me.spdCab.MaxRows = rs.RecordCount
  Me.spdCab.MaxCols = rs.Fields.Count
  
  'ENABLE   para que la grilla trabaje en background Lo que haces
  '         es que no se vean los refresh cuando se realizan cambios
  Me.spdCab.Redraw = True
  
  
  'INMOVILIZAR columnas
'  Me.spdCab.ColsFrozen = 1
  
  'SET ID de columna para luego poder extraer el dato x ID
  For intI = 0 To rs.Fields.Count - 1
    
    spdCab.Col = intI + 1
    spdCab.ColID = rs.Fields(intI).Name
    
  Next
  
    
  
  'CHECK si debe llenar la lista, por primera vez siempre se debe llenar
  If blnFillLista Then
  
    'CLEAR lista desplegable
    Me.cmb.Clear
    
    'FILL lista desplegable de Campos
    For intI = 0 To rs.Fields.Count - 1
      
      'CHECK si es BIT o NTEXT para marcarlo porque no se puede aplicar un Select Max
      Me.cmb.AddItem rs.Fields(intI).Name & " - (" & rs.Fields(intI).Type & ")"
      
    Next
  
  End If
  
'  'SHOW en barra de estado
  frmMenuPrincipal.info.Panels(3) = " fila: " & Me.spdCab.ActiveRow & " de " & Me.spdCab.MaxRows & " , columna: " & Me.spdCab.ActiveCol & " de " & Me.spdCab.MaxCols & " , valor: " & Me.spdCab.Text & " "
  frmMenuPrincipal.info.Panels(3).ToolTipText = " fila: " & Me.spdCab.ActiveRow & " de " & Me.spdCab.MaxRows & " , columna: " & Me.spdCab.ActiveCol & " de " & Me.spdCab.MaxCols & " , valor: " & Me.spdCab.Text & " "
  
  'recupero puntero mouse
  Screen.MousePointer = vbDefault

End Sub



Private Sub chkNulls_Click()


End Sub


Private Sub cmb_Click()
  
  Dim rs As ADODB.Recordset
  Dim strT, strNombre, strTipo As String
  Dim blnB As Boolean
  
  'CHECK si selecciono item
  If Me.cmb.ListIndex = -1 Then
    Exit Sub
  End If
  
  'SPLIT nombre de columna y tipo de datos, elimina parentesis
  strNombre = Left(Me.cmb, InStr(Me.cmb, "-") - 2)
  strTipo = Replace(Replace(Right(Me.cmb, Len(Me.cmb) - InStr(Me.cmb, "-")), "(", ""), ")", "")
  
  'CHECK si se puede consultar por máximo valor
  If strTipo = 11 Or strTipo = 203 Or strNombre = "Cant" Then
    
    'SHOW maximo valor
    frmMenuPrincipal.info.Panels(4) = " No aplica máximo valor "
    Exit Sub
    
  End If
  
  ' GET maximo valor de columna seleccionada
  '
  strT = "select MAX([" & strNombre & "]) as Valor " & _
         "from EIApozosDatos_vw p"
  
  Set rs = SQLexec(strT)
  
  'CHECK error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    Exit Sub
    End
  End If
  
  'CHECK si existe
  If Not rs.EOF Then
    
    'SHOW maximo valor
    frmMenuPrincipal.info.Panels(4) = " Máximo valor: " & rs(0) & " "
    
  End If

'  'SET foco en buscar
'  Me.txtDato.SetFocus

End Sub


Private Sub cmdAnterior_Click()
  
  Dim blnB As Boolean
  
  'CHECK si hay filas y no se llego al limite
  If Me.spdCab.ActiveRow > 0 And Me.spdCab.ActiveRow > lngFilaDesde Then
    
    'ASSIGN fila
    lngFila = Me.spdCab.ActiveRow - 1
    
    'SET celda activa
    Me.spdCab.SetActiveCell lngColumna, lngFila
    
    'CALL ficha de Pozo
    fichaPozo (lngFila)
    
  End If
   
End Sub

'Private Sub cmdActualizarTablaParaGIS_Click()
'BD.Execute "EXEC DELETE_POZOS_GIS"
'DoEvents
'BD.Execute "EXEC INSERT_POZOS_GIS"
''agregar codigo
'DoEvents
'Shell App.Path & "\LlamarModelBuilder.exe", vbNormalFocus
''MsgBox "Se actualizo la tabla para GIS satisfactoriamente", vbInformation, "Actualizacion Finalizada"
'End Sub

Private Sub cmdBuscar1_Click()
  
  Dim rs As ADODB.Recordset
  Dim strT, strNombre, strTipo, strValor, strCriterio, strWhere, strWhereAnd As String
  Dim blnB As Boolean
  
  'CHECK si selecciono item
  If Me.cmb.ListIndex = -1 Then
    blnB = MsgBox("Para poder buscar, debe seleccionar una columna.", vbCritical + vbOKOnly, "Atención...")
    Exit Sub
  End If
  
  
 'mouse reloj
  Screen.MousePointer = vbHourglass
  
  
  'SPLIT nombre de columna y tipo de datos, elimina parentesis
  strNombre = "[" & Left(Me.cmb, InStr(Me.cmb, "-") - 2) & "]"
  strTipo = Replace(Replace(Right(Me.cmb, Len(Me.cmb) - InStr(Me.cmb, "-")), "(", ""), ")", "")
  
  'CHECK si selecciono Cant, no permite buscar, columna que no existe, se genera para uso interno
  If strNombre = "Cant" Then
    Exit Sub
  End If
  
  'SET valor
  strValor = Me.txtDato
  
  'SET criterio estandar
  strCriterio = " = "
      
      
  'GENERATE valor para buscar
  'Si es texto:, agrega comillas simple, si es fecha: convierte a ISO, si es numero: nada
  Select Case strTipo
          
  'bit
  Case conBit
      
    strValor = IIf(strValor = "", "null", strValor)
      
  'numeros
  Case conSmallInt, conInt, conTinyInt, conMoney, conSmallMoney, conReal, conFloat, conNumeric, conDecimal
          
    strValor = IIf(strValor = "", "null", strValor)
    
  'fecha
  Case conSmallDateTime, conDateTime
      
    strValor = IIf(strValor = "", "null", "'" & dateToIso(strValor) & "'")
    
  'string
  Case conChar, conNchar, conVarchar, conText, conNVarchar, conNtext
      
    strValor = IIf(strValor = "", "null", "'%" & strValor & "%'")
    strCriterio = " like "
    
  End Select
  
  'CLEAR where
  strWhere = ""
  strWhereAnd = ""
  
  'CHECK si where ubicacion contiene algo
  If strWhereUbicacion <> "" Then
    
    strWhere = strWhereUbicacion
    strWhereAnd = " and "
    
  End If
  
  'CHECK si where ubicacion contiene algo
  If strValor <> "null" Then
    
    strWhere = strWhereUbicacion & strWhereAnd & strNombre & strCriterio & strValor
    
  End If
  
  'CALL get datos segun criterio de busquedam
    GET_datos (strWhere)
  
 
  'PAINT texto, esto es mas comodo para que no tenga que borrar el dato a buscar y sea reemplazado
  'Me.txtDato.Text = ""
  Me.txtDato.SelStart = 0
  Me.txtDato.SelLength = 50
  
  'CALL activate para que pinte celdas
  Call Form_Activate
    
  'SET foco en buscar
  Me.txtDato.SetFocus
  
  
  'mouse defa
  Screen.MousePointer = vbDefault
  
End Sub


Private Sub cmdColumnas_Click()

End Sub




Private Sub cmdDMA_Click()

  Dim blnB, blnRowPainted, blnFilaConMerge As Boolean
  Dim strT As String
  Dim lngRow, lngCol, lngColVisible As Long
  Dim varDato, varDatoColRig, varDatoColEstado As Variant
  Dim intI As Integer
  
  Dim varARR(11) As String
  
  Dim Obj As Object
  Dim Libro As Object
  Dim Hoja As Object
  
  'DEFINE columnas a exportar para el formato estandar
  varARR(0) = "Well ID"
  varARR(1) = "Ubicacion"
  varARR(2) = "Equipo"
  varARR(3) = "Status"
  varARR(4) = "Site Visit - Fecha"
  varARR(5) = "Site Visit - Acta"
  varARR(6) = "Site Visit - Comentario"
  varARR(7) = "DMA Final Permit"
  varARR(8) = "Estado"
  varARR(9) = "Presentación DMA"
  varARR(10) = "Presentación SMA"
  
  'mouse reloj
  Screen.MousePointer = vbHourglass
  
'  '-----------------------------------------------------------------------------
'  'APPLY filtro para exportacion estandar - DRILLING INVENTORY - RIGS SCHEDULER - GONE / DONE
'
'  Me.spdFiltro.SetText 1, 1, 1
'  Me.spdFiltro.SetText 2, 1, 1
'  Me.spdFiltro.SetText 3, 1, 1
'
'  'CALL filtrado de info
'  Call cmdFiltrar_Click
'  '-----------------------------------------------------------------------------
  
  
  'CREATE excel
  Set Obj = CreateObject("Excel.application")
  Set Libro = Obj.Workbooks.Add()
  Set Hoja = Libro.Sheets(1)
  
  
  'ENABLE excel visible
  Obj.Visible = False
  
  'WHILE grilla - filas
  For lngRow = 0 To Me.spdCab.DataRowCnt
    
    'GET  valor celda para columna ESTADO, para saber si va pintada
    lngCol = Me.spdCab.GetColFromID("Estado")
    Me.spdCab.GetText lngCol, lngRow, varDatoColEstado
    
    'GET  valor celda para columna NUMBER, para chequear
    '     si cada columna debe ser pintada en paso siguiente
    lngCol = Me.spdCab.GetColFromID("Number")
    Me.spdCab.GetText lngCol, lngRow, varDatoColRig
    
    'PAINT celda
    If varDatoColEstado = "APROBADO" Then
      Hoja.Cells(lngRow + 1, 9).Interior.Color = RGB(196, 215, 155)
    End If
    
    'WHILE array con columnas a exportar
    lngColVisible = 0
    
    'CLEAR flags para saber si se le debe hacer un merge a la fila
    blnFilaConMerge = False
    
    For intI = 0 To UBound(varARR) - 1
      
      'GET valor celda para columna RIGoficial, para chequear
      '     si Pozo pertenece a RIG oficial y va pintado
      lngCol = Me.spdCab.GetColFromID(varARR(intI))
      
      'CHECK si columna visible y se debe exportar
      Me.spdCab.Col = lngCol
      
      'ADD 1 a columnas a exportar
      lngColVisible = lngColVisible + 1
      
      'GET celda de grilla Spread
      Me.spdCab.GetText lngCol, lngRow, varDato
      
      'FORMAT si dato de tipo fecha, formateo dd/mm/yyyy
      If IsDate(varDato) And Len(varDato) = 10 Then
        Hoja.Cells(lngRow + 1, lngColVisible) = Format(varDato, "DD/MM/YYYY")
      Else
        Hoja.Cells(lngRow + 1, lngColVisible) = varDato
      End If
      
      'PAINT celda
      If varDatoColRig = "" Then
        Hoja.Cells(lngRow + 1, lngColVisible).Interior.Color = RGB(196, 215, 155)
        blnFilaConMerge = True
      End If
      
      'BORDER pinta borde gris oscuro
      Hoja.Cells(lngRow + 1, lngColVisible).Borders.Color = RGB(79, 79, 79)
      
      'CENTER celda
      Hoja.Cells(lngRow + 1, lngColVisible).HorizontalAlignment = xlVAlignCenter
      
    Next
    
    'CHECK si se debe aplicar merge a la fila
    If blnFilaConMerge Then
      strT = "A" & lngRow + 1 & ":I" & lngRow + 1
      Hoja.Range(strT).Merge
      Hoja.Cells(lngRow + 1, 1).HorizontalAlignment = xlHAlignLeft
    End If
    
  Next
  
  
  'ADJUST ancho de columna al mas ancho
  Obj.Cells.EntireColumn.AutoFit
  
  'FORMAT encabezado de filas
  lngCol = 1
  Do
    Hoja.Cells(1, lngCol).Font.ColorIndex = 2
    Hoja.Cells(1, lngCol).Font.Bold = True
    Hoja.Cells(1, lngCol).Interior.ColorIndex = 50
    Hoja.Rows(1).RowHeight = 20
    
    lngCol = lngCol + 1
  Loop Until Hoja.Cells(1, lngCol) = ""
  
  'ADJUST ancho de columna, le sumo 4 para que quede mas prolijo
  For lngCol = 1 To Me.spdCab.DataColCnt
  
    'ADJUST mientras no sea comentario
    If lngCol = 7 Then
        Hoja.Columns(lngCol).ColumnWidth = 80
        Hoja.Columns(lngCol).HorizontalAlignment = xlGeneral
        Hoja.Columns(lngCol).VerticalAlignment = xlCenter
        Hoja.Columns(lngCol).WrapText = True
    Else
        Hoja.Columns(lngCol).ColumnWidth = Hoja.Columns(lngCol).ColumnWidth + 4
    End If
  
  Next
  
  'ENABLE excel visible
  Obj.Visible = True
  
  'mouse defa
  Screen.MousePointer = vbDefault
  
  'CLOSE objetos
  Set Obj = Nothing
  Set Libro = Nothing
  Set Hoja = Nothing


End Sub

Private Sub cmdFiltrar_Click()


  Dim rs As ADODB.Recordset
  Dim intI As Integer
  Dim varDato, varColNombre As Variant
  Dim strColNombre, strT  As String
  
  'mouse reloj
  Screen.MousePointer = vbHourglass
  
  strWhereUbicacion = ""
    
  'BUILD where
  For intI = 1 To Me.spdFiltro.MaxCols
    
    'GET nombre de columna y tilde o no
    Me.spdFiltro.GetText intI, 0, varColNombre
    Me.spdFiltro.GetText intI, 1, varDato
    
    'CHECK si celda con tilde
    If varDato = 1 Then
        strWhereUbicacion = strWhereUbicacion & "'" & varColNombre & "', "
    End If
    
  Next
    
  'CHECK si existe where
  If strWhereUbicacion <> "" Then
    
    'BUILD filtro
    strWhereUbicacion = "Ubicacion in (" & Left(strWhereUbicacion, Len(strWhereUbicacion) - 2) & ")"
    
    'GET datos sin Where
    Call GET_datos(strWhereUbicacion)
      
  Else
  
    'GET datos sin Where
    Call GET_datos("")
  
  End If
    
  'CALL activate para que pinte celdas
  Call Form_Activate

  'SET foco en grilla
  Me.spdCab.SetFocus
  
  'mouse defa
  Screen.MousePointer = vbDefault

End Sub


Private Sub cmdPER_Click()
  
  
  Dim blnB, blnRowPainted, blnFilaConMerge As Boolean
  Dim strT As String
  Dim lngRow, lngCol, lngColVisible As Long
  Dim varDato, varDatoColRig, varDatoColRigOficial As Variant
  
  Dim Obj As Object
  Dim Libro As Object
  Dim Hoja As Object
  
  'mouse reloj
  Screen.MousePointer = vbHourglass
  
  'CREATE excel
  Set Obj = CreateObject("Excel.application")
  Set Libro = Obj.Workbooks.Add()
  Set Hoja = Libro.Sheets(1)
  
  'ENABLE excel visible
  Obj.Visible = False
  
  'WHILE grilla - filas
  For lngRow = 0 To Me.spdCab.DataRowCnt
    
    'GET  valor celda para columna NUMBER, para chequear
    '     si cada columna debe ser pintada en paso siguiente
    lngCol = Me.spdCab.GetColFromID("Number")
    Me.spdCab.GetText lngCol, lngRow, varDatoColRig
    
    'GET  valor celda para columna RIGoficial, para chequear
    '     si Pozo pertenece a RIG oficial y va pintado
    lngCol = Me.spdCab.GetColFromID("RIGoficial")
    Me.spdCab.GetText lngCol, lngRow, varDatoColRigOficial
    
    'PAINT celda
    If varDatoColRigOficial = "1" Then
      Hoja.Cells(lngRow + 1, 1).Interior.Color = RGB(196, 215, 155)
    End If
    
    'WHILE grilla - columnas
    lngColVisible = 0
    For lngCol = 1 To Me.spdCab.DataColCnt - 5
      
      'CHECK si columna visible y se debe exportar
      Me.spdCab.Col = lngCol
      If Not Me.spdCab.ColHidden Then
        
        'ADD 1 a columnas a exportar
        lngColVisible = lngColVisible + 1
        
        'GET celda de grilla Spread
        Me.spdCab.GetText lngCol, lngRow, varDato
        
        'FORMAT si dato de tipo fecha, formateo dd/mm/yyyy
        If IsDate(varDato) Then
          Hoja.Cells(lngRow + 1, lngColVisible) = Format(varDato, "DD/MM/YYYY")
        Else
          Hoja.Cells(lngRow + 1, lngColVisible) = varDato
        End If
        
        'CLEAR
        blnFilaConMerge = False
        
        'PAINT  celda
        '       Esto es para que pinte la fila completa para agrupar por Rigs Scheduler
        If varDatoColRig = "" Then
          Hoja.Cells(lngRow + 1, lngColVisible).Interior.Color = RGB(196, 215, 155)
          blnFilaConMerge = True
        End If
        
        'BORDER pinta borde gris oscuro
        Hoja.Cells(lngRow + 1, lngColVisible).Borders.Color = RGB(79, 79, 79)
        
        'CENTER celda
        Hoja.Cells(lngRow + 1, lngColVisible).HorizontalAlignment = xlVAlignCenter
      
      End If
      
    Next
    
    'CHECK si se debe aplicar merge a la fila
    If blnFilaConMerge Then
      strT = "A" & lngRow + 1 & ":I" & lngRow + 1
      Hoja.Range(strT).Merge
      Hoja.Cells(lngRow + 1, 1).HorizontalAlignment = xlHAlignLeft
    End If
    
  Next
  
  
  'ADJUST ancho de columna al mas ancho
  Obj.Cells.EntireColumn.AutoFit
  
  'FORMAT encabezado de filas
  lngCol = 1
  Do
    Hoja.Cells(1, lngCol).Font.ColorIndex = 2
    Hoja.Cells(1, lngCol).Font.Bold = True
    Hoja.Cells(1, lngCol).Interior.ColorIndex = 50
    Hoja.Rows(1).RowHeight = 20
    
    lngCol = lngCol + 1
  Loop Until Hoja.Cells(1, lngCol) = ""
  
  'ADJUST ancho de columna, le sumo 4 para que quede mas prolijo
  For lngCol = 1 To Me.spdCab.DataColCnt
    
    If Hoja.Cells(1, lngCol) <> "Site Visit - Comentario" Then
'      Hoja.Columns(lngCol).ColumnWidth = Hoja.Columns(lngCol).ColumnWidth + 4
    Else
'      Hoja.Columns(lngCol).ColumnWidth = 100
      Hoja.Columns(lngCol).WrapText = True
    End If
    
  Next
  
  'ENABLE excel visible
  Obj.Visible = True
  
  'mouse defa
  Screen.MousePointer = vbDefault
  
  'CLOSE objetos
  Set Obj = Nothing
  Set Libro = Nothing
  Set Hoja = Nothing

End Sub

  


Private Sub cmdSiguiente_Click()

  Dim blnB As Boolean
  
  'CHECK si hay filas y no se llego al limite
  If Me.spdCab.ActiveRow > 0 And Me.spdCab.ActiveRow < lngFilaHasta Then
    
    'ASSIGN fila
    lngFila = Me.spdCab.ActiveRow + 1
    
    'SET celda activa
    Me.spdCab.SetActiveCell lngColumna, lngFila
    
    'CALL ficha de Pozo
    fichaPozo (lngFila)
    
  End If

End Sub


Private Sub cmdSTD_Click()
  
  Dim blnB, blnRowPainted, blnFilaConMerge As Boolean
  Dim strT As String
  Dim lngRow, lngCol, lngColVisible As Long
  Dim varDato, varDatoColRig, varDatoColEstado As Variant
  Dim intI As Integer
  
  Dim varARR(9) As String
  
  Dim Obj As Object
  Dim Libro As Object
  Dim Hoja As Object
  
  'DEFINE columnas a exportar para el formato estandar
  varARR(0) = "Ubicacion"
  varARR(1) = "Well ID"
  varARR(2) = "Esperada EIA"
  varARR(3) = "DMA Final Permit"
  varARR(4) = "Estado"
  varARR(5) = "Area"
  varARR(6) = "Number"
  varARR(7) = "Field Manifold"
  varARR(8) = "Battery Assigned"
  
  'mouse reloj
  Screen.MousePointer = vbHourglass
  
  '-----------------------------------------------------------------------------
  'APPLY filtro para exportacion estandar - DRILLING INVENTORY - RIGS SCHEDULER - GONE / DONE
  
  Me.spdFiltro.SetText 1, 1, 1
  Me.spdFiltro.SetText 2, 1, 1
  Me.spdFiltro.SetText 3, 1, 1
  
  'CALL filtrado de info
  Call cmdFiltrar_Click
  '-----------------------------------------------------------------------------
  
  
  'CREATE excel
  Set Obj = CreateObject("Excel.application")
  Set Libro = Obj.Workbooks.Add()
  Set Hoja = Libro.Sheets(1)
  
  
  'ENABLE excel visible
  Obj.Visible = False
  
  'WHILE grilla - filas
  For lngRow = 0 To Me.spdCab.DataRowCnt
    
    'GET  valor celda para columna ESTADO, para saber si va pintada
    lngCol = Me.spdCab.GetColFromID("Estado")
    Me.spdCab.GetText lngCol, lngRow, varDatoColEstado
    
    'GET  valor celda para columna NUMBER, para chequear
    '     si cada columna debe ser pintada en paso siguiente
    lngCol = Me.spdCab.GetColFromID("Number")
    Me.spdCab.GetText lngCol, lngRow, varDatoColRig
    
    'PAINT celda
    If varDatoColEstado = "APROBADO" Then
      Hoja.Cells(lngRow + 1, 5).Interior.Color = RGB(196, 215, 155)
    End If
    
    'WHILE array con columnas a exportar
    lngColVisible = 0
    
    'CLEAR flags para saber si se le debe hacer un merge a la fila
    blnFilaConMerge = False
    
    For intI = 0 To UBound(varARR) - 1
      
      'GET valor celda para columna RIGoficial, para chequear
      '     si Pozo pertenece a RIG oficial y va pintado
      lngCol = Me.spdCab.GetColFromID(varARR(intI))
      
      'CHECK si columna visible y se debe exportar
      Me.spdCab.Col = lngCol
      
      'ADD 1 a columnas a exportar
      lngColVisible = lngColVisible + 1
      
      'GET celda de grilla Spread
      Me.spdCab.GetText lngCol, lngRow, varDato
      
      'FORMAT si dato de tipo fecha, formateo dd/mm/yyyy
      If IsDate(varDato) Then
        Hoja.Cells(lngRow + 1, lngColVisible) = Format(varDato, "DD/MM/YYYY")
      Else
        Hoja.Cells(lngRow + 1, lngColVisible) = varDato
      End If
      
      'PAINT celda
      If varDatoColRig = "" Then
        Hoja.Cells(lngRow + 1, lngColVisible).Interior.Color = RGB(196, 215, 155)
        blnFilaConMerge = True
      End If
      
      'BORDER pinta borde gris oscuro
      Hoja.Cells(lngRow + 1, lngColVisible).Borders.Color = RGB(79, 79, 79)
      
      'CENTER celda
      Hoja.Cells(lngRow + 1, lngColVisible).HorizontalAlignment = xlVAlignCenter
      
    Next
    
    'CHECK si se debe aplicar merge a la fila
    If blnFilaConMerge Then
      strT = "A" & lngRow + 1 & ":I" & lngRow + 1
      Hoja.Range(strT).Merge
      Hoja.Cells(lngRow + 1, 1).HorizontalAlignment = xlHAlignLeft
    End If
    
  Next
  
  
  'ADJUST ancho de columna al mas ancho
  Obj.Cells.EntireColumn.AutoFit
  
  'FORMAT encabezado de filas
  lngCol = 1
  Do
    Hoja.Cells(1, lngCol).Font.ColorIndex = 2
    Hoja.Cells(1, lngCol).Font.Bold = True
    Hoja.Cells(1, lngCol).Interior.ColorIndex = 50
    Hoja.Rows(1).RowHeight = 20
    
    lngCol = lngCol + 1
  Loop Until Hoja.Cells(1, lngCol) = ""
  
  'ADJUST ancho de columna, le sumo 4 para que quede mas prolijo
  For lngCol = 1 To Me.spdCab.DataColCnt
    Hoja.Columns(lngCol).ColumnWidth = Hoja.Columns(lngCol).ColumnWidth + 4
  Next
  
  'ENABLE excel visible
  Obj.Visible = True
  
  'mouse defa
  Screen.MousePointer = vbDefault
  
  'CLOSE objetos
  Set Obj = Nothing
  Set Libro = Nothing
  Set Hoja = Nothing


End Sub



Private Sub cmdTree_Click()
  
  'CHECK si no visible
  If Not Me.Tree.Visible Then
    
    'ENABLE tree view
    Me.FrameTree.Width = 4095
    Me.Tree.Visible = True
    
    'LAYOUT pantalla
    Me.Frame3.Left = FrameTree.Width + 100
    Me.Frame3.Width = frmMenuPrincipal.Width - (FrameTree.Width + 350)
    Me.spdCab.Width = Me.Frame3.Width - 180
    
  Else
    
    'DISABLE tree view
    Me.FrameTree.Width = 0
    Me.Tree.Visible = False
    
    'LAYOUT pantalla
    Me.Frame3.Left = FrameTree.Width + 100
    Me.Frame3.Width = frmMenuPrincipal.Width - (FrameTree.Width + 350)
    Me.spdCab.Width = Me.Frame3.Width - 180
    
  End If
  
End Sub


Private Sub Editar_fila_detalle_Click()

  'SET operacion
  frmEditorComments.dsiOpera = "U"
  
  'CALL formulario edición
  frmEditorComments.Show vbModal

End Sub

Private Sub Editar_pozos_seleccionados_Click()


  Dim c1, r1, r2 As Variant
  Dim blnB As Boolean
  
  'GET filas seleccionadas, uso la misma variable para las columnas, porque no interesan, solo las filas son importantes
  Me.spdCab.GetSelection 0, c1, r1, c1, r2
  
  'SAVE fila desde, fila hasta
  lngFilaDesde = r1
  lngFilaHasta = r2
  
  
End Sub

Private Sub Eliminar_fila_detalle_Click()

  'SET operacion
  frmEditorComments.dsiOpera = "D"
  
  'CALL formulario edición
  frmEditorComments.Show vbModal

  Me.spdDet2.DeleteRows Me.spdDet2.ActiveRow, 1

End Sub

Private Sub Form_Activate()
  
  
  Dim lngL, lngLANT, lngIni, lngFin, lngLimiteRigs, lngColUbi As Long
  Dim strT As String
  Dim varDato, varDatoMarca, varEquipo, varEquipoANT As Variant
  Dim intUno, intDos, intTres, intI As Integer
  Dim blnCambiaColor  As Boolean
  Dim rs As ADODB.Recordset
    
    
    'HIDE columnas que sirven para uso interno
    lngL = Me.spdCab.GetColFromID("IDpozo")
    Me.spdCab.Col = lngL
    Me.spdCab.ColHidden = True
    lngL = Me.spdCab.GetColFromID("EIAIDpozo")
    Me.spdCab.Col = lngL
    Me.spdCab.ColHidden = True
    lngL = Me.spdCab.GetColFromID("OrdenEnGrilla")
    Me.spdCab.Col = lngL
    Me.spdCab.ColHidden = True
    lngL = Me.spdCab.GetColFromID("Ubi")
    Me.spdCab.Col = lngL
    lngColUbi = lngL
    Me.spdCab.ColHidden = True
    lngL = Me.spdCab.GetColFromID("RIGoficial")
    Me.spdCab.Col = lngL
    Me.spdCab.ColHidden = True
    
    
    'FIND DRILLING INVENTORY
    lngL = Me.spdCab.SearchCol(lngColUbi, 0, -1, "2", SearchFlagsCaseSensitive)
    
    'SAVE cuando encuentro DRILLING INVENTORY, para saber el límite de RIGS SCHEDULER
    lngLimiteRigs = IIf(lngL = -1, 0, lngL)
    
    'CALL pinta fila
    Call pinta_Fila(lngL, "DRILLING INVENTORY")
    
    'FIND GONE / DONE
    lngL = Me.spdCab.SearchCol(lngColUbi, 0, -1, "3", SearchFlagsCaseSensitive)
    
    'CALL pinta fila
    Call pinta_Fila(lngL, "GONE / DONE")
    
    'FIND DISCHARGES
    lngL = Me.spdCab.SearchCol(lngColUbi, 0, -1, "4", SearchFlagsCaseSensitive)
    
    'CALL pinta fila
    Call pinta_Fila(lngL, "DISCHARGES")
    
    'FIND SIN UBICACION
    lngL = Me.spdCab.SearchCol(lngColUbi, 0, -1, "5", SearchFlagsCaseSensitive)
    
    'CALL pinta fila
    Call pinta_Fila(lngL, "SIN UBICACION")
   
    
    'FIND RIGS SCHEDULE
    lngL = Me.spdCab.SearchCol(3, 0, -1, "RIGS SCHEDULE", SearchFlagsCaseSensitive)
    
    'CHECK so encontro
    If lngL <> -1 Then
      
      'GET equipo
      varEquipoANT = ""
      lngLANT = 0

      'WHILE grilla
      For lngL = 1 To Me.spdCab.DataRowCnt

        'GET ubicacion
        Me.spdCab.GetText 3, lngL, varDato
        Me.spdCab.GetText Me.spdCab.DataColCnt, lngL, varDatoMarca

        'GET equipo
        Me.spdCab.GetText 4, lngL, varEquipo

        'CHECK si <> de RIG SCHEDULE, corto iteración para acelerar proceso
        If varDato <> "RIGS SCHEDULE" Then
          Exit For
        End If
        
        
        'CHECK si <> equipo
        If varEquipo <> varEquipoANT Then
          varEquipoANT = varEquipo
          lngLANT = lngL

          'PAINT fila para indicar comienzo de equipo

          'CALL pinta fila
          Call pinta_Fila(lngL, varEquipo)

        End If

        'CHECK si debe ser marcado como pozo perteneciente el Rig Oficial
        If varDatoMarca = 1 Or lngL = lngLANT Then
          Me.spdCab.Row = lngL
          Me.spdCab.Col = 1
          Me.spdCab.BackColor = RGB(196, 215, 155)
        End If

      Next
'
    End If
    
    
    'CHECK si es la primera vez
    If lngPrimeraVez Then
    
    'GET Pozos Nuevos
    '
    strT = "SELECT * From EIApozosNuevos_vw order by 1"
    Set rs = SQLexec(strT)
  
    'chequeo error
    If Not SQLparam.CnErrNumero = -1 Then
      SQLError
      SQLclose
      End
    End If
  
    strT = ""
  
    'WHILE rs
    Do While Not rs.EOF
  
      strT = strT & rs(0).Value & vbCr & vbLf
  
      'MOVE puntero proximo
      rs.MoveNext
  
    Loop
      
    'CHECK si no encontro pozos nuevos
    If strT = "" Then
      strT = "No se encontraron Pozos nuevos."
    End If
      
    'SHOW pozos nuevos
    intI = MsgBox(strT, vbApplicationModal + vbInformation, "Atención - Pozos Nuevos")
  
  
    'SET foco en grilla
    Me.spdCab.SetFocus
    
    
    'CHECK si hay filas
    If Me.spdCab.ActiveRow > 0 Then
      
      'SET celda activa - la primera que contanga informacion
      Me.spdCab.SetActiveCell 1, 1
      
      'ASSIGN columna activa
      lngColumna = Me.spdCab.ActiveCol
      
      'SHOW en barra de estado
      frmMenuPrincipal.info.Panels(3) = " fila: " & Me.spdCab.ActiveRow & " de " & Me.spdCab.MaxRows & " , columna: " & Me.spdCab.ActiveCol & " de " & Me.spdCab.MaxCols & " , valor: " & Me.spdCab.Text & " "
      frmMenuPrincipal.info.Panels(3).ToolTipText = " fila: " & Me.spdCab.ActiveRow & " de " & Me.spdCab.MaxRows & " , columna: " & Me.spdCab.ActiveCol & " de " & Me.spdCab.MaxCols & " , valor: " & Me.spdCab.Text & " "
      
      'CALL ficha de Pozo
      fichaPozo (1)
      
    End If
  
    'FLAG que ya paso la primera vez
    lngPrimeraVez = False
    
    'SET Well ID - (202) en forma predeterminado
    Me.cmb = "Well ID - (202)"
    
  End If
  
  
End Sub

Private Sub Form_Initialize()
  
  lngPrimeraVez = True
  
End Sub

Private Sub Nueva_fila_detalle_Click()
  
  'SET operacion
  frmEditorComments.dsiOpera = "C"
  
  'CALL formulario edición
  frmEditorComments.Show vbModal
  
End Sub


Private Sub spdCab_Change(ByVal Col As Long, ByVal Row As Long)
  
  Dim intRes As Integer
  Dim lngFila As Long
  Dim varDato, varColumna, varIDpozo As Variant
  Dim strUpdate As String
  Dim blnB As Boolean
  
  
  
  'GET nombre de columna
  Me.spdCab.GetText Col, 0, varDato
  
  
  'SHOW pregunta confirmación
  intRes = MsgBox("Acaba de modificar la columna: " & UCase(varDato) & ". Desea reemplazar el valor en el resto de las filas seleccionadas ?", vbQuestion + vbYesNo, "Atención...")
  
  'CHECK si hizo clic en Si
  If intRes = 6 Then
    
    'GET nombre de columna
    Me.spdCab.GetText Col, 0, varColumna
    
    'GET valor modificado
    Me.spdCab.GetText Col, Row, varDato
    
    'WHILE filas
    For lngFila = lngFilaDesde To lngFilaHasta
      
      'PUT dato modificado
      Me.spdCab.SetText Col, lngFila, varDato
      
    Next
    
    'SHOW pregunta confirmación
    intRes = MsgBox("Desea guardar los cambios que acaba de realizar ?", vbQuestion + vbYesNo, "Atención...")
    
    'CHECK si hizo clic en Si
    If intRes = 6 Then
      
      
    'WHILE filas
    For lngFila = lngFilaDesde To lngFilaHasta
      
      
      'PUT dato modificado
      Me.spdCab.SetText Col, lngFila, varDato
      
      'GET ID de pozo
      Me.spdCab.GetText spdCab.GetColFromID("IDpozo"), lngFila, varIDpozo
      
      
      'EXEC store procedure - modifica columna
      SQLexec ("exec dbo.EIApozosDatos_UPD_sp '" & varColumna & "','" & IIf(IsDate(varDato), dateToIso(varDato), CStr(varDato)) & "'," & str(varIDpozo))
    
    
      'CHECK error
      If Not SQLparam.CnErrNumero = -1 Then
        SQLError
        SQLclose
        End
      End If
      
      
    Next
      
      
    End If
    
  End If
  
End Sub



Private Sub spdCab_DblClick(ByVal Col As Long, ByVal Row As Long)
  
  Dim varDato As Variant
  Dim varColValue As Variant
  Dim varColWellID As Variant
  Dim varDischargeID As Variant
    
  'GET nombre de columna
  Me.spdCab.GetText 3, Row, varColValue
  
  'GET nombre de Pozo
  Me.spdCab.GetText 1, Row, varColWellID
  
  'GET ID
  Me.spdCab.GetText spdCab.GetColFromID("IDpozo"), Row, varDischargeID
  
  'CHECK si selecciono columna 1 y es DISCHARGES, debe tener nombre de Pozo, si no, no lo dejo ir a la proxima pantalla.
  
  If Col = 1 And varColValue = "DISCHARGES" And varColWellID = "" Then
        
        Dim Val As String
        Val = InputBox("Para la Ubicaciòn: DISCHARGES, primero debe agregar el Nombre de Pozo, para poder hacer un seguimiento del mismo.", "Informaciòn...")
        
        If Val = "" Then
            Exit Sub
        End If
                
        'EXEC store procedure - modifica columna
        SQLexec ("exec dbo.EIApozosDatos_UPD_sp 'WellID', '" & Val & "'," & str(varDischargeID))
    
        'CHECK error
        If Not SQLparam.CnErrNumero = -1 Then
          SQLError
          SQLclose
          End
        End If
        
        'SET WELL ID
        Me.spdCab.SetText spdCab.GetColFromID("Well ID"), Row, Val
        
                
  End If
  
  
  If Col = 1 And Me.spdCab.ActiveRow > 0 Then
    
'    'GET nombre de columna
'    Me.spdCab.GetText Col, Row, varDato
    
    'ASSIGN fila y columna activa
    lngFila = Me.spdCab.ActiveRow
    lngColumna = Me.spdCab.ActiveCol
    
    'SET puntero interno en fila columna
    Me.spdCab.Row = Row
    Me.spdCab.Col = Col
    
    fichaPozo (Row)
    
    'ACTIVATE solapas
    Me.SSTab.Tab = 1
    
  End If
 
End Sub

Private Sub spdCab_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    
    'CHECK si el foco esta en la grilla
    If Me.ActiveControl.Name = "spdCab" Then
      
      'CHECK si grilla contiene filas
      If Me.spdCab.ActiveRow > 0 Then
        
        'ASSIGN fila y columna activa
        lngFila = Me.spdCab.ActiveRow
        lngColumna = Me.spdCab.ActiveCol
        
        'SET puntero interno en fila columna
        Me.spdCab.Row = NewRow
        Me.spdCab.Col = NewCol
        
        'PUT datos de fila a ficha de pozo
        'Si cambia el control activo, fuera de la grilla, NewRow es -1, si no controlo esto, da error.
        If NewRow <> -1 Then
          fichaPozo (NewRow)
        Else
          fichaPozo (Row)
        End If
        
        'SHOW en barra de estado
        frmMenuPrincipal.info.Panels(3) = " fila: " & NewRow & " de " & Me.spdCab.MaxRows & " , columna: " & NewCol & " de " & Me.spdCab.MaxCols & " , valor: " & Me.spdCab.Text & " "
        frmMenuPrincipal.info.Panels(3).ToolTipText = " fila: " & NewRow & " de " & Me.spdCab.MaxRows & " , columna: " & NewCol & " de " & Me.spdCab.MaxCols & " , valor: " & Me.spdCab.Text & " "
        
      End If
      
    End If
    
    
End Sub



Private Sub Form_Load()
  
  Dim rs As ADODB.Recordset
  Dim strT As String
  Dim intI As Integer
  Dim blnB As Boolean
  
  
  'CHANGE apariencia grilla
  blnB = Me.dsiCambiaApariencia(spdCab)
  blnB = Me.dsiCambiaAparienciaDet(spdDet1)
  blnB = Me.dsiCambiaAparienciaDet(spdDet2)
  
  'GET Ubicaciones
  '
  strT = "SELECT DISTINCT Ubicacion From EIApozosDatos_vw Where Not Ubicacion Is Null order by 1"
  Set rs = SQLexec(strT)
  
  'chequeo error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    End
  End If
  
  
  'WHILE ubicaciones
  intI = 0
  While Not rs.EOF
    
    'ADD 1
    intI = intI + 1
    
    'ADD ubicaciones a Lista Desplegable
    Me.cmbUbicacion.AddItem rs(0)
    
    'ADD ubicaciones a grilla para filtro
    Me.spdFiltro.SetText intI, 0, rs(0)
    
    'SET tipo de dato en celda tipo BIT para que se pueda hacer clic y filtrar
    Me.spdFiltro.Col = intI
    Me.spdFiltro.Row = 1
    Me.spdFiltro.CellType = CellTypeCheckBox
    Me.spdFiltro.TypeCheckCenter = True
    
    'SET ancho de columna
    Me.spdFiltro.ColWidth(intI) = 16
    
    'NEXT
    rs.MoveNext
    
  Wend
  
  'SET limite columnas
  Me.spdFiltro.MaxCols = intI
  
  
  'SET seleccion
  If cmbUbicacion.ListCount > 0 Then
    cmbUbicacion.ListIndex = 0
  End If
  
  'cierro
  SQLclose
  
  'ACTIVATE solapa 0
  Me.SSTab.Tab = 0
  
  'GET datos sin Where el True significa llenar la lista de columnas por primera vez
  Call GET_datos("", True)
  
  'GET datos para tree View
  Call CargarArbolColumnas
  
  'SET ancho 0 y visible = false
  Me.FrameTree.Width = 0
  Me.Tree.Visible = False
  
  
End Sub


Private Sub CargarArbolColumnas()
  
  'Carga el arbol con los nombres de las columnas de la lista metidas en grupos
  Const imagenGrupo As Long = 1
  Const imagenItem  As Long = 2
  Const imagenLogo As Long = 3
  
  Dim i As Long
  
  Set Tree.ImageList = iml
  
  Tree.Nodes.Add , , "ROOT", "Gestión Pozos", imagenLogo
  
  Tree.Nodes.Add "ROOT", tvwChild, "G-1", "Datos Generales", imagenGrupo
  
    Tree.Nodes.Add "G-1", tvwChild, "C-1", "Well ID", imagenItem
    Tree.Nodes.Add "G-1", tvwChild, "C-2", "Number", imagenItem
    Tree.Nodes.Add "G-1", tvwChild, "C-3", "Ubicacion", imagenItem
    Tree.Nodes.Add "G-1", tvwChild, "C-4", "Equipo", imagenItem
    Tree.Nodes.Add "G-1", tvwChild, "C-5", "Area", imagenItem
    Tree.Nodes.Add "G-1", tvwChild, "C-7", "Pozo Tipo", imagenItem
    Tree.Nodes.Add "G-1", tvwChild, "C-8", "Prospect", imagenItem
    Tree.Nodes.Add "G-1", tvwChild, "C-9", "Primer Monografía", imagenItem
    Tree.Nodes.Add "G-1", tvwChild, "C-10", "Monografías", imagenItem
    Tree.Nodes.Add "G-1", tvwChild, "C-11", "Ultima Monografía", imagenItem
    Tree.Nodes.Add "G-1", tvwChild, "C-12", "Well Informado", imagenItem
    Tree.Nodes.Add "G-1", tvwChild, "C-14", "Informado Por", imagenItem
    Tree.Nodes.Add "G-1", tvwChild, "C-15", "Definitiva", imagenItem
    Tree.Nodes.Add "G-1", tvwChild, "C-16", "Xsurf WGS84", imagenItem
    Tree.Nodes.Add "G-1", tvwChild, "C-17", "Ysurf WGS84", imagenItem
  
  Tree.Nodes.Add "ROOT", tvwChild, "G-2", "Plan de 8 y 13", imagenGrupo
  
    Tree.Nodes.Add "G-2", tvwChild, "C-18", "Monografía", imagenItem
    Tree.Nodes.Add "G-2", tvwChild, "C-19", "Acuifero", imagenItem
    Tree.Nodes.Add "G-2", tvwChild, "C-20", "Plan de 8", imagenItem
    Tree.Nodes.Add "G-2", tvwChild, "C-21", "Plan de 13", imagenItem
    Tree.Nodes.Add "G-2", tvwChild, "C-22", "Solicitud", imagenItem
    Tree.Nodes.Add "G-2", tvwChild, "C-23", "Prioridad", imagenItem
    Tree.Nodes.Add "G-2", tvwChild, "C-24", "Preparar Doc.", imagenItem
    Tree.Nodes.Add "G-2", tvwChild, "C-25", "Field Manifold", imagenItem
    Tree.Nodes.Add "G-2", tvwChild, "C-26", "Battery Assigned", imagenItem
    Tree.Nodes.Add "G-2", tvwChild, "C-27", "Dictamen Técnico", imagenItem
    Tree.Nodes.Add "G-2", tvwChild, "C-28", "Producción Inicial", imagenItem
  
  Tree.Nodes.Add "ROOT", tvwChild, "G-3", "Rigs Scheduler", imagenGrupo
  
    Tree.Nodes.Add "G-3", tvwChild, "C-29", "TD", imagenItem
    Tree.Nodes.Add "G-3", tvwChild, "C-30", "Total Days", imagenItem
    Tree.Nodes.Add "G-3", tvwChild, "C-31", "Remaining Days", imagenItem
    Tree.Nodes.Add "G-3", tvwChild, "C-32", "Status", imagenItem
    Tree.Nodes.Add "G-3", tvwChild, "C-33", "Start", imagenItem
    Tree.Nodes.Add "G-3", tvwChild, "C-34", "End", imagenItem
    Tree.Nodes.Add "G-3", tvwChild, "C-35", "Land Owner", imagenItem
    Tree.Nodes.Add "G-3", tvwChild, "C-36", "Land Owner Permit", imagenItem
    Tree.Nodes.Add "G-3", tvwChild, "C-37", "Rig Order Checked", imagenItem
    Tree.Nodes.Add "G-3", tvwChild, "C-38", "Rig Order", imagenItem
  
  Tree.Nodes.Add "ROOT", tvwChild, "G-4", "Aprobacion", imagenGrupo
  
    Tree.Nodes.Add "G-4", tvwChild, "C-39", "Consultora", imagenItem
    Tree.Nodes.Add "G-4", tvwChild, "C-40", "Aprobación Tipo", imagenItem
    Tree.Nodes.Add "G-4", tvwChild, "C-41", "Pedido EIA", imagenItem
    Tree.Nodes.Add "G-4", tvwChild, "C-42", "Esperada EIA", imagenItem
    Tree.Nodes.Add "G-4", tvwChild, "C-43", "Esperada EIA", imagenItem
    Tree.Nodes.Add "G-4", tvwChild, "C-44", "Recomendación", imagenItem
    Tree.Nodes.Add "G-4", tvwChild, "C-441", "Site Visit - Numero", imagenItem
    Tree.Nodes.Add "G-4", tvwChild, "C-442", "Site Visit - Fecha", imagenItem
    Tree.Nodes.Add "G-4", tvwChild, "C-443", "Site Visit - Acta", imagenItem
    Tree.Nodes.Add "G-4", tvwChild, "C-444", "Site Visit - Autor", imagenItem
    Tree.Nodes.Add "G-4", tvwChild, "C-445", "Site Visit - Comentario", imagenItem
    
    Tree.Nodes.Add "G-4", tvwChild, "C-45", "DMA Final Permit", imagenItem
    Tree.Nodes.Add "G-4", tvwChild, "C-46", "Estado", imagenItem
  
  Tree.Nodes.Add "ROOT", tvwChild, "G-5", "Estado", imagenGrupo
  
    Tree.Nodes.Add "G-5", tvwChild, "C-47", "ID Manifiesto", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-48", "Manifiesto", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-49", "Entrega Consultora", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-50", "EIA Presentado", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-51", "Envio a CS", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-52", "Presentación DMA", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-53", "Presentación SMA", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-54", "Pago Tasa Admin.", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-55", "Infor. Adicional", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-56", "Informe Técnico", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-57", "Tasa Admin. ($)", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-58", "Tasa Contralor ($)", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-59", "Estudio ($)", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-60", "Adenda ($)", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-61", "Inicio DIA", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-62", "Fin DIA", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-63", "Avance de Obra 50%", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-64", "Avance de Obra 100%", imagenItem
    Tree.Nodes.Add "G-5", tvwChild, "C-65", "Evaluación Arqueologica", imagenItem
    
  Tree.Nodes.Add "ROOT", tvwChild, "G-6", "Calculos", imagenGrupo
  
    Tree.Nodes.Add "G-6", tvwChild, "C-66", "Pedido y Recep. EIA", imagenItem
    Tree.Nodes.Add "G-6", tvwChild, "C-67", "Recep. EIA y Pres. DMA", imagenItem
    Tree.Nodes.Add "G-6", tvwChild, "C-68", "Visita y Aprob. DMA", imagenItem
    Tree.Nodes.Add "G-6", tvwChild, "C-69", "Monog. y Pedido EIA", imagenItem
    Tree.Nodes.Add "G-6", tvwChild, "C-70", "Monog. y Recep. EIA", imagenItem
    Tree.Nodes.Add "G-6", tvwChild, "C-71", "Pres. DMA y Visita", imagenItem
    Tree.Nodes.Add "G-6", tvwChild, "C-72", "Pres. Pozo y DMA", imagenItem
    Tree.Nodes.Add "G-6", tvwChild, "C-73", "Entrega EIA y Aprob.", imagenItem
    Tree.Nodes.Add "G-6", tvwChild, "C-74", "IDpozo", imagenItem
    Tree.Nodes.Add "G-6", tvwChild, "C-75", "EIAIDpozo", imagenItem
    
  'SET nodos chequeados y expandidos
  For i = 1 To Tree.Nodes.Count
    Tree.Nodes(i).Checked = True
    Tree.Nodes(i).Expanded = True
  Next i
  
End Sub



Private Sub cmdGuardar_Click()

  Dim lngL As Long
  Dim varDato As Variant
  Dim strT As String
  Dim varNumero, varFecha, varComen, varAutor, varActa As Variant
  Dim intRes As Boolean
    
  'SHOW pregunta confirmación
  intRes = MsgBox("Desea guardar los cambios realizados en el pozo: " & Me.txtWellID & " ?", vbQuestion + vbYesNo, "Atención...")
  
  
  'CHECK si hizo clic en Si
  If intRes Then
  
  '  'BD.BeginTrans
  '  If VerificarCampos Then
      
      'BUILD parametros cabecera para enviar a Store Procedure
      'los separo en 3 porque VB6 no soporta la concatenacion de muchas filas juntas
      '
      strT = Me.txtIDpozo & "," & _
      "'" & dateToIso(Me.dtpPrimerMonografia.Value) & "'," & _
      "'" & IIf(Me.txtMonografias = "", 0, Me.txtMonografias) & "'," & _
      "'" & dateToIso(Me.dtpUltimaMonografia.Value) & "'," & _
      str(Me.chkDefinitiva.Value) & "," & _
      str(Me.chkMonografia.Value) & "," & _
      str(Me.chkAcuifero.Value) & "," & _
      str(Me.chkPrognosis.Value) & "," & _
      str(Me.chkPrograma.Value) & "," & _
      "'" & dateToIso(Me.dtpFechaSolicitud.Value) & "'," & _
      "'" & dateToIso(Me.dtpFechaPrioridad.Value) & "'," & _
      "'" & Me.txtDocumentoAPreparar & "'," & _
      "'" & Me.txtFieldManifold & "'," & _
      "'" & Me.txtBatteryAssigned & "'," & _
      "'" & dateToIso(Me.dtpFechaEntregaDictamenTecnico.Value) & "'," & _
      "'" & dateToIso(Me.dtpLandOwnerPermitDate.Value) & "'," & _
      "'" & Me.txtConsult & "'," & _
      "'" & Me.txtType & "'," & _
      "'" & dateToIso(Me.dtpFechaPedidoETIA.Value) & "'," & _
      "'" & dateToIso(Me.dtpFechaEsperadaETIA.Value) & "'," & _
      "'" & Me.txtConsultantRecomendation & "'," & _
      "'" & dateToIso(Me.dtpDMAFinalPermit.Value) & "'," & _
      "'" & Me.txtEstado & "'," & _
      "'" & Me.txtIDManifiesto & "',"
      
      strT = strT & "'" & dateToIso(Me.dtpFechaManifiesto.Value) & "'," & _
      "'" & dateToIso(Me.dtpFechaEntregaEiaXConsultoraAOxy.Value) & "'," & _
      Me.chkEIAPresentado.Value & "," & _
      "'" & dateToIso(Me.dtpFechaEnvioACS.Value) & "'," & _
      "'" & dateToIso(Me.dtpFechaPresentacionDMA.Value) & "'," & _
      "'" & dateToIso(Me.dtpFechaPresentacionSMA.Value) & "'," & _
      "'" & dateToIso(Me.dtpPagoTasaAdministrativa.Value) & "'," & _
      "'" & dateToIso(Me.dtpFechaInfoComplementaria.Value) & "'," & _
      "'" & Me.txtTechnicalReport & "'," & _
      IIf(Me.txtTasaAdministrativa = "", 0, Me.txtTasaAdministrativa) & "," & _
      IIf(Me.txtTasaContralor = "", 0, Me.txtTasaContralor) & "," & _
      IIf(Me.txtEstudio = "", 0, Me.txtEstudio) & "," & _
      IIf(Me.txtAdenda = "", 0, Me.txtAdenda) & "," & _
      "'" & dateToIso(Me.dtpFECHAINICIODIA.Value) & "'," & _
      "'" & dateToIso(Me.dtpFECHAFINDIA.Value) & "'," & _
      "'" & dateToIso(Me.dtpInformeAvanceObra50PorCiento.Value) & "'," & _
      "'" & dateToIso(Me.dtpInformeAvanceObra100PorCiento.Value) & "'," & _
      "'" & Me.txtInformeEvaluacionArqueologico & "'," & _
      IIf(Me.txtTiempoEntrePedidoETIAyRecepcionETIA = "", 0, Me.txtTiempoEntrePedidoETIAyRecepcionETIA) & "," & _
      IIf(Me.txtTiempoEntreRecepcionETIAYpresentacionAnteDMA = "", 0, Me.txtTiempoEntreRecepcionETIAYpresentacionAnteDMA) & "," & _
      IIf(Me.txtTiempoEntreVisitaYAprobacionFinaldeDMA = "", 0, Me.txtTiempoEntreVisitaYAprobacionFinaldeDMA) & "," & _
      IIf(Me.txtTiempoEntrePrimerMonografiaYPedidoEIA = "", 0, Me.txtTiempoEntrePrimerMonografiaYPedidoEIA) & "," & _
      IIf(Me.txtTiempoEntrePrimerMonografiayRecepcionETIA = "", 0, Me.txtTiempoEntrePrimerMonografiayRecepcionETIA) & "," & _
      IIf(Me.txtTiempoEntrePresentacionAnteDMAyVisita = "", 0, Me.txtTiempoEntrePresentacionAnteDMAyVisita) & "," & _
      IIf(Me.txtTiempoEntrePresentacionDePozoYAprobacionFinalDMA = "", 0, Me.txtTiempoEntrePresentacionDePozoYAprobacionFinalDMA) & ","
      
      strT = strT & IIf(Me.txtTiempoEntreEntregaEIAaDMAyAprobacion = "", 0, Me.txtTiempoEntreEntregaEIAaDMAyAprobacion)
      
      
      'SAVE cabecera
      SQLexec ("exec dbo.EIApozosDatos_INS_sp " & strT)
    
      'CHECK error
      If Not SQLparam.CnErrNumero = -1 Then
        SQLError
        SQLclose
        End
      End If
        
        
      'DELETE detalle completo
      '       Lo hago de este modo porque es mas fácil y rápido eliminar y luego insertar todo
      SQLexec ("exec dbo.EIApozosDatosDetalle_ELI_sp " & Me.txtIDpozo)
    
      'CHECK error
      If Not SQLparam.CnErrNumero = -1 Then
        SQLError
        SQLclose
        End
      End If
      
      
      'CHECK si detalle 1, contiene info
      If Me.spdDet1.DataRowCnt > 0 Then
        
        For lngL = 1 To Me.spdDet1.DataRowCnt
          
          'GET datos de grippa
          Me.spdDet1.GetText 2, lngL, varNumero
          Me.spdDet1.GetText 3, lngL, varFecha
          Me.spdDet1.GetText 4, lngL, varComen
          Me.spdDet1.GetText 5, lngL, varAutor
          
          'BUILD parametros detalle 1 para enviar a Store Procedure
          strT = Me.txtIDpozo & "," & _
          "'" & "LANDOWNER" & "'," & _
          varNumero & "," & _
          "'" & dateToIso(varFecha) & "'," & _
          "'" & varComen & "'," & _
          "'" & varAutor & "'," & _
          "null"
          
          'CALL store procedure con parametros
          SQLexec ("exec dbo.EIApozosDatosDetalle_INS_sp " & strT)
        
          'CHECK error
          If Not SQLparam.CnErrNumero = -1 Then
            SQLError
            SQLclose
            End
          End If
          
        Next
        
      End If
      
      
      'CHECK si detalle 2, contiene info
      If Me.spdDet2.DataRowCnt > 0 Then
        
        For lngL = 1 To Me.spdDet2.DataRowCnt
          
          'GET datos de grippa
          Me.spdDet2.GetText 2, lngL, varNumero
          Me.spdDet2.GetText 3, lngL, varFecha
          Me.spdDet2.GetText 4, lngL, varComen
          Me.spdDet2.GetText 5, lngL, varAutor
          Me.spdDet2.GetText 6, lngL, varActa
          
          'BUILD parametros detalle 1 para enviar a Store Procedure
          strT = Me.txtIDpozo & "," & _
          "'" & "SITEVISIT" & "'," & _
          varNumero & "," & _
          "'" & dateToIso(varFecha) & "'," & _
          "'" & varComen & "'," & _
          "'" & varAutor & "'," & _
          "'" & IIf(varActa = "", "null", varActa) & "'"
          
          'CALL store procedure con parametros
          SQLexec ("exec dbo.EIApozosDatosDetalle_INS_sp " & strT)
          
          'CHECK error
          If Not SQLparam.CnErrNumero = -1 Then
            SQLError
            SQLclose
            End
          End If
          
        Next
        
      End If
      
      
      'UPDATE grilla cabecera con los datos que se acaban de modificar y guardar
      
      'SET fila actual
      Me.spdCab.SetActiveCell lngColumna, lngFila
      
      Me.spdCab.SetText spdCab.GetColFromID("Ubicacion"), Me.spdCab.ActiveRow, Me.cmbUbicacion
      
      Me.spdCab.SetText spdCab.GetColFromID("Primer Monografía"), Me.spdCab.ActiveRow, Me.dtpPrimerMonografia.Value
      Me.spdCab.SetText spdCab.GetColFromID("Monografías"), Me.spdCab.ActiveRow, Me.txtMonografias
      Me.spdCab.SetText spdCab.GetColFromID("Ultima Monografía"), Me.spdCab.ActiveRow, Me.dtpUltimaMonografia.Value
      
      Me.spdCab.SetText spdCab.GetColFromID("Definitiva"), Me.spdCab.ActiveRow, Me.chkDefinitiva
      
      Me.spdCab.SetText spdCab.GetColFromID("Monografía"), Me.spdCab.ActiveRow, Me.chkMonografia
      Me.spdCab.SetText spdCab.GetColFromID("Acuifero"), Me.spdCab.ActiveRow, Me.chkAcuifero
      Me.spdCab.SetText spdCab.GetColFromID("Plan de 8"), Me.spdCab.ActiveRow, Me.chkPrognosis
      Me.spdCab.SetText spdCab.GetColFromID("Plan de 13"), Me.spdCab.ActiveRow, Me.chkPrograma
  
      Me.spdCab.SetText spdCab.GetColFromID("Solicitud"), Me.spdCab.ActiveRow, Me.dtpFechaSolicitud.Value
      Me.spdCab.SetText spdCab.GetColFromID("Prioridad"), Me.spdCab.ActiveRow, Me.dtpFechaPrioridad.Value
      Me.spdCab.SetText spdCab.GetColFromID("Preparar Doc."), Me.spdCab.ActiveRow, Me.txtDocumentoAPreparar
      Me.spdCab.SetText spdCab.GetColFromID("Field Manifold"), Me.spdCab.ActiveRow, Me.txtFieldManifold
      Me.spdCab.SetText spdCab.GetColFromID("Battery Assigned"), Me.spdCab.ActiveRow, Me.txtBatteryAssigned
      Me.spdCab.SetText spdCab.GetColFromID("Dictamen Técnico"), Me.spdCab.ActiveRow, Me.dtpFechaEntregaDictamenTecnico.Value
  
      Me.spdCab.SetText spdCab.GetColFromID("Land Owner Permit"), Me.spdCab.ActiveRow, Me.dtpLandOwnerPermitDate.Value
      Me.spdCab.SetText spdCab.GetColFromID("Consultora"), Me.spdCab.ActiveRow, Me.txtConsult
      Me.spdCab.SetText spdCab.GetColFromID("Aprobación Tipo"), Me.spdCab.ActiveRow, Me.txtType
      Me.spdCab.SetText spdCab.GetColFromID("Pedido EIA"), Me.spdCab.ActiveRow, Me.dtpFechaPedidoETIA.Value
      Me.spdCab.SetText spdCab.GetColFromID("Esperada EIA"), Me.spdCab.ActiveRow, Me.dtpFechaEsperadaETIA.Value
  
      Me.spdCab.SetText spdCab.GetColFromID("Recomendación"), Me.spdCab.ActiveRow, Me.txtConsultantRecomendation
  
      Me.spdCab.SetText spdCab.GetColFromID("DMA Final Permit"), Me.spdCab.ActiveRow, Me.dtpDMAFinalPermit.Value
      Me.spdCab.SetText spdCab.GetColFromID("Estado"), Me.spdCab.ActiveRow, Me.txtEstado
  
      Me.spdCab.SetText spdCab.GetColFromID("ID Manifiesto"), Me.spdCab.ActiveRow, Me.txtIDManifiesto
      Me.spdCab.SetText spdCab.GetColFromID("Manifiesto"), Me.spdCab.ActiveRow, Me.dtpFechaManifiesto.Value
      Me.spdCab.SetText spdCab.GetColFromID("Entrega Consultora"), Me.spdCab.ActiveRow, Me.dtpFechaEntregaEiaXConsultoraAOxy.Value
      Me.spdCab.SetText spdCab.GetColFromID("EIA Presentado"), Me.spdCab.ActiveRow, Me.chkEIAPresentado.Value
      Me.spdCab.SetText spdCab.GetColFromID("Envio a CS"), Me.spdCab.ActiveRow, Me.dtpFechaEnvioACS.Value
      Me.spdCab.SetText spdCab.GetColFromID("Presentación DMA"), Me.spdCab.ActiveRow, Me.dtpFechaPresentacionDMA.Value
      Me.spdCab.SetText spdCab.GetColFromID("Presentación SMA"), Me.spdCab.ActiveRow, Me.dtpFechaPresentacionSMA.Value
      Me.spdCab.SetText spdCab.GetColFromID("Pago Tasa Admin."), Me.spdCab.ActiveRow, Me.dtpPagoTasaAdministrativa.Value
      Me.spdCab.SetText spdCab.GetColFromID("Infor. Adicional"), Me.spdCab.ActiveRow, Me.dtpFechaInfoComplementaria.Value
      Me.spdCab.SetText spdCab.GetColFromID("Informe Técnico"), Me.spdCab.ActiveRow, Me.txtTechnicalReport
      Me.spdCab.SetText spdCab.GetColFromID("Tasa Admin. ($)"), Me.spdCab.ActiveRow, Me.txtTasaAdministrativa
      Me.spdCab.SetText spdCab.GetColFromID("Tasa Contralor ($)"), Me.spdCab.ActiveRow, Me.txtTasaContralor
      Me.spdCab.SetText spdCab.GetColFromID("Estudio ($)"), Me.spdCab.ActiveRow, Me.txtEstudio
      Me.spdCab.SetText spdCab.GetColFromID("Adenda ($)"), Me.spdCab.ActiveRow, Me.txtAdenda
      Me.spdCab.SetText spdCab.GetColFromID("Inicio DIA"), Me.spdCab.ActiveRow, Me.dtpFECHAINICIODIA.Value
      Me.spdCab.SetText spdCab.GetColFromID("Fin DIA"), Me.spdCab.ActiveRow, Me.dtpFECHAFINDIA.Value
      Me.spdCab.SetText spdCab.GetColFromID("Avance de Obra 50%"), Me.spdCab.ActiveRow, Me.dtpInformeAvanceObra50PorCiento.Value
      Me.spdCab.SetText spdCab.GetColFromID("Avance de Obra 100%"), Me.spdCab.ActiveRow, Me.dtpInformeAvanceObra100PorCiento.Value
      Me.spdCab.SetText spdCab.GetColFromID("Evaluación Arqueologica"), Me.spdCab.ActiveRow, Me.txtInformeEvaluacionArqueologico
  
      Me.spdCab.SetText spdCab.GetColFromID("Pedido y Recep. EIA"), Me.spdCab.ActiveRow, Me.txtTiempoEntrePedidoETIAyRecepcionETIA
      Me.spdCab.SetText spdCab.GetColFromID("Recep. EIA y Pres. DMA"), Me.spdCab.ActiveRow, Me.txtTiempoEntreRecepcionETIAYpresentacionAnteDMA
      Me.spdCab.SetText spdCab.GetColFromID("Visita y Aprob. DMA"), Me.spdCab.ActiveRow, Me.txtTiempoEntreVisitaYAprobacionFinaldeDMA
      Me.spdCab.SetText spdCab.GetColFromID("Monog. y Pedido EIA"), Me.spdCab.ActiveRow, Me.txtTiempoEntrePrimerMonografiaYPedidoEIA
      Me.spdCab.SetText spdCab.GetColFromID("Monog. y Recep. EIA"), Me.spdCab.ActiveRow, Me.txtTiempoEntrePrimerMonografiayRecepcionETIA
      Me.spdCab.SetText spdCab.GetColFromID("Pres. DMA y Visita"), Me.spdCab.ActiveRow, Me.txtTiempoEntrePresentacionAnteDMAyVisita
      Me.spdCab.SetText spdCab.GetColFromID("Pres. Pozo y DMA"), Me.spdCab.ActiveRow, Me.txtTiempoEntrePresentacionDePozoYAprobacionFinalDMA
      Me.spdCab.SetText spdCab.GetColFromID("Entrega EIA y Aprob."), Me.spdCab.ActiveRow, Me.txtTiempoEntreEntregaEIAaDMAyAprobacion
    
    End If
    
    
End Sub


Private Function VerificarCampos() As Boolean
'Verifica los datos
  On Error GoTo ErrorHandler
  Dim Errores As String
  
  If IDModificacion = 0 And InfoGlobal.IDTipoUsuario <> enmTipoUsuarioAdministrador Then
    Errores = Errores & Chr(vbKeyReturn) & "Su nivel de usuario no le permite dar de alta nuevos pozos."
    MsgBox "Imposible guardar: " & Chr(vbKeyReturn) & Errores, vbOKOnly + vbCritical, "Error"
    VerificarCampos = False
    Exit Function
  End If
  
  
  If cmbUbicacion = "RIGS SCHED." And txtEquipo = "" Then
    Errores = Errores & Chr(vbKeyReturn) & "El campo Equipo es requerido si la ubicacion es RIGS SCHED."
  End If
  
  If txtWellID = "" Then
    Errores = Errores & Chr(vbKeyReturn) & "El campo Well ID es requerido"
  Else
    If Not IsNull(ObtenerValorCampo("POZOS", "WELLID", "IDPOZO <> " & IDModificacion & " AND WELLID LIKE '" & txtWellID & "'")) Then
      Errores = Errores & Chr(vbKeyReturn) & "Ya existe un Pozo con el mismo WellID"
    End If
  End If
  
  If txtTD <> "" Then
    If Not IsNumeric(txtTD) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo TD debe ser numerico"
    Else
      If CLng(txtTD) < 0 Then
        Errores = Errores & Chr(vbKeyReturn) & "El campo TD debe ser mayor que cero"
      End If
    End If
  End If
  
  If txtMonografias <> "" Then
    If Not IsNumeric(txtMonografias) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo Monografias debe ser numerico"
    Else
      If CLng(txtMonografias) < 0 Then
        Errores = Errores & Chr(vbKeyReturn) & "El campo Monografias debe ser mayor que cero"
      End If
    End If
  End If
  
  If txtTotDays <> "" Then
    If Not IsNumeric(txtTotDays) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo Tot Days debe ser numerico"
    Else
      If CLng(txtTotDays) < 0 Then
        Errores = Errores & Chr(vbKeyReturn) & "El campo Tot Days debe ser mayor que cero"
      End If
    End If
  End If
  
  If txtRemDays <> "" Then
    If Not IsNumeric(txtRemDays) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo Rem Days debe ser numerico"
    Else
      If CLng(txtRemDays) < 0 Then
        Errores = Errores & Chr(vbKeyReturn) & "El campo Rem Days debe ser mayor que cero"
      End If
    End If
  End If
  
  If txtTiempoEntrePedidoETIAyRecepcionETIA <> "" Then
    If Not IsNumeric(txtTiempoEntrePedidoETIAyRecepcionETIA) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo EIA Duration debe ser numerico"
    End If
  End If
  
  If txtX_PDC <> "" Then
    If Not IsNumeric(txtX_PDC) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo X PDC debe ser numerico"
    Else
      If CDbl(txtX_PDC) < 0 Then
        Errores = Errores & Chr(vbKeyReturn) & "El campo X PDC debe ser mayor que cero"
      End If
    End If
  End If
  
  If txtY_PDC <> "" Then
    If Not IsNumeric(txtY_PDC) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo Y PDC debe ser numerico"
    Else
      If CDbl(txtY_PDC) < 0 Then
        Errores = Errores & Chr(vbKeyReturn) & "El campo Y PDC debe ser mayor que cero"
      End If
    End If
  End If
  
    
  If txtX_Pos94 <> "" Then
    If Not IsNumeric(txtX_Pos94) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo X POS94 debe ser numerico"
    Else
      If CDbl(txtX_Pos94) < 0 Then
        Errores = Errores & Chr(vbKeyReturn) & "El campo X POS94 debe ser mayor que cero"
      End If
    End If
  End If
    
  If txtY_Pos94 <> "" Then
    If Not IsNumeric(txtY_Pos94) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo Y POS94 debe ser numerico"
    Else
      If CDbl(txtY_Pos94) < 0 Then
        Errores = Errores & Chr(vbKeyReturn) & "El campo Y POS94 debe ser mayor que cero"
      End If
    End If
  End If
  
'  If txtAFE <> "" Then
'    If Not IsNumeric(txtAFE) Then
'      Errores = Errores & Chr(vbKeyReturn) & "El campo AFE debe ser numerico"
'    Else
'      If CDbl(txtAFE) < 0 Then
'        Errores = Errores & Chr(vbKeyReturn) & "El campo AFE debe ser mayor que cero"
'      End If
'    End If
'  End If
'  MsgBox "1"
  
  If txtTasaAdministrativa <> "" Then
    If Not IsNumeric(txtTasaAdministrativa) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo Tasa Administrativa debe ser numerico"
    Else
      If CDbl(txtTasaAdministrativa) < 0 Then
        Errores = Errores & Chr(vbKeyReturn) & "El campo Tasa Administrativa debe ser numerico"
      End If
    End If
  End If
  
  If txtEstudio <> "" Then
    If Not IsNumeric(txtEstudio) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo Estudio debe ser numerico"
    Else
      If CDbl(txtEstudio) < 0 Then
        Errores = Errores & Chr(vbKeyReturn) & "El campo Estudio debe ser numerico"
      End If
    End If
  End If
  
  If txtAdenda <> "" Then
    If Not IsNumeric(txtAdenda) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo Adenda debe ser numerico"
    Else
      If CDbl(txtAdenda) < 0 Then
        Errores = Errores & Chr(vbKeyReturn) & "El campo Adenda debe ser numerico"
      End If
    End If
  End If
  
  If txtTiempoEntrePedidoETIAyRecepcionETIA <> "" Then
    If Not IsNumeric(txtTiempoEntrePedidoETIAyRecepcionETIA) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo Tiempo Entre pedido y recepcion ETIA debe ser numerico"
    End If
  End If
  
    
  If txtTiempoEntreRecepcionETIAYpresentacionAnteDMA <> "" Then
    If Not IsNumeric(txtTiempoEntreRecepcionETIAYpresentacionAnteDMA) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo Tiempo Entre Recepcion ETIA y Presentacion DMA debe ser numerico"
    End If
  End If
  
  If txtTiempoEntrePresentacionAnteDMAyVisita <> "" Then
    If Not IsNumeric(txtTiempoEntrePresentacionAnteDMAyVisita) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo Tiempo Entre Presentacion Ante DMA y Visita debe ser numerico"
    End If
  End If
  
  If txtTiempoEntreVisitaYAprobacionFinaldeDMA <> "" Then
    If Not IsNumeric(txtTiempoEntreVisitaYAprobacionFinaldeDMA) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo Tiempo Entre Visita y Aprobacion Final de DMA debe ser numerico"
    End If
  End If
  
  If txtTiempoEntrePresentacionDePozoYAprobacionFinalDMA <> "" Then
    If Not IsNumeric(txtTiempoEntrePresentacionDePozoYAprobacionFinalDMA) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo Tiempo Entre Presentacion de pozo y Aprobacion Final de DMA debe ser numerico"
    End If
  End If
  
  If txtTiempoEntrePrimerMonografiaYPedidoEIA <> "" Then
    If Not IsNumeric(txtTiempoEntrePrimerMonografiaYPedidoEIA) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo Tiempo Entre Primer monografia y pedido EIA debe ser numerico"
    End If
  End If
 
  If txtTiempoEntreEntregaEIAaDMAyAprobacion <> "" Then
    If Not IsNumeric(txtTiempoEntreEntregaEIAaDMAyAprobacion) Then
      Errores = Errores & Chr(vbKeyReturn) & "El campo Tiempo Entre Entrega EIA a DMA y Aprobacion debe ser numerico"
    End If
  End If
  
  VerificarCampos = (Errores = "")
  If Errores <> "" Then
    MsgBox "Se han encontrado los siguientes errores al intentar guardar: " & Chr(vbKeyReturn) & Errores, vbOKOnly + vbCritical, "Error"
  End If
      
ErrorHandler:
  ErrorHandler
End Function



Private Sub spdCab_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
  
    Dim varColumna As Variant
    Dim varColValue As Variant
    Dim blnEdit As Boolean
    
  
    'GET nombre de columna
    Me.spdCab.GetText Col, 0, varColumna
    
    'CHECK que columna se esta actualizando
    Select Case varColumna
    
    Case "Primer Monografía"
      blnEdit = True
      
    Case "Ultima Monografía"
      blnEdit = True
    
    Case "Solicitud"
      blnEdit = True
    
    Case "Prioridad"
      blnEdit = True
    
    Case "Preparar Doc."
      blnEdit = True
    
    Case "Field Manifold"
      blnEdit = True
    
    Case "Battery Assigned"
      blnEdit = True
    
    Case "Dictamen Técnico"
      blnEdit = True
    
    Case "Land Owner Permit"
      blnEdit = True
    
    Case "Consultora"
      blnEdit = True
    
    Case "Aprobación Tipo"
      blnEdit = True
    
    Case "Pedido EIA"
      blnEdit = True
    
    Case "Esperada EIA"
      blnEdit = True
    
    Case "Recomendación"
      blnEdit = True
    
    Case "ID Manifiesto"
      blnEdit = True
    
    Case "Manifiesto"
      blnEdit = True
    
    Case "Entrega Consultora"
      blnEdit = True
    
    Case "Envio a CS"
      blnEdit = True
    
    Case "Presentación DMA"
      blnEdit = True
    
    Case "Presentación SMA"
      blnEdit = True
    
    Case "Pago Tasa Admin."
      blnEdit = True
    
    Case "Infor. Adicional"
      blnEdit = True
    
    Case "Informe Técnico"
      blnEdit = True
    
    Case "Tasa Admin. ($)"
      blnEdit = True
    
    Case "Tasa Contralor ($)"
      blnEdit = True
    
    Case "Estudio ($)"
      blnEdit = True
    
    Case "Adenda ($)"
      blnEdit = True
    
    Case "Inicio DIA"
      blnEdit = True
    
    Case "Fin DIA"
    
    Case "Avance de Obra 50%"
      blnEdit = True
    
    Case "Avance de Obra 100%"
      blnEdit = True
    
    Case "Evaluación Arqueologica"
      blnEdit = True
    
    Case Else
      blnEdit = False
    
    End Select
      
    'CHECK si columna permite edición
    If Not blnEdit Then
    
        Exit Sub
    
    End If
  
    'CALL menu
    Me.PopupMenu VariosPozos
  
End Sub



Private Sub spdDet1_DblClick(ByVal Col As Long, ByVal Row As Long)
  
  Dim varDato As Variant
  Dim blnB As Boolean
  
  'GET dato de celda
  Me.spdDet1.GetText Me.spdDet1.ActiveCol, Me.spdDet1.ActiveRow, varDato
  
  'CHECK si hay datos en la fila en donde se hizo doble clic
  If varDato = "" Then
    blnB = MsgBox("La fila seleccionada, no contiene información.", vbCritical + vbOKOnly, "Atención...")
    Exit Sub
  End If
  
  'SET detalle
  frmEditorComments.dsiDetalle = "LANDOWNER"
  
  'SET operacion
  frmEditorComments.dsiOpera = "R"
  
  'CALL formulario edición
  frmEditorComments.Show vbModal
  
End Sub

Private Sub spdDet1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  'CHECK si se hizo clic con boton derecho del mouse
  If Button = 2 Then
    
    'SET detalle
    'Esto lo hago porque el formulario de edicion de detalle esta
    'compartido para las 2 grillas de detalle, para saber con cual trabajar
    frmEditorComments.dsiDetalle = "LANDOWNER"
    
    'CALL menu
    Me.PopupMenu Grilla
    
  End If
  
End Sub


Private Sub spdDet2_DblClick(ByVal Col As Long, ByVal Row As Long)

  Dim varDato As Variant
  Dim blnB As Boolean
  
  'GET dato de celda
  Me.spdDet2.GetText Me.spdDet2.ActiveCol, Me.spdDet2.ActiveRow, varDato
  
  'CHECK si hay datos en la fila en donde se hizo doble clic
  If varDato = "" Then
    blnB = MsgBox("La fila seleccionada, no contiene información.", vbCritical + vbOKOnly, "Atención...")
    Exit Sub
  End If
  
  'SET detalle
  frmEditorComments.dsiDetalle = "SITEVISIT"
  
  'SET operacion
  frmEditorComments.dsiOpera = "R"
  
  'CALL formulario edición
  frmEditorComments.Show vbModal
  
End Sub

Private Sub spdDet2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'CHECK si se hizo clic con boton derecho del mouse
  If Button = 2 Then
    
    'SET detalle
    'Esto lo hago porque el formulario de edicion de detalle esta
    'compartido para las 2 grillas de detalle, para saber con cual trabajar
    frmEditorComments.dsiDetalle = "SITEVISIT"
    
    'CALL menu
    Me.PopupMenu Grilla
    
  End If

End Sub


Private Sub HabilitarControl(Key As String, Checked As Boolean)

End Sub









Private Sub CalcularValoresAutomaticos()

'Autocompleta los valores en los campos necesarios
  On Error GoTo ErrorHandler
  
  If CargandoSolapa = True Then
    Exit Sub
  End If
  
  
  
  'ChkMonografia
  If Not IsNull(dtpPrimerMonografia.Value) Then
    chkMonografia.Value = vbChecked
  Else
    chkMonografia.Value = vbUnchecked
  End If
  
  'Documento a Preparar
  txtDocumentoAPreparar = ""
  If chkMonografia.Value = vbUnchecked Then
    txtDocumentoAPreparar = txtDocumentoAPreparar & ", " & chkMonografia.Caption
  End If
  If chkAcuifero.Value = vbUnchecked Then
    txtDocumentoAPreparar = txtDocumentoAPreparar & ", " & chkAcuifero.Caption
  End If
  If chkPrognosis.Value = vbUnchecked Then
    txtDocumentoAPreparar = txtDocumentoAPreparar & ", " & chkPrognosis.Caption
  End If
  If chkPrograma.Value = vbUnchecked Then
    txtDocumentoAPreparar = txtDocumentoAPreparar & ", " & chkPrograma.Caption
  End If
  txtDocumentoAPreparar = Mid(txtDocumentoAPreparar, 3)
  
  'Estado
  If Not IsNull(dtpDMAFinalPermit.Value) And dtpDMAFinalPermit.Value <> "01/01/1900" Then
    txtEstado = "APROBADO"
    txtEstado.BackColor = &H80FF80
  Else
    txtEstado = ""
    txtEstado.BackColor = vbWhite
  End If
  
  'txtTiempoEntrePedidoETIAyRecepcionETIA
  If Not IsNull(dtpFechaEntregaEiaXConsultoraAOxy.Value) Then
    If Not IsNull(dtpFechaPedidoETIA.Value) Then
      txtTiempoEntrePedidoETIAyRecepcionETIA = DateDiff("d", dtpFechaPedidoETIA.Value, dtpFechaEntregaEiaXConsultoraAOxy.Value)
    Else
      txtTiempoEntrePedidoETIAyRecepcionETIA = ""
    End If
  Else
    txtTiempoEntrePedidoETIAyRecepcionETIA = ""
  End If
  'txtTiempoEntrePrimerMonografiayRecepcionETIA
  If Not IsNull(dtpFechaEntregaEiaXConsultoraAOxy.Value) Then
    If Not IsNull(dtpPrimerMonografia.Value) Then
      txtTiempoEntrePrimerMonografiayRecepcionETIA = DateDiff("d", dtpPrimerMonografia.Value, dtpFechaEntregaEiaXConsultoraAOxy.Value)
    Else
      txtTiempoEntrePrimerMonografiayRecepcionETIA = ""
    End If
  Else
    txtTiempoEntrePrimerMonografiayRecepcionETIA = ""
  End If
  'txtTiempoEntreRecepcionYPresentacionAnteDMA
  If Not IsNull(dtpFechaEntregaEiaXConsultoraAOxy.Value) Then
    If Not IsNull(dtpFechaPresentacionDMA.Value) Then
      txtTiempoEntreRecepcionETIAYpresentacionAnteDMA = DateDiff("d", dtpFechaEntregaEiaXConsultoraAOxy.Value, dtpFechaPresentacionDMA.Value)
    Else
      txtTiempoEntreRecepcionETIAYpresentacionAnteDMA = ""
    End If
  Else
    txtTiempoEntreRecepcionETIAYpresentacionAnteDMA = ""
  End If
  
  
  'txtTiempoEntrePresentacionDePozoYAprobacionFinalDMA
  If Not IsNull(dtpDMAFinalPermit.Value) Then
    If Not IsNull(dtpPrimerMonografia.Value) Then
      txtTiempoEntrePresentacionDePozoYAprobacionFinalDMA = DateDiff("d", dtpPrimerMonografia.Value, dtpDMAFinalPermit.Value)
    Else
      txtTiempoEntrePresentacionDePozoYAprobacionFinalDMA = ""
    End If
  Else
    txtTiempoEntrePresentacionDePozoYAprobacionFinalDMA = ""
  End If
  
  'txtTiempoEntrePrimerMonografiaYPedidoEIA
  If Not IsNull(dtpFechaPedidoETIA.Value) Then
    If Not IsNull(dtpPrimerMonografia.Value) Then
      txtTiempoEntrePrimerMonografiaYPedidoEIA = DateDiff("d", dtpPrimerMonografia.Value, dtpFechaPedidoETIA.Value)
    Else
      txtTiempoEntrePrimerMonografiaYPedidoEIA = ""
    End If
  Else
    txtTiempoEntrePrimerMonografiaYPedidoEIA = ""
  End If
  'txtTiempoEntreEntregaEIAaDMAyAprobacion
  If Not IsNull(dtpDMAFinalPermit.Value) Then
    If Not IsNull(dtpFechaPresentacionDMA.Value) Then
      txtTiempoEntreEntregaEIAaDMAyAprobacion = DateDiff("d", dtpFechaPresentacionDMA.Value, dtpDMAFinalPermit.Value)
    Else
      txtTiempoEntreEntregaEIAaDMAyAprobacion = ""
    End If
  Else
    txtTiempoEntreEntregaEIAaDMAyAprobacion = ""
  End If
  
  
ErrorHandler:
  ErrorHandler
End Sub


Private Sub Tree_NodeCheck(ByVal Node As MSComctlLib.Node)

'Marca y desmarca los nodos hijos, ademas oculta o muestra las columnas segun el caso
  Dim n As Node
  
  Me.spdCab.Redraw = False
  
  If Node.Children = 0 Then
  
    'HabilitarControl Node.Key, Node.Checked
    
    If Node.Checked Then
      
      Node.Checked = True
       
      'HABILITA columna, puntero a columna y oculta columna
      Me.spdCab.Col = Me.spdCab.GetColFromID(Node)
      Me.spdCab.ColHidden = False
      
    Else
      
      Node.Checked = False
            
      'DESHABILITA columna, puntero a columna y oculta columna
      Me.spdCab.Col = Me.spdCab.GetColFromID(Node)
      Me.spdCab.ColHidden = True
      
    End If
    
'    mfg.Refresh
  
  Else
  
    Set n = Node.Child
    
    While Not n Is Nothing
      n.Checked = Node.Checked
      Tree_NodeCheck n
      Set n = n.Next
    Wend
    
  End If
  
  Me.spdCab.Redraw = False
  
End Sub

Private Sub txtDato_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
    
    'CALL boton Buscar
    Call cmdBuscar1_Click
    
  End If
  
End Sub

Private Sub txtEstudio_Change()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub


Private Sub txtTasaAdministrativa_Change()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  txtTasaAdministrativa = Replace(txtTasaAdministrativa, ",", ".")
  txtTasaAdministrativa.SelStart = Len(txtTasaAdministrativa)
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub txtTasaContralor_Change()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  txtTasaContralor = Replace(txtTasaContralor, ",", ".")
  txtTasaContralor.SelStart = Len(txtTasaContralor)
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub txtTechnicalReport_Change()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  txtTechnicalReport = Replace(txtTechnicalReport, ",", ".")
  txtTechnicalReport.SelStart = Len(txtTechnicalReport)
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub txtRemDays_Change()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub optAvanzada_Click()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub optDesarrollo_Click()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub optExploratorio_Click()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub cmbPozo_Change()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub cmbYacimiento_Click()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub cmbPozo_Click()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub chkMonografia_Click()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub chkAcuifero_Click()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub chkPrognosis_Click()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub chkPrograma_Click()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub dtpPrimerMonografia_Change()

'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub dtpWellInformed_Change()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub


Private Sub dtpDMAFinalPermit_Change()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub dtpFechaEntregaEiaXConsultoraAOxy_Change()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub dtpSiteVisitWithDMAConducted_Change()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub dtpFechaPedidoETIA_Change()
'llama al evento calcular valores automaticos
  On Error GoTo ErrorHandler
  
  CalcularValoresAutomaticos
  
ErrorHandler:
  ErrorHandler
End Sub

Private Function ObtenerTipoYacimiento() As String
'Obtiene el tipo de yacimiento segun que opt se ha clickeado
  On Error GoTo ErrorHandler
  
  If optDesarrollo.Value = True Then
    ObtenerTipoYacimiento = ""
  ElseIf optExploratorio.Value = True Then
    ObtenerTipoYacimiento = "X"
  ElseIf optAvanzada.Value = True Then
    ObtenerTipoYacimiento = "A"
  End If
  
ErrorHandler:
  ErrorHandler
End Function

Private Sub cmdExportarPlanilla_Click()
'Exporta la grilla a excel
  On Error GoTo ErrorHandler
  Const menExportaciones = 5
  Const mListado = 0
  Const mEstadoLineas = 1
  Const mTTM = 2
  Const mPermitting = 3
  
  frmMenuPrincipal.menExportar(mEstadoLineas).Enabled = (InfoGlobal.IDTipoUsuario = enmTipoUsuarioAdministrador)
  frmMenuPrincipal.menExportar(mTTM).Enabled = (InfoGlobal.IDTipoUsuario = enmTipoUsuarioAdministrador)
  frmMenuPrincipal.menExportar(mPermitting).Enabled = (InfoGlobal.IDTipoUsuario = enmTipoUsuarioAdministrador)
   frmMenuPrincipal.PopupMenu frmMenuPrincipal.menPrincipal(menExportaciones)
    
ErrorHandler:
  ErrorHandler
End Sub


Private Sub Form_Resize()

  Me.SSTab.Width = frmMenuPrincipal.Width - 178
  Me.SSTab.Height = IIf(frmMenuPrincipal.Height - 1300 < 0, 0, frmMenuPrincipal.Height - 1300)
  
  Me.FrameTree.Height = IIf(frmMenuPrincipal.Height - 2900 < 0, 0, frmMenuPrincipal.Height - 2900)
  Me.Tree.Height = IIf(Me.FrameTree.Height - 250 < 0, 0, Me.FrameTree.Height - 250)
  
  Me.Frame3.Left = Me.FrameTree.Left + 50
  Me.Frame3.Width = frmMenuPrincipal.Width - 380
  Me.Frame3.Height = IIf(frmMenuPrincipal.Height - 2900 < 0, 0, frmMenuPrincipal.Height - 2900)
  
  Me.spdCab.Height = IIf(Me.Frame3.Height - 250 < 0, 0, Me.Frame3.Height - 250)
  Me.spdCab.Width = IIf(Me.Frame3.Width - 170 < 0, 0, Me.Frame3.Width - 170)

End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
  
  'CHECK si solapa activa es la 0
  If Me.SSTab.Tab = 0 Then
    
    'SET FOCOS en grilla
    Me.spdCab.SetFocus
    
    'SET active col
    Me.spdCab.SetActiveCell 1, Me.spdCab.Row
    
  End If
  
  
End Sub


'
' PINTA FILA
'
Private Sub pinta_Fila(ByVal lngL As Long, ByVal txtT As String)
    
    'CHECK si no encontro, exit
    If lngL = -1 Then
      Exit Sub
    End If

    'PAINT fila para indicar comienzo de equipo
    
    'ADD fila
    Me.spdCab.MaxRows = Me.spdCab.MaxRows + 1
    Me.spdCab.InsertRows lngL, 1
    
    'MERGE celdas
    Me.spdCab.AddCellSpan 1, lngL, Me.spdCab.DataColCnt - 3, 1
    
    'ASSIGN nombre de equipo
    Me.spdCab.SetText 1, lngL, txtT
    
    'BOLD
    Me.spdCab.Row = lngL
    Me.spdCab.Col = 1
    Me.spdCab.FontBold = True
    
    'SET rango grilla
    Me.spdCab.Col = 1
    Me.spdCab.Col2 = Me.spdCab.DataColCnt
    Me.spdCab.Row = lngL
    Me.spdCab.Row2 = lngL
    
    'CHANGE color aplicado a bloque
    Me.spdCab.BlockMode = True
    Me.spdCab.BackColor = RGB(183, 183, 183)
    Me.spdCab.BlockMode = False
    
End Sub


Private Sub fichaPozo(lngFilaParaEditar)

  Dim lngL As Long
  Dim varDato As Variant
  Dim strT As String
  Dim intI As Integer
  Dim rs As ADODB.Recordset
  
  'GET  datos de columna en donde me indica si hay info a editar o es una linea que
  '     identifica a un grupo, por lo tanto si la edito, da error, entonces hago un Exit
  Me.spdCab.GetText spdCab.GetColFromID("Ubi"), lngFilaParaEditar, varDato
    
  If varDato = "" Then
    Exit Sub
  End If
  
  'GET info completa de grilla
  
  Me.spdCab.GetText spdCab.GetColFromID("IDpozo"), lngFilaParaEditar, varDato
  Me.txtIDpozo = varDato
  
  Me.spdCab.GetText spdCab.GetColFromID("Well ID"), lngFilaParaEditar, varDato
  Me.txtWellID = varDato
  Me.txtWellID.Locked = True
  Me.txtWellID.BackColor = RGB(231, 231, 231)
  
  Me.spdCab.GetText spdCab.GetColFromID("Ubicacion"), lngFilaParaEditar, varDato
  Me.cmbUbicacion = varDato
  Me.cmbUbicacion.Locked = True
  Me.cmbUbicacion.BackColor = RGB(231, 231, 231)

  Me.spdCab.GetText spdCab.GetColFromID("Equipo"), lngFilaParaEditar, varDato
  Me.txtEquipo = varDato
  Me.txtEquipo.Locked = True
  Me.txtEquipo.BackColor = RGB(231, 231, 231)
  
  Me.spdCab.GetText spdCab.GetColFromID("Area"), lngFilaParaEditar, varDato
  Me.txtArea = varDato
  Me.txtArea.Locked = True
  Me.txtArea.BackColor = RGB(231, 231, 231)
  
  Me.spdCab.GetText spdCab.GetColFromID("Pozo Tipo"), lngFilaParaEditar, varDato
  
  Select Case UCase(varDato)
  
    Case "X"
      Me.optExploratorio.Value = 1
    
    Case "A"
      Me.optAvanzada.Value = 1
    
    Case "D"
      Me.optDesarrollo.Value = 1
  
  End Select
  
  Me.optExploratorio.Enabled = False
  Me.optAvanzada.Enabled = False
  Me.optDesarrollo.Enabled = False
  
  
  Me.spdCab.GetText spdCab.GetColFromID("Prospect"), lngFilaParaEditar, varDato
  Me.txtProspect = varDato
  Me.txtProspect.Locked = True
  Me.txtProspect.BackColor = RGB(231, 231, 231)
  
  Me.spdCab.GetText spdCab.GetColFromID("Primer Monografía"), lngFilaParaEditar, varDato
  Me.dtpPrimerMonografia.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpPrimerMonografia = IIf(InStr(varDato, "1900") <> 0, "", varDato)

  Me.spdCab.GetText spdCab.GetColFromID("Monografías"), lngFilaParaEditar, varDato
  Me.txtMonografias = varDato
  
  Me.spdCab.GetText spdCab.GetColFromID("Ultima Monografía"), lngFilaParaEditar, varDato
  Me.dtpUltimaMonografia.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpUltimaMonografia = IIf(InStr(varDato, "1900") <> 0, "", varDato)
  
  Me.spdCab.GetText spdCab.GetColFromID("Well Informado"), lngFilaParaEditar, varDato
  Me.txtWellInformed = varDato
  Me.txtWellInformed.Locked = True
  Me.txtWellInformed.BackColor = RGB(231, 231, 231)
  
  Me.spdCab.GetText spdCab.GetColFromID("Informado Por"), lngFilaParaEditar, varDato
  Me.txtInformedBy = varDato
  Me.txtInformedBy.Locked = True
  Me.txtInformedBy.BackColor = RGB(231, 231, 231)
  
  Me.spdCab.GetText spdCab.GetColFromID("Definitiva"), lngFilaParaEditar, varDato
  Me.chkDefinitiva = varDato

  Me.spdCab.GetText spdCab.GetColFromID("X PDC"), lngFilaParaEditar, varDato
  Me.txtX_PDC = varDato
  Me.txtX_PDC.Locked = True
  Me.txtX_PDC.BackColor = RGB(231, 231, 231)

  Me.spdCab.GetText spdCab.GetColFromID("Y PDC"), lngFilaParaEditar, varDato
  Me.txtY_PDC = varDato
  Me.txtY_PDC.Locked = True
  Me.txtY_PDC.BackColor = RGB(231, 231, 231)

  Me.spdCab.GetText spdCab.GetColFromID("Xsurf WGS84"), lngFilaParaEditar, varDato
  Me.txtX_Pos94 = varDato
  Me.txtX_Pos94.Locked = True
  Me.txtX_Pos94.BackColor = RGB(231, 231, 231)

  Me.spdCab.GetText spdCab.GetColFromID("Ysurf WGS84"), lngFilaParaEditar, varDato
  Me.txtY_Pos94 = varDato
  Me.txtY_Pos94.Locked = True
  Me.txtY_Pos94.BackColor = RGB(231, 231, 231)

  Me.spdCab.GetText spdCab.GetColFromID("Monografía"), lngFilaParaEditar, varDato
  Me.chkMonografia = varDato

  Me.spdCab.GetText spdCab.GetColFromID("Acuifero"), lngFilaParaEditar, varDato
  Me.chkAcuifero = varDato

  Me.spdCab.GetText spdCab.GetColFromID("Plan de 8"), lngFilaParaEditar, varDato
  Me.chkPrognosis = varDato

  Me.spdCab.GetText spdCab.GetColFromID("Plan de 13"), lngFilaParaEditar, varDato
  Me.chkPrograma = varDato

  Me.spdCab.GetText spdCab.GetColFromID("Solicitud"), lngFilaParaEditar, varDato
  Me.dtpFechaSolicitud.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpFechaSolicitud = IIf(InStr(varDato, "1900") <> 0, "", varDato)

  Me.spdCab.GetText spdCab.GetColFromID("Prioridad"), lngFilaParaEditar, varDato
  Me.dtpFechaPrioridad.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpFechaPrioridad = IIf(InStr(varDato, "1900") <> 0, "", varDato)

  Me.spdCab.GetText spdCab.GetColFromID("Preparar Doc."), lngFilaParaEditar, varDato
  Me.txtDocumentoAPreparar = varDato

  Me.spdCab.GetText spdCab.GetColFromID("Field Manifold"), lngFilaParaEditar, varDato
  Me.txtFieldManifold = varDato

  Me.spdCab.GetText spdCab.GetColFromID("Battery Assigned"), lngFilaParaEditar, varDato
  Me.txtBatteryAssigned = varDato

  Me.spdCab.GetText spdCab.GetColFromID("Dictamen Técnico"), lngFilaParaEditar, varDato
  Me.dtpFechaEntregaDictamenTecnico.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpFechaEntregaDictamenTecnico = IIf(InStr(varDato, "1900") <> 0, "", varDato)
  
  Me.spdCab.GetText spdCab.GetColFromID("Producción Inicial"), lngFilaParaEditar, varDato
  Me.txtFirstProd = varDato
  Me.txtFirstProd.Locked = True
  Me.txtFirstProd.BackColor = RGB(231, 231, 231)

  Me.spdCab.GetText spdCab.GetColFromID("TD"), lngFilaParaEditar, varDato
  Me.txtTD = varDato
  Me.txtTD.Locked = True
  Me.txtTD.BackColor = RGB(231, 231, 231)

  Me.spdCab.GetText spdCab.GetColFromID("Total Days"), lngFilaParaEditar, varDato
  Me.txtTotDays = varDato
  Me.txtTotDays.Locked = True
  Me.txtTotDays.BackColor = RGB(231, 231, 231)

  Me.spdCab.GetText spdCab.GetColFromID("Remaining Days"), lngFilaParaEditar, varDato
  Me.txtRemDays = varDato
  Me.txtRemDays.Locked = True
  Me.txtRemDays.BackColor = RGB(231, 231, 231)

  Me.spdCab.GetText spdCab.GetColFromID("Status"), lngFilaParaEditar, varDato
  Me.txtStatus = varDato
  Me.txtStatus.Locked = True
  Me.txtStatus.BackColor = RGB(231, 231, 231)

  Me.spdCab.GetText spdCab.GetColFromID("Start"), lngFilaParaEditar, varDato
  Me.dtpStartDate.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpStartDate = IIf(InStr(varDato, "1900") <> 0, "", varDato)
  Me.dtpStartDate.Enabled = False
  
  Me.spdCab.GetText spdCab.GetColFromID("End"), lngFilaParaEditar, varDato
  Me.dtpEndDate.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpEndDate = IIf(InStr(varDato, "1900") <> 0, "", varDato)
  Me.dtpEndDate.Enabled = False
  
  Me.spdCab.GetText spdCab.GetColFromID("Land Owner"), lngFilaParaEditar, varDato
  Me.txtLandOwner = varDato
  Me.txtLandOwner.Locked = True
  Me.txtLandOwner.BackColor = RGB(231, 231, 231)
  
  Me.spdCab.GetText spdCab.GetColFromID("Land Owner Permit"), lngFilaParaEditar, varDato
  Me.dtpLandOwnerPermitDate.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpLandOwnerPermitDate = IIf(InStr(varDato, "1900") <> 0, "", varDato)
  
  Me.spdCab.GetText spdCab.GetColFromID("Consultora"), lngFilaParaEditar, varDato
  Me.txtConsult = varDato
  
  Me.spdCab.GetText spdCab.GetColFromID("Aprobación Tipo"), lngFilaParaEditar, varDato
  Me.txtType = varDato
  
  Me.spdCab.GetText spdCab.GetColFromID("Pedido EIA"), lngFilaParaEditar, varDato
  Me.dtpFechaPedidoETIA.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpFechaPedidoETIA = IIf(InStr(varDato, "1900") <> 0, "", varDato)
  
  Me.spdCab.GetText spdCab.GetColFromID("Esperada EIA"), lngFilaParaEditar, varDato
  Me.dtpFechaEsperadaETIA.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpFechaEsperadaETIA = IIf(InStr(varDato, "1900") <> 0, "", varDato)
  
  Me.spdCab.GetText spdCab.GetColFromID("Recomendación"), lngFilaParaEditar, varDato
  Me.txtConsultantRecomendation = varDato
  
  Me.spdCab.GetText spdCab.GetColFromID("DMA Final Permit"), lngFilaParaEditar, varDato
  Me.dtpDMAFinalPermit.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpDMAFinalPermit = IIf(InStr(varDato, "1900") <> 0, "", varDato)
  
  Me.spdCab.GetText spdCab.GetColFromID("Estado"), lngFilaParaEditar, varDato
  Me.txtEstado = varDato
  
  Me.spdCab.GetText spdCab.GetColFromID("ID Manifiesto"), lngFilaParaEditar, varDato
  Me.txtIDManifiesto = varDato

  Me.spdCab.GetText spdCab.GetColFromID("Manifiesto"), lngFilaParaEditar, varDato
  Me.dtpFechaManifiesto.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpFechaManifiesto = IIf(InStr(varDato, "1900") <> 0, "", varDato)

  Me.spdCab.GetText spdCab.GetColFromID("Entrega Consultora"), lngFilaParaEditar, varDato
  Me.dtpFechaEntregaEiaXConsultoraAOxy.Value = IIf(InStr(varDato, "1900") <> 0, "", varDato)

  Me.spdCab.GetText spdCab.GetColFromID("EIA Presentado"), lngFilaParaEditar, varDato
  Me.chkEIAPresentado.Value = varDato


  Me.spdCab.GetText spdCab.GetColFromID("Envio a CS"), lngFilaParaEditar, varDato
  Me.dtpFechaEnvioACS.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpFechaEnvioACS = IIf(InStr(varDato, "1900") <> 0, "", varDato)

  Me.spdCab.GetText spdCab.GetColFromID("Presentación DMA"), lngFilaParaEditar, varDato
  Me.dtpFechaPresentacionDMA.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpFechaPresentacionDMA = IIf(InStr(varDato, "1900") <> 0, "", varDato)
  
  Me.spdCab.GetText spdCab.GetColFromID("Presentación SMA"), lngFilaParaEditar, varDato
  Me.dtpFechaPresentacionSMA.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpFechaPresentacionSMA = IIf(InStr(varDato, "1900") <> 0, "", varDato)

  Me.spdCab.GetText spdCab.GetColFromID("Pago Tasa Admin."), lngFilaParaEditar, varDato
  Me.dtpPagoTasaAdministrativa.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpPagoTasaAdministrativa = IIf(InStr(varDato, "1900") <> 0, "", varDato)

  Me.spdCab.GetText spdCab.GetColFromID("Infor. Adicional"), lngFilaParaEditar, varDato
  Me.dtpFechaInfoComplementaria.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpFechaInfoComplementaria = IIf(InStr(varDato, "1900") <> 0, "", varDato)

  Me.spdCab.GetText spdCab.GetColFromID("Informe Técnico"), lngFilaParaEditar, varDato
  Me.txtTechnicalReport = varDato

  Me.spdCab.GetText spdCab.GetColFromID("Tasa Admin. ($)"), lngFilaParaEditar, varDato
  Me.txtTasaAdministrativa = varDato

  Me.spdCab.GetText spdCab.GetColFromID("Tasa Contralor ($)"), lngFilaParaEditar, varDato
  Me.txtTasaContralor = varDato

  Me.spdCab.GetText spdCab.GetColFromID("Estudio ($)"), lngFilaParaEditar, varDato
  Me.txtEstudio = varDato

  Me.spdCab.GetText spdCab.GetColFromID("Adenda ($)"), lngFilaParaEditar, varDato
  Me.txtAdenda = varDato

  Me.spdCab.GetText spdCab.GetColFromID("Inicio DIA"), lngFilaParaEditar, varDato
  Me.dtpFECHAINICIODIA.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpFECHAINICIODIA = IIf(InStr(varDato, "1900") <> 0, "", varDato)
  
  Me.spdCab.GetText spdCab.GetColFromID("Fin DIA"), lngFilaParaEditar, varDato
  Me.dtpFECHAFINDIA.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpFECHAFINDIA = IIf(InStr(varDato, "1900") <> 0, "", varDato)
  
  Me.spdCab.GetText spdCab.GetColFromID("Avance de Obra 50%"), lngFilaParaEditar, varDato
  Me.dtpInformeAvanceObra50PorCiento.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpInformeAvanceObra50PorCiento = IIf(InStr(varDato, "1900") <> 0, "", varDato)
  
  Me.spdCab.GetText spdCab.GetColFromID("Avance de Obra 100%"), lngFilaParaEditar, varDato
  Me.dtpInformeAvanceObra100PorCiento.Value = IIf(varDato = "", Now(), varDato)
  Me.dtpInformeAvanceObra100PorCiento = IIf(InStr(varDato, "1900") <> 0, "", varDato)
    
  Me.spdCab.GetText spdCab.GetColFromID("Evaluación Arqueologica"), lngFilaParaEditar, varDato
  Me.txtInformeEvaluacionArqueologico = varDato
  
  Me.spdCab.GetText spdCab.GetColFromID("Pedido y Recep. EIA"), lngFilaParaEditar, varDato
  Me.txtTiempoEntrePedidoETIAyRecepcionETIA = varDato
  Me.txtTiempoEntrePedidoETIAyRecepcionETIA.BackColor = RGB(216, 237, 223)
  Me.txtTiempoEntrePedidoETIAyRecepcionETIA.Locked = True
  
  Me.spdCab.GetText spdCab.GetColFromID("Recep. EIA y Pres. DMA"), lngFilaParaEditar, varDato
  Me.txtTiempoEntreRecepcionETIAYpresentacionAnteDMA = varDato
  Me.txtTiempoEntreRecepcionETIAYpresentacionAnteDMA.BackColor = RGB(216, 237, 223)
  Me.txtTiempoEntreRecepcionETIAYpresentacionAnteDMA.Locked = True
  
  Me.spdCab.GetText spdCab.GetColFromID("Visita y Aprob. DMA"), lngFilaParaEditar, varDato
  Me.txtTiempoEntreVisitaYAprobacionFinaldeDMA = varDato
  Me.txtTiempoEntreVisitaYAprobacionFinaldeDMA.BackColor = RGB(216, 237, 223)
  Me.txtTiempoEntreVisitaYAprobacionFinaldeDMA.Locked = True
  
  Me.spdCab.GetText spdCab.GetColFromID("Monog. y Pedido EIA"), lngFilaParaEditar, varDato
  Me.txtTiempoEntrePrimerMonografiaYPedidoEIA = varDato
  Me.txtTiempoEntrePrimerMonografiaYPedidoEIA.BackColor = RGB(216, 237, 223)
  Me.txtTiempoEntrePrimerMonografiaYPedidoEIA.Locked = True
  
  Me.spdCab.GetText spdCab.GetColFromID("Monog. Y Recep. EIA"), lngFilaParaEditar, varDato
  Me.txtTiempoEntrePrimerMonografiayRecepcionETIA = varDato
  Me.txtTiempoEntrePrimerMonografiayRecepcionETIA.BackColor = RGB(216, 237, 223)
  Me.txtTiempoEntrePrimerMonografiayRecepcionETIA.Locked = True
  
  Me.spdCab.GetText spdCab.GetColFromID("Pres. DMA y Visita"), lngFilaParaEditar, varDato
  Me.txtTiempoEntrePresentacionAnteDMAyVisita = varDato
  Me.txtTiempoEntrePresentacionAnteDMAyVisita.BackColor = RGB(216, 237, 223)
  Me.txtTiempoEntrePresentacionAnteDMAyVisita.Locked = True
  
  Me.spdCab.GetText spdCab.GetColFromID("Pres. Pozo y DMA"), lngFilaParaEditar, varDato
  Me.txtTiempoEntrePresentacionDePozoYAprobacionFinalDMA = varDato
  Me.txtTiempoEntrePresentacionDePozoYAprobacionFinalDMA.BackColor = RGB(216, 237, 223)
  Me.txtTiempoEntrePresentacionDePozoYAprobacionFinalDMA.Locked = True
  
  Me.spdCab.GetText spdCab.GetColFromID("Entrega EIA y Aprob."), lngFilaParaEditar, varDato
  Me.txtTiempoEntreEntregaEIAaDMAyAprobacion = varDato
  Me.txtTiempoEntreEntregaEIAaDMAyAprobacion.BackColor = RGB(216, 237, 223)
  Me.txtTiempoEntreEntregaEIAaDMAyAprobacion.Locked = True
  
  
  '--------------------------------------------------------------------------------------------
  'GET detalle1
  strT = "select * " & _
         "from EIApozosDatosDetalle1_vw " & _
         "where IDpozo = " & Me.txtIDpozo
  
  Set rs = SQLexec(strT)
  
  'CHECK error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    End
  End If
  
  'BINDING grilla
  Set Me.spdDet1.DataSource = rs
  
  'SET limite en grilla
  Me.spdDet1.MaxCols = rs.Fields.Count
  
  'LOCK grilla completa
  Me.spdDet1.Row = -1
  Me.spdDet1.Col = -1
  Me.spdDet1.Lock = True
  Me.spdDet1.Protect = True
  
  'SET ID de columna para luego poder extraer el dato x ID
  For intI = 0 To rs.Fields.Count - 1
    
    spdDet1.Col = intI + 1
    spdDet1.ColID = rs.Fields(intI).Name
    
  Next
  
  'HIDE columnas que son para trabajar internamente
  Me.spdDet1.Col = 1
  Me.spdDet1.ColHidden = True
  Me.spdDet1.Col = Me.spdDet1.MaxCols
  Me.spdDet1.ColHidden = True
  
  
  '--------------------------------------------------------------------------------------------
  'GET detalle2
  strT = "select * " & _
         "from EIApozosDatosDetalle2_vw " & _
         "where IDpozo = " & Me.txtIDpozo
  
  Set rs = SQLexec(strT)
  
  'CHECK error
  If Not SQLparam.CnErrNumero = -1 Then
    SQLError
    SQLclose
    End
  End If
  
  'BINDING grilla
  Set Me.spdDet2.DataSource = rs
  
  'SET limite en grilla
  Me.spdDet2.MaxCols = rs.Fields.Count
  
  'LOCK grilla completa
  Me.spdDet2.Row = -1
  Me.spdDet2.Col = -1
  Me.spdDet2.Lock = True
  Me.spdDet2.Protect = True
  
  'SET ID de columna para luego poder extraer el dato x ID
  For intI = 0 To rs.Fields.Count - 1
    
    spdDet2.Col = intI + 1
    spdDet2.ColID = rs.Fields(intI).Name
    
  Next
  
  'HIDE columnas que son para trabajar internamente
  Me.spdDet2.Col = 1
  Me.spdDet2.ColHidden = True
  Me.spdDet2.Col = Me.spdDet2.MaxCols
  Me.spdDet2.ColHidden = True
  
End Sub

