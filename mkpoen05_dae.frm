VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form mkpoen05_dae 
   Caption         =   "결재"
   ClientHeight    =   2565
   ClientLeft      =   7935
   ClientTop       =   3030
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   3420
   Begin Threed.SSPanel SSPanel1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      _Version        =   65536
      _ExtentX        =   5953
      _ExtentY        =   4471
      _StockProps     =   15
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txt_dat 
         Height          =   360
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox chk_all 
         Caption         =   "대결"
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   735
      End
      Begin Threed.SSCommand cmd_del 
         Height          =   375
         Left            =   1740
         TabIndex        =   2
         Top             =   2040
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "취소"
      End
      Begin Threed.SSCommand cmd_ins 
         Height          =   375
         Left            =   135
         TabIndex        =   3
         Top             =   2040
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "등록"
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   360
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1770
         _Version        =   65536
         _ExtentX        =   3122
         _ExtentY        =   635
         _StockProps     =   15
         Caption         =   "사장님 결재일자"
         ForeColor       =   4210752
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "결재일자를 입력하세요.(8자리)"
         Height          =   255
         Left            =   460
         TabIndex        =   7
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "결재권자가 부재중일때 긴급히 처리해야 하는 경우 대신하여 결재하며, 문서는 결재권자에게 사후 보고하여야 함."
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   2895
      End
   End
End
Attribute VB_Name = "mkpoen05_dae"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chk_all_Click()
    
    If chk_all.Value = 1 Then
        txt_dat.Text = 99999999
        txt_dat.Enabled = False
    Else
        txt_dat.Text = ""
        txt_dat.Enabled = True
    End If
    
End Sub

Private Sub cmd_del_Click()

    mkpoen05_print.txt_dae.Text = 0
    
    Unload Me
    
End Sub

Private Sub cmd_ins_Click()

    mkpoen05_print.txt_dae.Text = Val(txt_dat.Text)
    
    Unload Me
    
End Sub
