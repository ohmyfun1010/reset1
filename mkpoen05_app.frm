VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form mkpoen05_app 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   1560
   ClientTop       =   1905
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   7725
   Begin Threed.SSPanel SSPanel1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _Version        =   65536
      _ExtentX        =   13573
      _ExtentY        =   9551
      _StockProps     =   15
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSCommand cmd_cancle 
         Height          =   375
         Left            =   6360
         TabIndex        =   5
         Top             =   4920
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "초기화"
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   4695
         Left            =   3360
         TabIndex        =   1
         Top             =   120
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
         _ExtentY        =   8281
         _StockProps     =   14
         Caption         =   "부서원 리스트"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin FPSpreadADO.fpSpread spd_sablist 
            Height          =   4215
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   3975
            _Version        =   458752
            _ExtentX        =   7011
            _ExtentY        =   7435
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
            MaxCols         =   5
            MaxRows         =   0
            SpreadDesigner  =   "mkpoen05_app.frx":0000
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   4695
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   8281
         _StockProps     =   14
         Caption         =   "전체부서"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSComctlLib.TreeView trv2 
            Height          =   4215
            Left            =   100
            TabIndex        =   4
            Top             =   300
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   7435
            _Version        =   393217
            Style           =   7
            Appearance      =   1
         End
      End
   End
End
Attribute VB_Name = "mkpoen05_app"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpCol As Long
Dim tmprow As Long
Dim tmpprt As Integer

Private Sub cmd_cancle_Click()
            
    On Error GoTo err_rtn
    
    Dim resp        As Integer
        
    resp = MsgBox("결재자를 초기화 하시겠습니까??", vbQuestion + vbYesNo)
    If resp = vbNo Then Exit Sub
            
    Ws.BeginTrans
        
    If tmpprt = 2 Then
    
        mkpoen05_print.sprd_print2.Col = tmpCol
        mkpoen05_print.sprd_print2.row = 2
        
        mkpoen05_print.sprd_print2.TypeButtonText = "click"
        mkpoen05_print.sprd_print2.CellTag = 0
        
        sss = "update oth_applist set"
        If tmpCol = 26 Then sss = sss & " apl_1sab = 0"
        If tmpCol = 30 Then sss = sss & " apl_2sab = 0"
        If tmpCol = 34 Then sss = sss & " apl_3sab = 0"
        sss = sss & " where apl_table = 'man_tooltesthd'"
        sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
        sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
        
        db.Execute sss, 64
        
    ElseIf tmpprt = 1 Then
        
        mkpoen05_print.sprd_print1.Col = tmpCol
        mkpoen05_print.sprd_print1.row = tmprow
        
        mkpoen05_print.sprd_print1.TypeButtonText = "Click"
        mkpoen05_print.sprd_print1.CellTag = 0
        
        sss = "update oth_applist set"
        If tmpCol = 26 And tmprow = 2 Then sss = sss & " apl_1sab = 0"
        If tmpCol = 30 And tmprow = 2 Then sss = sss & " apl_2sab = 0"
        If tmpCol = 34 And tmprow = 2 Then sss = sss & " apl_3sab = 0"
        If tmpCol = 34 And tmprow = 33 Then sss = sss & " apl_m1sab = 0"
        If tmpCol = 38 And tmprow = 33 Then sss = sss & " apl_m2sab = 0"
        If tmpCol = 34 And tmprow = 37 Then sss = sss & " apl_m3sab = 0"
        If tmpCol = 38 And tmprow = 37 Then sss = sss & " apl_m4sab = 0"
        sss = sss & " where apl_table = 'man_tooltesthd'"
        sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
        sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
        
        db.Execute sss, 64
        
    End If
    
    Ws.CommitTrans
    
    Unload Me
    
    Exit Sub

err_rtn:
    Ws.Rollback
    MsgBox (Err.Description)
    
End Sub

Private Sub Form_Load()
    
    Dim root_code As String
    
    sss = " select sok_code,sok_name,sok_next,sok_level,sok_upcd "
    sss = sss & " from peo_sokcd "
    sss = sss & " where sok_yn = 'Y' "
    If tmpCol = 26 And tmprow = 2 Then sss = sss & " start with sok_code = 'H000' "
    If tmpCol = 30 And tmprow = 2 Then sss = sss & " start with sok_code = 'H100' "
    If tmpCol = 34 And tmprow = 2 Then sss = sss & " start with sok_code = 'H100' "
    If (tmpCol = 34 Or tmpCol = 38) And (tmprow = 33 Or tmprow = 37) Then sss = sss & " start with sok_code = 'AAAA' "
    sss = sss & " connect by prior sok_code = sok_upcd "
    sss = sss & " order siblings by sok_seq "
    
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Not Rs.RecordCount < 1 Then
        
        Set nodex = trv2.Nodes.Add(, Rs!sok_code, Rs!sok_code, Rs!sok_name)
        Rs.MoveNext
        '
        Do While Not Rs.EOF
        
            root_code = Rs!sok_upcd
            trv2.Nodes.Add root_code, tvwChild, Rs!sok_code, Rs!sok_name & "(" & Rs!sok_code & ")"
       
            Rs.MoveNext
        
        Loop
            
    End If
    
    nodex.Expanded = True
    
    Rs.Close
    
End Sub

Public Sub index_Send(Col As Long, row As Long, prt As Integer)
    tmpCol = Col
    tmprow = row
    tmpprt = prt
End Sub


Private Sub spd_sablist_Click(ByVal Col As Long, ByVal row As Long)
    
    On Error GoTo err_rtn
    
    Dim resp        As Integer
    Dim appsab      As Integer
    Dim memo        As String   '쪽지발송 메모
    Dim title       As String   '쪽지발송 타이틀
    
    If Len(mkpoen05_print.txt_dat1) < 8 Then
        msg = "등록번호를 확인하세요."
        Exit Sub
    End If
    
    If Len(mkpoen05_print.txt_seq1) = 0 Then
        msg = "등록번호를 확인하세요."
        Exit Sub
    End If
      
    resp = MsgBox("결재자를 지정하시겠습니까??", vbQuestion + vbYesNo)
    If resp = vbNo Then Exit Sub
    
    appsab = 0
    
    Ws.BeginTrans
    
    '================================================================
    'print1 번(test 2개)=============================================
    '================================================================
    
    '담당부서
    
    If tmpCol = 26 And tmprow = 2 And tmpprt = 1 Then
        
        mkpoen05_print.sprd_print1.Col = 26
        mkpoen05_print.sprd_print1.row = 2
        
        spd_sablist.Col = 3
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print1.TypeButtonText = spd_sablist.Text
        
        spd_sablist.Col = 5
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print1.CellTag = spd_sablist.Text
        appsab = spd_sablist.Text
        
        sss = "update oth_applist set"
        sss = sss & " apl_1sab = " & spd_sablist.Text
        sss = sss & " where apl_table = 'man_tooltesthd'"
        sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
        sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
        
        db.Execute sss, 64
    
    ElseIf tmpCol = 30 And tmprow = 2 And tmpprt = 1 Then
        
        mkpoen05_print.sprd_print1.Col = 30
        mkpoen05_print.sprd_print1.row = 2
        
        spd_sablist.Col = 3
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print1.TypeButtonText = spd_sablist.Text
        
        spd_sablist.Col = 5
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print1.CellTag = spd_sablist.Text
        appsab = spd_sablist.Text
        
        sss = "update oth_applist set"
        sss = sss & " apl_2sab = " & spd_sablist.Text
        sss = sss & " where apl_table = 'man_tooltesthd'"
        sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
        sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
        
        db.Execute sss, 64
        
    ElseIf tmpCol = 34 And tmprow = 2 And tmpprt = 1 Then
        
        mkpoen05_print.sprd_print1.Col = 34
        mkpoen05_print.sprd_print1.row = 2
        
        spd_sablist.Col = 3
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print1.TypeButtonText = spd_sablist.Text
        
        spd_sablist.Col = 5
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print1.CellTag = spd_sablist.Text
        appsab = spd_sablist.Text
        
        sss = "update oth_applist set"
        sss = sss & " apl_3sab = " & spd_sablist.Text
        sss = sss & " where apl_table = 'man_tooltesthd'"
        sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
        sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
        
        db.Execute sss, 64
        
    '관련부서
        
    ElseIf tmpCol = 34 And tmprow = 33 And tmpprt = 1 Then
        
        mkpoen05_print.sprd_print1.Col = 34
        mkpoen05_print.sprd_print1.row = 33
        
        spd_sablist.Col = 3
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print1.TypeButtonText = spd_sablist.Text
        
        spd_sablist.Col = 5
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print1.CellTag = spd_sablist.Text
        appsab = spd_sablist.Text
        
        sss = "update oth_applist set"
        sss = sss & " apl_m1sab = " & spd_sablist.Text
        sss = sss & " where apl_table = 'man_tooltesthd'"
        sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
        sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
        
        db.Execute sss, 64
        
    ElseIf tmpCol = 38 And tmprow = 33 And tmpprt = 1 Then
        
        mkpoen05_print.sprd_print1.Col = 38
        mkpoen05_print.sprd_print1.row = 33
        
        spd_sablist.Col = 3
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print1.TypeButtonText = spd_sablist.Text
        
        spd_sablist.Col = 5
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print1.CellTag = spd_sablist.Text
        appsab = spd_sablist.Text
        
        sss = "update oth_applist set"
        sss = sss & " apl_m2sab = " & spd_sablist.Text
        sss = sss & " where apl_table = 'man_tooltesthd'"
        sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
        sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
        
        db.Execute sss, 64
    
     ElseIf tmpCol = 34 And tmprow = 37 And tmpprt = 1 Then
        
        mkpoen05_print.sprd_print1.Col = 34
        mkpoen05_print.sprd_print1.row = 37
        
        spd_sablist.Col = 3
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print1.TypeButtonText = spd_sablist.Text
        
        spd_sablist.Col = 5
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print1.CellTag = spd_sablist.Text
        appsab = spd_sablist.Text
        
        sss = "update oth_applist set"
        sss = sss & " apl_m3sab = " & spd_sablist.Text
        sss = sss & " where apl_table = 'man_tooltesthd'"
        sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
        sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
        
        db.Execute sss, 64
        
     ElseIf tmpCol = 38 And tmprow = 37 And tmpprt = 1 Then
        
        mkpoen05_print.sprd_print1.Col = 38
        mkpoen05_print.sprd_print1.row = 37
        
        spd_sablist.Col = 3
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print1.TypeButtonText = spd_sablist.Text
        
        spd_sablist.Col = 5
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print1.CellTag = spd_sablist.Text
        appsab = spd_sablist.Text
        
        sss = "update oth_applist set"
        sss = sss & " apl_m4sab = " & spd_sablist.Text
        sss = sss & " where apl_table = 'man_tooltesthd'"
        sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
        sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
        
        db.Execute sss, 64
    
    End If
    
    '================================================================
    'print2 번(test 3개)=============================================
    '================================================================
    
    '담당부서
    
    If tmpCol = 26 And tmprow = 2 And tmpprt = 2 Then
        
        mkpoen05_print.sprd_print2.Col = 26
        mkpoen05_print.sprd_print2.row = 2
        
        spd_sablist.Col = 3
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print2.TypeButtonText = spd_sablist.Text
        
        spd_sablist.Col = 5
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print2.CellTag = spd_sablist.Text
        appsab = spd_sablist.Text
        
        sss = "update oth_applist set"
        sss = sss & " apl_1sab = " & spd_sablist.Text
        sss = sss & " where apl_table = 'man_tooltesthd'"
        sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
        sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
        
        db.Execute sss, 64
    
    ElseIf tmpCol = 30 And tmprow = 2 And tmpprt = 2 Then
        
        mkpoen05_print.sprd_print2.Col = 30
        mkpoen05_print.sprd_print2.row = 2
        
        spd_sablist.Col = 3
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print2.TypeButtonText = spd_sablist.Text
        
        spd_sablist.Col = 5
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print2.CellTag = spd_sablist.Text
        appsab = spd_sablist.Text
        
        sss = "update oth_applist set"
        sss = sss & " apl_2sab = " & spd_sablist.Text
        sss = sss & " where apl_table = 'man_tooltesthd'"
        sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
        sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
        
        db.Execute sss, 64
        
    ElseIf tmpCol = 34 And tmprow = 2 And tmpprt = 2 Then
        
        mkpoen05_print.sprd_print2.Col = 34
        mkpoen05_print.sprd_print2.row = 2
        
        spd_sablist.Col = 3
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print2.TypeButtonText = spd_sablist.Text
        
        spd_sablist.Col = 5
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print2.CellTag = spd_sablist.Text
        appsab = spd_sablist.Text
        
        sss = "update oth_applist set"
        sss = sss & " apl_3sab = " & spd_sablist.Text
        sss = sss & " where apl_table = 'man_tooltesthd'"
        sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
        sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
        
        db.Execute sss, 64
        
    '관련부서
        
    ElseIf tmpCol = 34 And tmprow = 33 And tmpprt = 2 Then
        
        mkpoen05_print.sprd_print2.Col = 34
        mkpoen05_print.sprd_print2.row = 33
        
        spd_sablist.Col = 3
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print2.TypeButtonText = spd_sablist.Text
        
        spd_sablist.Col = 5
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print2.CellTag = spd_sablist.Text
        appsab = spd_sablist.Text
        
        sss = "update oth_applist set"
        sss = sss & " apl_m1sab = " & spd_sablist.Text
        sss = sss & " where apl_table = 'man_tooltesthd'"
        sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
        sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
        
        db.Execute sss, 64
        
    ElseIf tmpCol = 38 And tmprow = 33 And tmpprt = 2 Then
        
        mkpoen05_print.sprd_print2.Col = 38
        mkpoen05_print.sprd_print2.row = 33
        
        spd_sablist.Col = 3
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print2.TypeButtonText = spd_sablist.Text
        
        spd_sablist.Col = 5
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print2.CellTag = spd_sablist.Text
        appsab = spd_sablist.Text
        
        sss = "update oth_applist set"
        sss = sss & " apl_m2sab = " & spd_sablist.Text
        sss = sss & " where apl_table = 'man_tooltesthd'"
        sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
        sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
        
        db.Execute sss, 64
    
     ElseIf tmpCol = 34 And tmprow = 37 And tmpprt = 2 Then
        
        mkpoen05_print.sprd_print2.Col = 34
        mkpoen05_print.sprd_print2.row = 37
        
        spd_sablist.Col = 3
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print2.TypeButtonText = spd_sablist.Text
        
        spd_sablist.Col = 5
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print2.CellTag = spd_sablist.Text
        appsab = spd_sablist.Text
        
        sss = "update oth_applist set"
        sss = sss & " apl_m3sab = " & spd_sablist.Text
        sss = sss & " where apl_table = 'man_tooltesthd'"
        sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
        sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
        
        db.Execute sss, 64
        
     ElseIf tmpCol = 38 And tmprow = 37 And tmpprt = 2 Then
        
        mkpoen05_print.sprd_print2.Col = 38
        mkpoen05_print.sprd_print2.row = 37
        
        spd_sablist.Col = 3
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print2.TypeButtonText = spd_sablist.Text
        
        spd_sablist.Col = 5
        spd_sablist.row = row
        
        mkpoen05_print.sprd_print2.CellTag = spd_sablist.Text
        appsab = spd_sablist.Text
        
        sss = "update oth_applist set"
        sss = sss & " apl_m4sab = " & spd_sablist.Text
        sss = sss & " where apl_table = 'man_tooltesthd'"
        sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
        sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
        
        db.Execute sss, 64
    
        
    End If
    
    mkpoen05_print.msg = "결재자가 지정되었습니다."
    
    Ws.CommitTrans
    
    Unload Me
    
    Exit Sub

err_rtn:
    Ws.Rollback
    MsgBox (Err.Description)
    
End Sub

Private Sub trv2_Click()
    
    Dim tmpsokcode As String
    
    tmpsokcode = trv2.SelectedItem.Key
    
    sss = "select sin_name,jik_name,sok_name,sin_sab from peo_sinbun,peo_jiksuda,peo_sokcd where sin_sok = '" & tmpsokcode & "'"
    sss = sss & " and sin_jik = jik_up and sin_sok = sok_code and sin_taedt = 0  order by sin_jik,sin_entdt"
    
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    
    spd_sablist.MaxRows = 0
    
    If Not Rs.RecordCount < 1 Then
    
        With spd_sablist
            Do While Not Rs.EOF
                .MaxRows = .MaxRows + 1: .row = .MaxRows
            
                .Col = 2: If Not IsNull(Rs!sok_name) Then .Text = Rs!sok_name
                .Col = 3: If Not IsNull(Rs!sin_name) Then .Text = Rs!sin_name
                .Col = 4: If Not IsNull(Rs!jik_name) Then .Text = Rs!jik_name
                .Col = 5: If Not IsNull(Rs!sin_sab) Then .Text = Val(Rs!sin_sab)
            Rs.MoveNext
            Loop
        End With
    End If
    
    Rs.Close
 
    
End Sub
