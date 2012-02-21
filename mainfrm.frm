VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mainfrm 
   Caption         =   "调查问卷统计系统"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   7065
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "退    出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2070
      TabIndex        =   5
      Top             =   3480
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "选项"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2085
      TabIndex        =   3
      Top             =   2610
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   270
      Left            =   705
      TabIndex        =   2
      Top             =   3225
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton tj 
      Caption         =   "开始统计"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2085
      TabIndex        =   0
      Top             =   1740
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   255
      TabIndex        =   6
      Top             =   120
      Width           =   6675
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3090
      TabIndex        =   4
      Top             =   630
      Width           =   1065
   End
End
Attribute VB_Name = "mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    optionfrm.Show vbModal, Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    File1.Path = docPath
End Sub

Private Sub Form_Load()
    dbName = App.Path
    If Right(dbName, 1) <> "\" Then dbName = dbName & "\"
    dbName = dbName & "data.mdb"
    Set conn = New ADODB.Connection
    docPath = App.Path & "\doc"
    File1.Path = docPath
    Label1.Caption = ""
    Label2.Caption = ""
End Sub

Private Sub tj_Click()
'On Error GoTo errmsg
    Dim errFileName(1000) As String
    Dim errFileNum As Integer
    Dim fileCount
    Dim wApp As Word.Application
    
    Dim answer(21, 10) As Integer  '答题
    
    If Dir(dbName) = "" Then CreateDB
    
    If DirExists(App.Path & "\errFiles") = 0 Then
        MkDir App.Path & "\errFiles"
    End If
    
    
    
    DBConnect
    conn.Execute "delete from tj"
    
    Set wApp = New Word.Application
    
    wApp.Visible = False
    fileCount = File1.ListCount
    
    PBar.Min = 0
    PBar.Max = fileCount
    PBar.Value = PBar.Min
    errFileNum = 0
    For j = 0 To fileCount - 1
    Label1.Caption = File1.List(j)
        wApp.Documents.Open docPath & "\" & File1.List(j)
    'For j = 0 To 0
    'wApp.Documents.Open docPath & "\wj(4).doc"
        
        '1-5题
        For a = 1 To 5
            answer(a, wApp.ActiveDocument.FormFields(a).DropDown.Value) = answer(a, wApp.ActiveDocument.FormFields(a).DropDown.Value) + 1
        Next
        
        '第6题
        For a = 6 To 12
            If wApp.ActiveDocument.FormFields(a).CheckBox.Value Then answer(6, a - 5) = answer(6, a - 5) + 1
        Next
        
        '7
        For a = 15 To 22
            If wApp.ActiveDocument.FormFields(a).CheckBox.Value Then answer(7, a - 14) = answer(7, a - 14) + 1
        Next
        
        '81
        For a = 24 To 28
            If wApp.ActiveDocument.FormFields(a).CheckBox.Value Then answer(8, a - 23) = answer(8, a - 23) + 1
        Next
        
        '82
        For a = 29 To 33
            If wApp.ActiveDocument.FormFields(a).CheckBox.Value Then answer(9, a - 28) = answer(9, a - 28) + 1
        Next
        
        '83
        For a = 34 To 38
            If wApp.ActiveDocument.FormFields(a).CheckBox.Value Then answer(10, a - 33) = answer(10, a - 33) + 1
        Next
        
        '84
        For a = 39 To 43
            If wApp.ActiveDocument.FormFields(a).CheckBox.Value Then answer(11, a - 38) = answer(11, a - 38) + 1
        Next
        
        '85
        For a = 44 To 48
            If wApp.ActiveDocument.FormFields(a).CheckBox.Value Then answer(12, a - 43) = answer(12, a - 43) + 1
        Next
        
        '13
        For a = 49 To 54
            If wApp.ActiveDocument.FormFields(a).CheckBox.Value Then answer(13, a - 48) = answer(8, a - 48) + 1
        Next
        
        '14
        For a = 56 To 62
            If wApp.ActiveDocument.FormFields(a).CheckBox.Value Then answer(14, a - 55) = answer(14, a - 55) + 1
        Next
        
        '15
        For a = 64 To 69
            If wApp.ActiveDocument.FormFields(a).CheckBox.Value Then answer(15, a - 63) = answer(15, a - 63) + 1
        Next
        
        '16
        For a = 71 To 80
            If wApp.ActiveDocument.FormFields(a).CheckBox.Value Then answer(16, a - 70) = answer(16, a - 70) + 1
        Next
        
        '17
        For a = 82 To 90
            If wApp.ActiveDocument.FormFields(a).CheckBox.Value Then answer(17, a - 81) = answer(17, a - 81) + 1
        Next
        
        '18
        For a = 92 To 99
            If wApp.ActiveDocument.FormFields(a).CheckBox.Value Then answer(18, a - 91) = answer(18, a - 91) + 1
        Next
        
        '19
        For a = 101 To 105
            If wApp.ActiveDocument.FormFields(a).CheckBox.Value Then answer(19, a - 100) = answer(19, a - 100) + 1
        Next
        
        '20
        answer(20, wApp.ActiveDocument.FormFields(107).DropDown.Value) = answer(20, wApp.ActiveDocument.FormFields(107).DropDown.Value) + 1
        For a = 108 To 113
            If wApp.ActiveDocument.FormFields(a).CheckBox.Value Then answer(21, a - 107) = answer(21, a - 107) + 1
        Next
    
        Label1.Caption = "问卷数量：" & j + 1 & "份"
        
        PBar.Value = PBar + 1
        Label2.Caption = Int(PBar.Value / PBar.Max * 100) & "%"
        sql = sql & fieldvalue & ")"
    
    wApp.ActiveDocument.Close
      
    Next
    
    For t = 1 To 21
        sql = "insert into tj(rs,aa,ab,ac,ad,ae,af,ag,ah,ai,aj) values("
        fieldvalue = j
        For xx = 1 To 10
            fieldvalue = fieldvalue & "," & answer(t, xx)
        Next
        sql = sql & fieldvalue & ")"
        conn.Execute sql
    Next
    
    'Label2.Caption = "100%"
    
    'Label1.Caption = "正在生成统计表..."
    
    
      
      
    tongji
        
    conn.Close
    Set conn = Nothing
    
errmsg:
    If wApp <> "" Then wApp.Quit
    Set wApp = Nothing
    
    Label1.Caption = "统计完成！保存到" & App.Path & "文件夹。"
    
    errFilemsg = ""
    
Exit Sub

    For i = 1 To errFileNum
        FileCopy docPath & "\" & errFileName(i), App.Path & "\errFiles" & errFileName(i)
        errFilemsg = errFilemsg & errFileName(i) & Chr(13)
    Next
    errFilemsg = "下列文件未统计，已复制到" & App.Path & "\errFiles" & Chr(13) & Chr(13) & errFilemsg
    MsgBox errFilemsg, vbCritical, "错误文件列表"
    
End Sub

Private Sub CreateDB()
    '菜单“工程”-->"引用"-->"Microsoft   ActiveX   Data   Objects   2.8   Library"
    '                    -->  Microsoft   ADO   Ext.2.8   for   DDL   ado   Security
    Dim cat     As ADOX.Catalog
    Set cat = New ADOX.Catalog
    cat.Create ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbName & ";")
    MsgBox "数据库创建成功！"
    Dim tbl     As ADOX.Table
    Set tbl = New ADOX.Table
    tbl.ParentCatalog = cat
    tbl.Name = "tj"
    
    '增加一个自动增长的字段
    Dim col     As ADOX.Column
    Set col = New ADOX.Column
    col.ParentCatalog = cat
    col.Type = ADOX.DataTypeEnum.adInteger       '   //   必须先设置字段类型
    col.Name = "id"
    col.Properties("Jet OLEDB:Allow Zero Length").Value = False
    col.Properties("AutoIncrement").Value = True
    tbl.Columns.Append col, ADOX.DataTypeEnum.adInteger, 0
    
    '增加一个文本字段
    Dim col2     As ADOX.Column
    Set col2 = New ADOX.Column
    col2.ParentCatalog = cat
    col2.Name = "dw"   '单位名称
    col2.Properties("Jet OLEDB:Allow Zero Length").Value = True
    tbl.Columns.Append col2, ADOX.DataTypeEnum.adVarChar, 50
    
    '增加一个数值型字段
    Dim col4     As ADOX.Column
    Set col4 = New ADOX.Column
    col4.ParentCatalog = cat
    col4.Name = "rs"   '人数
    tbl.Columns.Append col4, ADOX.DataTypeEnum.adVarChar, 5
    
    '增加一个数值型字段
    Dim col5     As ADOX.Column
    Set col5 = New ADOX.Column
    col5.ParentCatalog = cat
    col5.Type = ADOX.DataTypeEnum.adInteger
    col5.Name = "aa"   '答案A
    tbl.Columns.Append col5, ADOX.DataTypeEnum.adInteger
    
    '增加一个数值型字段
    Dim col6     As ADOX.Column
    Set col6 = New ADOX.Column
    col6.ParentCatalog = cat
    col6.Type = ADOX.DataTypeEnum.adInteger
    col6.Name = "ab"   '
    tbl.Columns.Append col6, ADOX.DataTypeEnum.adInteger
    
    '增加一个数值型字段
    Dim col7     As ADOX.Column
    Set col7 = New ADOX.Column
    col7.ParentCatalog = cat
    col7.Type = ADOX.DataTypeEnum.adInteger
    col7.Name = "ac"   '
    tbl.Columns.Append col7, ADOX.DataTypeEnum.adInteger
    
    '增加一个数值型字段
    Dim col8     As ADOX.Column
    Set col8 = New ADOX.Column
    col8.ParentCatalog = cat
    col8.Type = ADOX.DataTypeEnum.adInteger
    col8.Name = "ad"   '1d
    tbl.Columns.Append col8, ADOX.DataTypeEnum.adInteger
    
    '增加一个数值型字段
    Dim col9     As ADOX.Column
    Set col9 = New ADOX.Column
    col9.ParentCatalog = cat
    col9.Type = ADOX.DataTypeEnum.adInteger
    col9.Name = "ae"   '2a
    tbl.Columns.Append col9, ADOX.DataTypeEnum.adInteger
    
    '增加一个数值型字段
    Dim col10     As ADOX.Column
    Set col10 = New ADOX.Column
    col10.ParentCatalog = cat
    col10.Type = ADOX.DataTypeEnum.adInteger
    col10.Name = "af"   '2b
    tbl.Columns.Append col10, ADOX.DataTypeEnum.adInteger
    
    '增加一个数值型字段
    Dim col11     As ADOX.Column
    Set col11 = New ADOX.Column
    col11.ParentCatalog = cat
    col11.Type = ADOX.DataTypeEnum.adInteger
    col11.Name = "ag"   '2c
    tbl.Columns.Append col11, ADOX.DataTypeEnum.adInteger
    
    '增加一个数值型字段
    Dim col12     As ADOX.Column
    Set col12 = New ADOX.Column
    col12.ParentCatalog = cat
    col12.Type = ADOX.DataTypeEnum.adInteger
    col12.Name = "ah"   '2d
    tbl.Columns.Append col12, ADOX.DataTypeEnum.adInteger
    
    '增加一个文本字段
    Dim col13     As ADOX.Column
    Set col13 = New ADOX.Column
    col13.ParentCatalog = cat
    col13.Name = "ai"   '2e
    col13.Properties("Jet OLEDB:Allow Zero Length").Value = True
    tbl.Columns.Append col13, ADOX.DataTypeEnum.adVarChar, 255
    
    '增加一个数值型字段
    Dim col14     As ADOX.Column
    Set col14 = New ADOX.Column
    col14.ParentCatalog = cat
    col14.Type = ADOX.DataTypeEnum.adInteger
    col14.Name = "aj"   '3a
    tbl.Columns.Append col14, ADOX.DataTypeEnum.adInteger
    
    '增加一个货币型字段
    'Dim col4     As ADOX.Column
    'Set col4 = New ADOX.Column
    'col4.ParentCatalog = cat
    'col4.Type = ADOX.DataTypeEnum.adCurrency
    'col4.Name = "xx"
    'tbl.Columns.Append col4, ADOX.DataTypeEnum.adCurrency
    
    '增加一个OLE字段
    'Dim col5     As ADOX.Column
    'Set col5 = New ADOX.Column
    'col5.ParentCatalog = cat
    'col5.Type = ADOX.DataTypeEnum.adLongVarBinary
    'col5.Name = "OLD_FLD"
    'tbl.Columns.Append col5, ADOX.DataTypeEnum.adLongVarBinary
    
    '增加一个数值型字段
    'Dim col3     As ADOX.Column
    'Set col3 = New ADOX.Column
    'col3.ParentCatalog = cat
    'col3.Type = ADOX.DataTypeEnum.adDouble
    'col3.Name = "ll"
    'tbl.Columns.Append col3, ADOX.DataTypeEnum.adDouble
    'Dim p     As ADOX.Property
    'For Each p In col3.Properties
    '      Debug.Print p.Name & ":" & p.Value & ":" & p.Type & ":" & p.Attributes
    'Next
    
    '设置主键
    tbl.Keys.Append "PrimaryKey", ADOX.KeyTypeEnum.adKeyPrimary, "id", "", ""
    cat.Tables.Append tbl
    MsgBox "数据库表：" + tbl.Name + "已经创建成功！"
    Set tbl = Nothing
    Set cat = Nothing
    
End Sub

'连接ACCESS数据库
Sub DBConnect()
    strconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbName
    If conn.State <> 0 Then conn.Close
    conn.Open strconn
    
End Sub

Private Sub tongji()
    Dim t1a, t1b, t1c, t1d, t1e As Long
    Dim t2a, t2b, t2c, t2d, t2f As Long
    Dim t3a, t3b, t3c, t3d, t3e As Long
    Dim t2e, t4, t5 As String
    
    Dim count, fp As Long
    
    Dim wordapp As Word.Application
    
    Dim a(21, 10) As Integer
    Dim zrs As Integer '总人数
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sql = "select * from tj"
    rs.Open sql, conn, 1, 1
    
    Set wordapp = New Word.Application
    wordapp.Visible = False
    wordapp.Documents.Open App.Path & "\tjb.doc"
    
    zrs = rs("rs")
    i = 0
    Do While Not rs.EOF
        i = i + 1
        a(i, 1) = rs("aa")
        a(i, 2) = rs("ab")
        a(i, 3) = rs("ac")
        a(i, 4) = rs("ad")
        a(i, 5) = rs("ae")
        a(i, 6) = rs("af")
        a(i, 7) = rs("ag")
        a(i, 8) = rs("ah")
        a(i, 9) = rs("ai")
        a(i, 10) = rs("aj")
        rs.MoveNext
    Loop
    
    count = rs.RecordCount
    rs.Close
    Set rs = Nothing
        wordapp.ActiveDocument.Tables(1).Cell(1, 1).Range.Text = Space(20) & "单位：                           调查人数：" & zrs & "                  "
        
    For i = 1 To 7
    
        wordapp.ActiveDocument.Tables(1).Cell(2 + i, 2).Range.Text = a(i, 1)
        wordapp.ActiveDocument.Tables(1).Cell(2 + i, 3).Range.Text = a(i, 2)
        wordapp.ActiveDocument.Tables(1).Cell(2 + i, 4).Range.Text = a(i, 3)
        wordapp.ActiveDocument.Tables(1).Cell(2 + i, 5).Range.Text = a(i, 4)
        wordapp.ActiveDocument.Tables(1).Cell(2 + i, 6).Range.Text = a(i, 5)
        wordapp.ActiveDocument.Tables(1).Cell(2 + i, 7).Range.Text = a(i, 6)
        wordapp.ActiveDocument.Tables(1).Cell(2 + i, 8).Range.Text = a(i, 7)
        wordapp.ActiveDocument.Tables(1).Cell(2 + i, 9).Range.Text = a(i, 8)
        wordapp.ActiveDocument.Tables(1).Cell(2 + i, 10).Range.Text = a(i, 9)
        wordapp.ActiveDocument.Tables(1).Cell(2 + i, 11).Range.Text = a(i, 10)
    Next
    
    
    For i = 9 To 13
    
        wordapp.ActiveDocument.Tables(1).Cell(2 + i, 3).Range.Text = a(i - 1, 1)
        wordapp.ActiveDocument.Tables(1).Cell(2 + i, 4).Range.Text = a(i - 1, 2)
        wordapp.ActiveDocument.Tables(1).Cell(2 + i, 5).Range.Text = a(i - 1, 3)
        wordapp.ActiveDocument.Tables(1).Cell(2 + i, 6).Range.Text = a(i - 1, 4)
        wordapp.ActiveDocument.Tables(1).Cell(2 + i, 7).Range.Text = a(i - 1, 5)
    Next
    
    For i = 13 To 19
    
        wordapp.ActiveDocument.Tables(1).Cell(4 + i, 2).Range.Text = a(i, 1)
        wordapp.ActiveDocument.Tables(1).Cell(4 + i, 3).Range.Text = a(i, 2)
        wordapp.ActiveDocument.Tables(1).Cell(4 + i, 4).Range.Text = a(i, 3)
        wordapp.ActiveDocument.Tables(1).Cell(4 + i, 5).Range.Text = a(i, 4)
        wordapp.ActiveDocument.Tables(1).Cell(4 + i, 6).Range.Text = a(i, 5)
        wordapp.ActiveDocument.Tables(1).Cell(4 + i, 7).Range.Text = a(i, 6)
        wordapp.ActiveDocument.Tables(1).Cell(4 + i, 8).Range.Text = a(i, 7)
        wordapp.ActiveDocument.Tables(1).Cell(4 + i, 9).Range.Text = a(i, 8)
        wordapp.ActiveDocument.Tables(1).Cell(4 + i, 10).Range.Text = a(i, 9)
        wordapp.ActiveDocument.Tables(1).Cell(4 + i, 11).Range.Text = a(i, 10)
    Next
    
    For i = 20 To 21
    
        wordapp.ActiveDocument.Tables(1).Cell(5 + i, 3).Range.Text = a(i, 1)
        wordapp.ActiveDocument.Tables(1).Cell(5 + i, 4).Range.Text = a(i, 2)
        wordapp.ActiveDocument.Tables(1).Cell(5 + i, 5).Range.Text = a(i, 3)
        wordapp.ActiveDocument.Tables(1).Cell(5 + i, 6).Range.Text = a(i, 4)
        wordapp.ActiveDocument.Tables(1).Cell(5 + i, 7).Range.Text = a(i, 5)
        wordapp.ActiveDocument.Tables(1).Cell(5 + i, 8).Range.Text = a(i, 6)
    Next
    
    
    
    
    wordapp.ActiveDocument.SaveAs App.Path & "\广州市中小学教师继续教育工作调查问卷答案汇总表.doc"
    wordapp.ActiveDocument.Close
    wordapp.Quit
    Set wordapp = Nothing
    
End Sub
Public Function DirExists(ByVal strDirName As String) As Integer
    Const strWILDCARD$ = "*.*"
       
    Dim strDummy     As String
    
    On Error Resume Next
    If Trim(strDirName) = "" Then
          DirExists = 0
          Exit Function
    End If
    strDummy = Dir$(strDirName & strWILDCARD, vbDirectory)
    DirExists = Not (strDummy = vbNullString)
              
    Err = 0
End Function

