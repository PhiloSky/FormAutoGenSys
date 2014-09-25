VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "FAGS"
   ClientHeight    =   3030
   ClientLeft      =   1530
   ClientTop       =   1515
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "生成"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "自动生成系统"
      BeginProperty Font 
         Name            =   "仿宋"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'试验记录自动生成系统
'消防应急灯具（照明灯具、标志灯具）

Sub Command1_Click()
    Label1.Caption = "生成中……"
    Dim wApp As Word.Application
    Dim wDoc As Word.Document
    Dim wSel As Word.Selection
    Set wApp = New Word.Application

    Dim str As String
    Dim fname As String
    Dim num As String
    Dim mdl As String
    
    Open App.Path & "/任务单.csv" For Input As #1
    
    Do While Not EOF(1)
        Line Input #1, str
        If str = "2014照明," Then
            fname = "照明.doc"
        ElseIf str = "2014标志," Then
            fname = "标志.doc"
        Else
            num = Left(str, 9)
            mdl = Right(str, (Len(str) - 10))
            
            Set wDoc = Documents.Open(FileName:=App.Path & "/模板" & fname)
            Set wSel = Documents.Application.Selection
            
            With wSel.Find
                .Text = "123456789"
                .Replacement.Text = num
            End With
            wSel.Find.Execute Replace:=wdReplaceAll
            
            With wSel.Find
                .Text = "ABCDEFG"
                .Replacement.Text = mdl
            End With
            
            wSel.Find.Execute Replace:=wdReplaceAll
            wSel.GoTo wdGoToBookmark, , , "样品照片"
            wSel.InlineShapes.AddPicture FileName:=App.Path & "/" & num & ".jpg"
            
            wDoc.SaveAs App.Path & ("/" & num & "-" & fname)
            wDoc.Close
        End If
    Loop
    Close #1
    wApp.Quit
    Set wSel = Nothing
    Set wDoc = Nothing
    Set wApp = Nothing
    
    Label1.Caption = "已完成"
End Sub

Private Sub Command2_Click()
    End
End Sub
