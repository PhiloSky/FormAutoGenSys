VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "FAGS"
   ClientHeight    =   3030
   ClientLeft      =   1530
   ClientTop       =   1515
   ClientWidth     =   4575
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4575
   Begin VB.CommandButton Command3 
      Caption         =   "��ȡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "�Զ�����ϵͳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�����¼�Զ�����ϵͳ
'����������Ӧ���ƾߣ������ƾߡ���־�ƾߣ�

Sub Command1_Click()
    Label1.Caption = "�����С���"
    Dim wApp As Word.Application
    Dim wDoc As Word.Document
    Dim wSel As Word.Selection
    Set wApp = New Word.Application

    Dim str As String
    Dim fname As String
    Dim num As String
    Dim mdl As String
    
    fname = "����"      '�ƾ����ͳ�ʼ��
    
    Open App.Path & "\����.csv" For Input As #1
    
    Do While Not EOF(1)
        Line Input #1, str
        If str = "" Then
            fname = "��־"      '�ƾ����͸ı�
        Else
            num = Left(str, 9)
            mdl = Right(str, (Len(str) - 10))
            
            Set wDoc = Documents.Open(FileName:=App.Path & "\ģ��" & fname & ".doc")
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
            wSel.GoTo wdGoToBookmark, , , "��Ʒ��Ƭ"
            wSel.InlineShapes.AddPicture FileName:=App.Path & "\δ��ӡ" & fname & "\" & num & ".jpg"
            
            wDoc.SaveAs App.Path & "\δ��ӡ" & fname & "\" & num & "-" & fname & ".doc"
            wDoc.Close
        End If
    Loop
    Close #1
    wApp.Quit
    Set wSel = Nothing
    Set wDoc = Nothing
    Set wApp = Nothing
    
    Command3.Enabled = False
    Label1.Caption = " ��������ɡ� �����˳�"
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    Label1.Caption = "��ȡ�С���"
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim num As String
 
    Open App.Path & "\����.csv" For Output As #2
    Close #2
    Open App.Path & "\����.csv" For Append As #2
    Set fso = CreateObject("scripting.filesystemobject")
    
    Set folder = fso.getfolder(App.Path & "\δ��ӡ����")

    For Each file In folder.Files
       num = Right(file, 13)
       num = Left(num, 9) & ","
       Print #2, num
    Next
    
    Print #2,
    
    Set folder = fso.getfolder(App.Path & "\δ��ӡ��־")
    
    For Each file In folder.Files
       num = Right(file, 13)
       num = Left(num, 9) & ","
       Print #2, num
    Next
    
    Close #2
    Set fso = Nothing
    Set folder = Nothing
    Command1.Enabled = False
    Label1.Caption = " ����ȡ��ɡ� ���ֶ������ͺ�"
End Sub
