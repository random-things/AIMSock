VERSION 5.00
Object = "{A651B0EA-0B17-4C1E-AECA-DA083CD6169F}#9.0#0"; "AIMSock.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMining 
      Height          =   285
      Left            =   4680
      TabIndex        =   9
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CheckBox chkAFK 
      Caption         =   "AFK"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Qs 
      Caption         =   "?s"
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdCheap 
      Caption         =   "Cheap"
      Height          =   255
      Left            =   6000
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Timer timerQueue 
      Interval        =   10000
      Left            =   4680
      Top             =   4320
   End
   Begin VB.ListBox lstQueue 
      Height          =   840
      Left            =   0
      TabIndex        =   5
      Top             =   4680
      Width           =   4575
   End
   Begin VB.ComboBox cmbBuddy 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   4320
      Width           =   4575
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   3960
      Width           =   5895
   End
   Begin VB.TextBox txtStuff 
      Height          =   3735
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   7335
   End
   Begin AIMSock.AIMSockOCX AIMSockOCX1 
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   5280
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bSearching As Boolean
Private bSearchName As String

Private Function AddText(ByVal SomeString As String)
    txtStuff.Text = txtStuff.Text & SomeString & vbCrLf
    txtStuff.SelStart = Len(txtStuff.Text)
End Function

Private Sub AIMSockOCX1_GotIM(ByVal ScreenName As String, ByVal WarningLevel As Integer, ByVal Message As String)
    AddText ScreenName & " (" & WarningLevel & "): " & StripHTML(Message)

    If chkAFK.Value = vbChecked Then
        lstQueue.AddItem ScreenName & "ÿI am currently mining data on <b>" & txtMining.Text & "</b> right now.  Please ask your question again later."
        Exit Sub
    End If

    Dim i As Integer, bFound As Boolean
    For i = 0 To cmbBuddy.ListCount - 1
        If cmbBuddy.List(i) = ScreenName Then
            bFound = True
        End If
    Next i
    
    If bFound = False Then
        cmbBuddy.AddItem ScreenName
    End If
    
    If UCase$(Left$(StripHTML(Message), 3)) = "Q: " Then
        bSearchName = ScreenName
    End If
End Sub

Private Sub AIMSockOCX1_OnLoggedOnAs(ByVal ScreenName As String)
    AddText "Logged on as: " & ScreenName
End Sub

Private Sub chkAFK_Click()
    If chkAFK.Value = vbChecked Then
        AIMSockOCX1.AwayMessage = "I am currently mining data on <b>" & txtMining.Text & "</b>.  Please ask your question again later."
        AIMSockOCX1.SetProfileAway
    Else
        AIMSockOCX1.AwayMessage = ""
        AIMSockOCX1.SetProfileAway
    End If
End Sub

Private Sub cmdCheap_Click()
    lstQueue.AddItem cmbBuddy.Text & "ÿCurrently answering a question for " & bSearchName & ".  Try again shortly."
End Sub

Private Sub cmdSend_Click()
    lstQueue.AddItem cmbBuddy.Text & "ÿA: " & txtSend.Text
    txtSend.Text = ""
End Sub

Private Sub Form_Load()
    AIMSockOCX1.ScreenName = "The Answer Buddy"
    AIMSockOCX1.Password = "kewl"
    AIMSockOCX1.Profile = "Hello, and welcome to The Answer Buddy.  I am an AIM Bot that is designed to answer your questions.  To use me, preface your question with 'Q: ', so for example send 'Q: What color is the sky?' without the quotes.  I will reply eventually, although sometimes it takes up to a minute because of my queue speed.<br><br>Things that I will not discuss:<br>-My creator<br>-My 'intelligence'<br>-Modifying my own code"
    AIMSockOCX1.Connect "login.oscar.aol.com", 5190
End Sub

Public Function StripHTML(ByVal TheString As String) As String
    Dim Done As Boolean
    Done = False
    
    Dim LeftPos As Long, RightPos As Long, TagName As String
    Do Until Done
        LeftPos = InStr(1, TheString, "<")
        If LeftPos = 0 Then
            StripHTML = TheString
            Exit Function
        End If
        RightPos = InStr(LeftPos, TheString, ">")
        
        Dim TempString As String, WholeString As String
        WholeString = Mid(TheString, LeftPos, RightPos - LeftPos + 1)
        TempString = Mid(TheString, LeftPos + 1, RightPos - LeftPos - 1)
        TagName = Split(TempString, " ")(0)
        
        'Select Case UCase$(TagName)
        '    Case "HTML", "FONT", "/HTML", "/FONT", "BODY", "/BODY", "B", "/B", "A", "/A", "I", "/I", "U", "/U"
                TheString = Replace(TheString, WholeString, "")
        'End Select
    Loop
End Function

Private Sub Qs_Click()
    lstQueue.AddItem cmbBuddy.Text & "ÿQuestions must be prefaced with 'Q: '"
End Sub

Private Sub timerQueue_Timer()
    If lstQueue.ListCount > 0 Then
        AIMSockOCX1.SendMessage Split(lstQueue.List(0), "ÿ")(0), Split(lstQueue.List(0), "ÿ")(1)
        lstQueue.RemoveItem 0
    End If
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        lstQueue.AddItem cmbBuddy.Text & "ÿA: " & txtSend.Text
        txtSend.Text = ""
    End If
End Sub
