VERSION 5.00
Begin VB.Form Form2
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Local Messages - Protocol Optimization"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6030
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFilter
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdReloadFile
      Caption         =   "Reload File"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdReloadList
      Caption         =   "Reload List"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton cmdSaveFile
      Caption         =   "Save File"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   5040
      Width           =   2415
   End
   Begin VB.CommandButton cmdSaveIndex
      Caption         =   "Save Index"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox txtMessage
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   5535
   End
   Begin VB.ListBox lstMessages
      Height          =   2790
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label lblFilter
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Filter"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   4150
      Width           =   495
   End
   Begin VB.Label lblPlaceholder
      Caption         =   "[%N] = UserName"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Module-level variables
Private numMessages As Integer
Private messageFile As String
Private messages() As String

' Function to save the selected message index
Private Sub SaveMessageIndex()
    On Error GoTo ErrorHandler
    
    If lstMessages.ListIndex < 0 Then
        ShowMessageBox "You must select an item from the list."
        Exit Sub
    End If
    
    Dim selectedIndex As Integer
    selectedIndex = Val(ReadField(1, lstMessages.List(lstMessages.ListIndex), 45))
    messages(selectedIndex) = txtMessage.Text
    ReloadMessageList
    
    Exit Sub
    
ErrorHandler:
    LogError Err.Number, Err.Description, "SaveMessageIndex"
End Sub

' Function to save the messages to a file
Private Sub SaveMessagesToFile()
    On Error GoTo ErrorHandler
    
    Dim filePath As String
    filePath = App.Path & "\..\Recursos\init\LocalMsg.dat"
    
    WriteVar filePath, "INIT", "NumLocaleMsg", numMessages
    
    Dim i As Integer
    For i = 1 To numMessages
        DoEvents
        WriteVar filePath, "Msg", "Msg" & i, messages(i)
    Next i
    
    Exit Sub
    
ErrorHandler:
    LogError Err.Number, Err.Description, "SaveMessagesToFile"
End Sub

' Function to reload the message list
Private Sub ReloadMessageList()
    On Error GoTo ErrorHandler
    
    lstMessages.Clear
    
    If Len(txtFilter.Text) = 0 Then
        Dim i As Integer
        For i = 1 To numMessages
            lstMessages.AddItem i & "-" & messages(i)
        Next i
    Else
        Dim i As Integer
        For i = 1 To numMessages
            If InStr(1, UCase$(messages(i)), UCase$(txtFilter.Text)) Then
                lstMessages.AddItem i & "-" & messages(i)
            End If
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError Err.Number, Err.Description, "ReloadMessageList"
End Sub

' Function to reload the message file
Private Sub ReloadMessageFile()
    On Error GoTo ErrorHandler
    
    lstMessages.Clear
    
    Dim filePath As String
    filePath = App.Path & "\..\Recursos\init\LocalMsg.dat"
    
    If FileExist(filePath, vbNormal) Then
        messageFile = filePath
        numMessages = Val(GetVar(messageFile, "INIT", "NumLocaleMsg"))
        txtFilter.Text = ""
        ReDim messages(1 To numMessages)
        
        Dim i As Integer
        For i = 1 To numMessages
            messages(i) = GetVar(messageFile, "Msg", "Msg" & i)
            lstMessages.AddItem i & "-" & messages(i)
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError Err.Number, Err.Description, "ReloadMessageFile"
End Sub

' Event handler for the Filter textbox change
Private Sub FilterTextChanged()
    ReloadMessageList
End Sub

' Event handler for the form load
Private Sub Form_Load()
    ReloadMessageFile
End Sub

' Event handler for the message list click
Private Sub MessageListClicked()
    If lstMessages.ListIndex >= 0 Then
        txtMessage.Text = ReadField(2, lstMessages.Text, Asc("-"))
    End If
End Sub

' Function to show a message box
Private Sub ShowMessageBox(ByVal message As String)
    MsgBox message
End Sub

' Function to log errors
Private Sub LogError(ByVal errorNumber As Long, ByVal errorDescription As String, ByVal functionName As String)
    ' Implement your error logging logic here
    ' For example, you could write the error details to a log file or display them in a message box
    Dim errorMessage As String
    errorMessage = "Error in " & functionName & ": " & errorNumber & " - " & errorDescription
    Debug.Print errorMessage
End Sub

' Event handlers for the command buttons
Private Sub cmdSaveIndex_Click()
    SaveMessageIndex
End Sub

Private Sub cmdSaveFile_Click()
    SaveMessagesToFile
End Sub

Private Sub cmdReloadList_Click()
    ReloadMessageList
End Sub

Private Sub cmdReloadFile_Click()
    ReloadMessageFile
End Sub
