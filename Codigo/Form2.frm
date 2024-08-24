VERSION 5.00
Begin VB.Form Form2
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mensages Locales - Optimizacion de protocolo"
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
      Caption         =   "Recargar file"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdReloadList
      Caption         =   "Recargar lista"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton cmdSaveFile
      Caption         =   "Grabar archivo"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   5040
      Width           =   2415
   End
   Begin VB.CommandButton cmdSaveIndex
      Caption         =   "Guardar Index"
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
      Caption         =   "Filtrar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   4150
      Width           =   495
   End
   Begin VB.Label lblUserName
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
Private NumMsg As Integer
Private MsgFile As String
Private arrLocale_SMG() As String

' Constants
Private Const FILE_PATH As String = App.Path & "\..\Recursos\init\LocalMsg.dat"

' Helper functions
Private Function GetMessageIndex(ByVal listItem As String) As Integer
    GetMessageIndex = Val(ReadField(1, listItem, 45))
End Function

Private Function GetMessageText(ByVal listItem As String) As String
    GetMessageText = ReadField(2, listItem, Asc("-"))
End Function

Private Sub LoadMessages()
    ' Load messages from file
    If FileExist(FILE_PATH, vbNormal) Then
        MsgFile = FILE_PATH
        NumMsg = Val(GetVar(MsgFile, "INIT", "NumLocaleMsg"))
        ReDim arrLocale_SMG(1 To NumMsg)

        Dim i As Integer
        For i = 1 To NumMsg
            arrLocale_SMG(i) = GetVar(MsgFile, "Msg", "Msg" & i)
            lstMessages.AddItem i & "-" & arrLocale_SMG(i)
        Next i
    End If
End Sub

Private Sub SaveMessages()
    ' Save messages to file
    Dim arch As String = FILE_PATH
    Dim msg As Integer

    Call WriteVar(arch, "INIT", "NumLocaleMsg", NumMsg)

    For msg = 1 To NumMsg
        DoEvents
        Call WriteVar(arch, "Msg", "Msg" & msg, arrLocale_SMG(msg))
    Next msg
End Sub

Private Sub FilterMessages()
    ' Filter messages based on the filter text
    lstMessages.Clear

    Dim i As Integer
    Dim filterText As String = txtFilter.Text

    If filterText = vbNullString Then
        ' No filter, show all messages
        For i = 1 To NumMsg
            lstMessages.AddItem i & "-" & arrLocale_SMG(i)
        Next i
    Else
        ' Filter messages based on the filter text
        For i = 1 To NumMsg
            If InStr(1, UCase$(arrLocale_SMG(i)), UCase$(filterText)) Then
                lstMessages.AddItem i & "-" & arrLocale_SMG(i)
            End If
        Next i
    End If
End Sub

Private Sub cmdSaveIndex_Click()
    ' Save the selected message index
    Dim selectedIndex As Integer = lstMessages.ListIndex

    If selectedIndex >= 0 Then
        Dim messageIndex As Integer = GetMessageIndex(lstMessages.List(selectedIndex))
        arrLocale_SMG(messageIndex) = txtMessage.Text
        FilterMessages
    Else
        MsgBox "Debes seleccionar un elemento de la lista."
    End If
End Sub

Private Sub cmdSaveFile_Click()
    ' Save messages to file
    SaveMessages
End Sub

Private Sub cmdReloadList_Click()
    ' Reload the message list
    FilterMessages
End Sub

Private Sub cmdReloadFile_Click()
    ' Reload messages from file
    LoadMessages
    txtFilter.Text = ""
End Sub

Private Sub txtFilter_Change()
    ' Filter messages when the filter text changes
    FilterMessages
End Sub

Private Sub Form_Load()
    ' Load messages when the form loads
    LoadMessages
End Sub

Private Sub lstMessages_Click()
    ' Display the selected message text
    Dim selectedIndex As Integer = lstMessages.ListIndex
    If selectedIndex >= 0 Then
        txtMessage.Text = GetMessageText(lstMessages.List(selectedIndex))
    End If
End Sub
