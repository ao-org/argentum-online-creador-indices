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
   Begin VB.TextBox Filtro 
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Recargar file"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Recargar lista"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar archivo"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   5040
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar Index"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   5535
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label2 
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
   Begin VB.Label Label1 
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

Private NumMsg  As Integer
Private MsgFile As String

Private Sub Command1_Click()

    If List1.ListIndex < 0 Then
        MsgBox "Debes seleccionar un elemento de la lista."
        Exit Sub

    End If

    arrLocale_SMG(Val(ReadField(1, List1.List(List1.ListIndex), 45))) = Text1.Text
    Call Command3_Click

End Sub

Private Sub Command2_Click()
    Dim arch As String
    Dim msg  As Integer
    arch = App.Path & "\..\Recursos\init\" & "LocalMsg.dat"
    
    Call WriteVar(arch, "INIT", "NumLocaleMsg", NumMsg)
    
    For msg = 1 To NumMsg
        DoEvents
        Call WriteVar(arch, "Msg", "Msg" & msg, arrLocale_SMG(msg))
        
    Next msg

End Sub

Private Sub Command3_Click()
    List1.Clear
    
    Dim i As Integer

    If Filtro.Text = vbNullString Then
        
        For i = 1 To NumMsg
    
            List1.AddItem i & "-" & arrLocale_SMG(i)
    
        Next i

    Else

        For i = 1 To NumMsg
    
            If InStr(1, UCase$(arrLocale_SMG(i)), UCase$(Filtro.Text)) Then
                List1.AddItem i & "-" & arrLocale_SMG(i)

            End If
    
        Next i

    End If

End Sub

Private Sub Command4_Click()
    List1.Clear
    Dim i As Integer

    If FileExist(App.Path & "\..\Recursos\init\LocalMsg.dat", vbNormal) Then

        MsgFile = App.Path & "\..\Recursos\init\LocalMsg.dat"
        NumMsg = Val(GetVar(MsgFile, "INIT", "NumLocaleMsg"))

        Filtro.Text = ""

        ReDim arrLocale_SMG(1 To NumMsg) As String
    
        For i = 1 To NumMsg
        
            arrLocale_SMG(i) = GetVar(MsgFile, "Msg", "Msg" & i)
            List1.AddItem i & "-" & arrLocale_SMG(i)

        Next i

    End If

End Sub

Private Sub Filtro_Change()
    Call Command3_Click

End Sub

Private Sub Form_Load()
    Dim i As Integer

    If FileExist(App.Path & "\..\Recursos\init\LocalMsg.dat", vbNormal) Then

        MsgFile = App.Path & "\..\Recursos\init\LocalMsg.dat"
        NumMsg = Val(GetVar(MsgFile, "INIT", "NumLocaleMsg"))

        ReDim arrLocale_SMG(1 To NumMsg) As String
    
        For i = 1 To NumMsg

            arrLocale_SMG(i) = GetVar(MsgFile, "Msg", "Msg" & i)
            List1.AddItem i & "-" & arrLocale_SMG(i)

        Next i

    End If

End Sub

Private Sub List1_Click()
    Text1.Text = ReadField(2, List1.Text, Asc("-"))

End Sub

