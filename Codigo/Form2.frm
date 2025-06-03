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
Private SelectedLangCode As String
Private SelectedLangName As String
Private arrLocale_SMG() As String

' Declaraciones API ANSI para leer y escribir archivos INI
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String, _
     ByVal lpFileName As String) As Long

Private Sub Form_Load()
    Dim inputLang As String
    Dim i As Integer

    inputLang = InputBox( _
        "Seleccione el idioma:" & vbCrLf & _
        "1 = Espaï¿½ol" & vbCrLf & _
        "2 = Inglï¿½s" & vbCrLf & _
        "3 = Portuguï¿½s" & vbCrLf & _
        "4 = Francï¿½s" & vbCrLf & _
        "5 = Italiano", _
        "Idioma")

    Select Case Trim(inputLang)
        Case "1": SelectedLangCode = "ES": SelectedLangName = "Espaï¿½ol"
        Case "2": SelectedLangCode = "EN": SelectedLangName = "Inglï¿½s"
        Case "3": SelectedLangCode = "PT": SelectedLangName = "Portuguï¿½s"
        Case "4": SelectedLangCode = "FR": SelectedLangName = "Francï¿½s"
        Case "5": SelectedLangCode = "IT": SelectedLangName = "Italiano"
        Case Else
            MsgBox "Selecciï¿½n invï¿½lida.", vbCritical
            Unload Me
            Exit Sub
    End Select

    MsgFile = App.Path & "\..\Recursos\init\" & SelectedLangCode & "_LocalMsg.dat"

    If FileExist(MsgFile, vbNormal) Then
        NumMsg = Val(GetIniValue("INIT", "NumLocale" & SelectedLangCode & "_Msg", MsgFile))
        ReDim arrLocale_SMG(1 To NumMsg)

        For i = 1 To NumMsg
            arrLocale_SMG(i) = GetIniValue(SelectedLangCode & "_MSG", "Msg" & i, MsgFile)
            List1.AddItem i & "-" & arrLocale_SMG(i)
        Next i
    Else
        MsgBox "Archivo no encontrado: " & MsgFile, vbExclamation
    End If
End Sub

Private Sub Command1_Click()
    If List1.ListIndex < 0 Then
        MsgBox "Debes seleccionar un elemento de la lista."
        Exit Sub
    End If

    arrLocale_SMG(Val(ReadField(1, List1.List(List1.ListIndex), 45))) = Text1.Text
    Call Command3_Click
End Sub

Private Sub Command2_Click()

    Dim i As Integer

    Call WriteIniValue("INIT", "NumLocale" & SelectedLangCode & "_Msg", NumMsg, MsgFile)

    For i = 1 To NumMsg
        DoEvents
        Call WriteIniValue(SelectedLangCode & "_MSG", "Msg" & i, arrLocale_SMG(i), MsgFile)
    Next i


    MsgBox "Mensajes guardados correctamente en " & MsgFile, vbInformation
End Sub

Private Sub Command3_Click()

    Dim i As Integer
    List1.Clear


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

    Filtro.Text = ""
    Call Command3_Click
End Sub

Private Sub Filtro_Change()
    Call Command3_Click

End Sub

Private Sub List1_Click()
    Text1.Text = ReadField(2, List1.Text, Asc("-"))
End Sub

' Funciones para leer desde archivos INI (ANSI)
Private Function GetIniValue(ByVal section As String, ByVal Key As String, ByVal FileName As String) As String
    Dim buffer As String * 1024
    Dim length As Long
    length = GetPrivateProfileString(section, Key, "", buffer, Len(buffer), FileName)
    GetIniValue = Left$(buffer, length)
End Function

' Funciï¿½n para escribir en archivos INI (ANSI)
Private Function WriteIniValue(ByVal section As String, ByVal Key As String, ByVal Value As String, ByVal FileName As String) As Boolean
    WriteIniValue = (WritePrivateProfileString(section, Key, Value, FileName) <> 0)
End Function


