Attribute VB_Name = "Module1"
Option Explicit

Public OutputFile As String

Public ObjFile       As String
Public NpcFile       As String
Public ObjData()     As ObjDatas
Public NpcData()     As NpcDatas
Public HechizoData() As HechizoDatas
Public MapName()     As String
Public MapDesc()     As String

Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Type ModRaza
    Fuerza As Integer
    Agilidad As Integer
    Inteligencia As Integer
    Constitucion As Integer
    Carisma As Integer
End Type

Public Const NUMRAZAS = 6
Public ModRaza(1 To NUMRAZAS) As ModRaza

Public Sugerencia()           As String

Public QuestName()            As String
Public QuestDesc()            As String
Public QuestFin()             As String
Public QuestNext()            As String
Public QuestPos()             As Integer
Public QuestRepetible()       As Byte

Public RequiredLevel()        As Integer

Public Type ObjDatas

    grhindex As Long ' Indice del grafico que representa el obj
    Name As String
    texto As String
    Info As String
    en_Name As String
    en_texto As String
    en_Info As String
    MINDEF As Integer
    MaxDEF As Integer
    MinHit As Integer
    MaxHit As Integer
    ObjType As Byte
    CreaLuz As String
    CreaParticulaPiso As Integer
    CreaGRH  As String
    Raices As Integer
    Madera As Integer
    MaderaElfica As Integer
    PielLobo As Integer
    PielOsoPardo As Integer
    PielOsoPolar As Integer
    LingH As Integer
    LingP As Integer
    LingO As Integer
    Destruye As Byte
    Proyectil As Byte
    Municiones As Byte
    SkHerreria As Byte
    SkPociones As Byte
    Sksastreria As Byte
    Valor As Long
    Agarrable As Boolean
    Llave As Integer
    Cooldown As Long
    CdType As Integer

End Type

Public Type NpcDatas

    Name As String
    desc As String
    en_Name As String
    en_desc As String
    Body As Integer
    Head As Integer
    Hp As Long
    Exp As Long
    ExpClan As Long
    Oro As Long
    MinHit As Integer
    MaxHit As Integer
    NumQuiza As Byte
    PuedeInvocar As Byte
    QuizaDropea() As Integer

End Type

Public Type HechizoDatas

    Nombre As String ' Indice del grafico que representa el obj
    desc As String
    PalabrasMagicas As String
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
    StaRequerido As Integer
    ManaRequerido As Integer
    MinSkill As Byte
    IconoIndex As Long

End Type

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
    '*****************************************************************
    'Gets a field from a delimited string
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/15/2004
    '*****************************************************************
    Dim i          As Long
    Dim LastPos    As Long
    Dim CurrentPos As Long
    Dim delimiter  As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)

    End If

End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
    '*****************************************************************
    'Gets the number of fields in a delimited string
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 07/29/2007
    '*****************************************************************
    Dim count     As Long
    Dim curPos    As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        count = count + 1
    Loop While curPos <> 0
    
    FieldCount = count

End Function

Public Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Get a var to from a text file
    '*****************************************************************
    Dim L        As Long
    Dim Char     As String
    Dim sSpaces  As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
    
    sSpaces = Space$(5000)
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
    '*****************************************************************
    'Writes a var to a text file
    '*****************************************************************
    writeprivateprofilestring Main, Var, Value, File

End Sub

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Len(Dir$(File, FileType)) <> 0)

End Function

Public Sub Clean_File(ByVal file_path As String)
    '*****************************************************************
    'Author: Juan Martín Dotuyo Dodero
    'Last Modify Date: 10/12/2020 (Jopi)
    'Wipe out the contents of the file
    '*****************************************************************
    On Error GoTo Error_Handler
    
    Dim handle As Integer
    
    'We open the file to delete
    handle = FreeFile
    Open OutputFile For Output As handle
    Close handle

    Exit Sub
    
Error_Handler:
    Close handle
        
End Sub
