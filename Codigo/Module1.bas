Attribute VB_Name = "Module1"
Option Explicit

Public OutputFile    As String

Public ObjFile       As String

Public NpcFile       As String

Public ObjData()     As ObjDatas

Public NpcData()     As NpcDatas

Public HechizoData() As HechizoDatas

Public MapName()     As String

Public MapDesc()     As String

Public Declare Function WritePrivateProfileString _
               Lib "kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                   ByVal lpKeyName As Any, _
                                                   ByVal lpString As String, _
                                                   ByVal lpFileName As String) As Long

Public Declare Function GetPrivateProfileString _
               Lib "kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                 ByVal lpKeyName As Any, _
                                                 ByVal lpDefault As String, _
                                                 ByVal lpReturnedString As String, _
                                                 ByVal nSize As Long, _
                                                 ByVal lpFileName As String) As Long

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

' Datos de Objetos (Objs.dat)
Public Type ObjDatas

    grhindex As Long ' Índice del gráfico que representa el objeto
    ' Nombre por idioma
    Name     As String       ' Español
    en_name  As String
    pt_name  As String
    fr_name  As String
    it_name  As String
    ' Texto por idioma
    texto     As String      ' Español
    en_texto  As String
    pt_texto  As String
    fr_texto  As String
    it_texto  As String
    ' Descripción / Info por idioma
    Info     As String       ' Español
    en_Info  As String
    pt_Info  As String
    fr_Info  As String
    it_Info  As String
    
    Desc     As String       ' Español
    en_Desc  As String
    pt_Desc  As String
    fr_Desc  As String
    it_Desc  As String
    ' Atributos de combate y uso

    MINDEF As Integer
    MaxDEF As Integer
    MinHit As Long
    MaxHit As Long
    ObjType As Byte
    CreaLuz As String
    CreaParticulaPiso As Integer
    CreaGRH As String
    Hechizo As Integer
    ' Ingredientes y componentes (alquimia, crafteo)
    Raices As Integer
    Cuchara As Integer
    Botella As Integer
    Mortero As Integer
    FrascoAlq As Integer
    FrascoElixir As Integer
    Dosificador As Integer
    Orquidea As Integer
    Carmesi As Integer
    HongoDeLuz As Integer
    Esporas As Integer
    Tuna As Integer
    Cala As Integer
    ColaDeZorro As Integer
    FlorOceano As Integer
    FlorRoja As Integer
    Hierva As Integer
    HojasDeRin As Integer
    HojasRojas As Integer
    SemillasPros As Integer
    Pimiento As Integer
    Madera As Long
    MaderaElfica As Long
    PielLobo As Integer
    PielOsoPardo As Integer
    PielOsoPolar As Integer
    PielLoboNegro As Integer
    PielTigre As Integer
    PielTigreBengala As Integer
    LingH As Integer
    LingP As Integer
    LingO As Integer
    Coal As Integer
    ' Otros atributos
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
    SpellIndex As Integer

End Type

' Datos del NPC (Npcs.dat)
Public Type NpcDatas

    ' Nombre por idioma
    Name     As String        ' Español
    en_name  As String
    pt_name  As String
    fr_name  As String
    it_name  As String
    ' Descripción por idioma
    Desc     As String        ' Español
    en_Desc  As String
    pt_Desc  As String
    fr_Desc  As String
    it_Desc  As String
    ' Atributos generales
    Body          As Integer
    Head          As Integer
    Hp            As Long
    Exp           As Long
    ExpClan       As Long
    Oro           As Long
    MinHit        As Long
    MaxHit        As Long
    NumQuiza      As Byte
    QuizaProb     As Integer
    NoMapInfo     As Byte
    PuedeInvocar  As Byte
    QuizaDropea() As Integer

End Type


' Datos de Hechizos (Hechizos.dat)
Public Type HechizoDatas
    ' Nombre por idioma
    Nombre     As String       ' Español
    en_name    As String
    pt_name    As String
    fr_name    As String
    it_name    As String
    ' Descripción por idioma
    Desc       As String       ' Español
    en_Desc    As String
    pt_Desc    As String
    fr_Desc    As String
    it_Desc    As String
    ' Palabras mágicas
    PalabrasMagicas As String
    ' Mensajes del lanzador por idioma
    HechizeroMsg        As String
    en_HechizeroMsg     As String
    pt_HechizeroMsg     As String
    fr_HechizeroMsg     As String
    it_HechizeroMsg     As String
    ' Mensajes al objetivo por idioma
    TargetMsg           As String
    en_TargetMsg        As String
    pt_TargetMsg        As String
    fr_TargetMsg        As String
    it_TargetMsg        As String
    ' Mensajes propios por idioma
    PropioMsg           As String
    en_PropioMsg        As String
    pt_PropioMsg        As String
    fr_PropioMsg        As String
    it_PropioMsg        As String
    ' Parámetros técnicos
    StaRequerido   As Integer
    ManaRequerido  As Integer
    MinSkill       As Byte
    IconoIndex     As Long
    Cooldown       As Long


End Type


Function ReadField(ByVal Pos As Integer, _
                   ByRef Text As String, _
                   ByVal SepASCII As Byte) As String
        '*****************************************************************
        'Gets a field from a delimited string
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 11/15/2004
        '*****************************************************************
        Dim i          As Long
        Dim LastPos    As Long
        Dim CurrentPos As Long
        Dim delimiter  As String * 1
100     delimiter = Chr$(SepASCII)

102     For i = 1 To Pos
104         LastPos = CurrentPos
106         CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
108     Next i

110     If CurrentPos = 0 Then
112         ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
        Else
114         ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)

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

100     If LenB(Text) = 0 Then Exit Function
102     delimiter = Chr$(SepASCII)
104     curPos = 0
        Do
106         curPos = InStr(curPos + 1, Text, delimiter)
108         count = count + 1
110     Loop While curPos <> 0

112     FieldCount = count

End Function

Public Function GetVar(ByVal File As String, _
                       ByVal Main As String, _
                       ByVal Var As String) As String
        '*****************************************************************
        'Author: Aaron Perkins
        'Last Modify Date: 10/07/2002
        'Get a var to from a text file
        '*****************************************************************
        Dim L        As Long
        Dim Char     As String
        Dim sSpaces  As String 'Input that the program will retrieve
        Dim szReturn As String 'Default value if the string is not found
100     sSpaces = Space$(5000)
102     GetPrivateProfileString Main, Var, vbNullString, sSpaces, Len(sSpaces), File
104     GetVar = RTrim$(sSpaces)
106     GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function

Sub WriteVar(ByVal File As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal Value As String)
        '*****************************************************************
        'Writes a var to a text file
        '*****************************************************************
100     WritePrivateProfileString Main, Var, Value, File

End Sub

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
100     FileExist = (Len(Dir$(File, FileType)) <> 0)

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
100     handle = FreeFile
102     Open OutputFile For Output As handle
104     Close handle
        Exit Sub
Error_Handler:
106     Close handle

End Sub
