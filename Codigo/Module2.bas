Attribute VB_Name = "Module2"
Option Explicit

Public arrLocale_SMG() As String

Public CantMsg         As Integer

Public Function Load_Locales() As Boolean

        On Error GoTo ErrorHandler

        Dim strFile As String

        Dim tmpStr  As String

        Dim intFile As Integer

        Dim i       As Long

100     strFile = App.Path & "\..\Recursos\init\localmsg.dat"
102     ReDim arrLocale_SMG(1 To General_Get_Line_Count(strFile)) As String
104     intFile = FreeFile
106     Open strFile For Input As #intFile

108     Do While Not EOF(intFile)
110         i = i + 1
112         Line Input #intFile, arrLocale_SMG(i)
114         Form2.List1.AddItem (i & "-" & arrLocale_SMG(i))
        Loop
116     CantMsg = i
118     Close #intFile
120     Load_Locales = True
        Exit Function
ErrorHandler:

End Function

Public Function Locale_Parse_ServerMessage(ByVal bytHeader As Byte, _
                                           Optional ByVal strExtra As String = vbNullString) As String

        On Error GoTo ErrorHandler

        Dim strLocale As String

        Dim lngPos    As Long

100     If LenB(strExtra) = 0 Then
102         Locale_Parse_ServerMessage = arrLocale_SMG(bytHeader)
            Exit Function

        End If

104     strLocale = arrLocale_SMG(bytHeader)
106     lngPos = InStr(1, strLocale, "%N")

108     If lngPos > 0 Then
110         Locale_Parse_ServerMessage = Replace$(strLocale, "%N", strExtra)
            Exit Function

        End If

112     lngPos = InStr(1, strLocale, "¬")

114     Do While lngPos > 0
116         strLocale = Replace$(strLocale, mid$(strLocale, lngPos, 2), String_To_Byte(strExtra, CByte(mid$(strLocale, lngPos + 1, 1))))
118         lngPos = InStr(lngPos + 1, strLocale, "¬")
        Loop
120     lngPos = InStr(1, strLocale, "#")

122     Do While lngPos > 0
124         strLocale = Replace$(strLocale, mid$(strLocale, lngPos, 2), String_To_Long(strExtra, CByte(mid$(strLocale, lngPos + 1, 1))))
126         lngPos = InStr(lngPos + 1, strLocale, "#")
        Loop
128     lngPos = InStr(1, strLocale, "&")

130     Do While lngPos > 0
            'nombre del objeto debe ser
            ' strLocale = Replace$(strLocale, mid$(strLocale, lngPos, 2), Locale_UserItem(String_To_Integer(strExtra, CByte(mid$(strLocale, lngPos + 1, 1)))))
132         lngPos = InStr(lngPos + 1, strLocale, "&")
        Loop
134     lngPos = InStr(1, strLocale, "%")

136     If lngPos > 0 Then
138         strLocale = Replace$(strLocale, mid$(strLocale, lngPos, 2), mid$(strExtra, CByte(mid$(strLocale, lngPos + 1, 1))))

        End If

ErrorHandler:
140     Locale_Parse_ServerMessage = strLocale

End Function

Public Function General_Get_Line_Count(ByVal FileName As String) As Long

        '**************************************************************
        'Author: Augusto José Rando
        'Last Modify Date: 6/11/2005
        '
        '**************************************************************
        On Error GoTo ErrorHandler

        Dim N As Integer, tmpStr As String

100     If LenB(FileName) Then
102         N = FreeFile()
104         Open FileName For Input As #N

106         Do While Not EOF(N)
108             General_Get_Line_Count = General_Get_Line_Count + 1
110             Line Input #N, tmpStr
            Loop
112         Close N

        End If

        Exit Function
ErrorHandler:

End Function

Public Function Integer_To_String(ByVal Var As Integer) As String

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 3/12/2005
        '
        '**************************************************************
        Dim temp As String

        'Convertimos a hexa
100     temp = Hex$(Var)

        'Nos aseguramos tenga 4 bytes de largo
102     While Len(temp) < 4

104         temp = "0" & temp
        Wend
        'Convertimos a string
106     Integer_To_String = Chr$(Val("&H" & Left$(temp, 2))) & Chr$(Val("&H" & Right$(temp, 2)))
        Exit Function
ErrorHandler:

End Function

Public Function String_To_Integer(ByRef str As String, ByVal start As Integer) As Integer

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 3/12/2005
        '
        '**************************************************************
        On Error GoTo Error_Handler

        Dim temp_str As String

        'Asergurarse sea válido
100     If Len(str) < start - 1 Or Len(str) = 0 Then Exit Function
        'Convertimos a hexa el valor ascii del segundo byte
102     temp_str = Hex$(Asc(mid$(str, start + 1, 1)))

        'Nos aseguramos tenga 2 bytes (los ceros a la izquierda cuentan por ser el segundo byte)
104     While Len(temp_str) < 2

106         temp_str = "0" & temp_str
        Wend
        'Convertimos a integer
108     String_To_Integer = Val("&H" & Hex$(Asc(mid$(str, start, 1))) & temp_str)
        Exit Function
Error_Handler:

End Function

Public Function Byte_To_String(ByVal Var As Byte) As String
        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 3/12/2005
        'Convierte un byte a string
        '**************************************************************
100     Byte_To_String = Chr$(Val("&H" & Hex$(Var)))
        Exit Function
ErrorHandler:

End Function

Public Function String_To_Byte(ByRef str As String, ByVal start As Integer) As Byte

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 3/12/2005
        '
        '**************************************************************
        On Error GoTo Error_Handler

100     If Len(str) < start Then Exit Function
102     String_To_Byte = Asc(mid$(str, start, 1))
        Exit Function
Error_Handler:

End Function

Public Function Long_To_String(ByVal Var As Long) As String

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 3/12/2005
        '
        '**************************************************************
        'No aceptamos valores que usen los 4 últimos its
100     If Var > &HFFFFFFF Then GoTo ErrorHandler

        Dim temp As String

        'Vemos si el cuarto byte es cero
102     If (Var And &HFF&) = 0 Then Var = Var Or &H80000001

        'Vemos si el tercer byte es cero
104     If (Var And &HFF00&) = 0 Then Var = Var Or &H40000100

        'Vemos si el segundo byte es cero
106     If (Var And &HFF0000) = 0 Then Var = Var Or &H20010000

        'Vemos si el primer byte es cero
108     If Var < &H1000000 Then Var = Var Or &H10000000
        'Convertimos a hexa
110     temp = Hex$(Var)

        'Nos aseguramos tenga 8 bytes de largo
112     While Len(temp) < 8

114         temp = "0" & temp
        Wend
        'Convertimos a string
116     Long_To_String = Chr$(Val("&H" & Left$(temp, 2))) & Chr$(Val("&H" & mid$(temp, 3, 2))) & Chr$(Val("&H" & mid$(temp, 5, 2))) & Chr$(Val("&H" & mid$(temp, 7, 2)))
        Exit Function
ErrorHandler:

End Function

Public Function String_To_Long(ByRef str As String, ByVal start As Integer) As Long

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 3/12/2005
        '
        '**************************************************************
        On Error GoTo ErrorHandler

100     If Len(str) < start - 3 Then Exit Function

        Dim temp_str  As String

        Dim temp_str2 As String

        Dim temp_str3 As String

        'Tomamos los últimos 3 bytes y convertimos sus valroes ASCII a hexa
102     temp_str = Hex$(Asc(mid$(str, start + 1, 1)))
104     temp_str2 = Hex$(Asc(mid$(str, start + 2, 1)))
106     temp_str3 = Hex$(Asc(mid$(str, start + 3, 1)))

        'Nos aseguramos todos midan 2 bytes (los ceros a la izquierda cuentan por ser bytes 2, 3 y 4)
108     While Len(temp_str) < 2

110         temp_str = "0" & temp_str
        Wend

112     While Len(temp_str2) < 2

114         temp_str2 = "0" & temp_str2
        Wend

116     While Len(temp_str3) < 2

118         temp_str3 = "0" & temp_str3
        Wend
        'Convertimos a una única cadena hexa
120     String_To_Long = Val("&H" & Hex$(Asc(mid$(str, start, 1))) & temp_str & temp_str2 & temp_str3)

        'Si el cuarto byte era cero
122     If String_To_Long And &H80000000 Then String_To_Long = String_To_Long Xor &H80000001

        'Si el tercer byte era cero
124     If String_To_Long And &H40000000 Then String_To_Long = String_To_Long Xor &H40000100

        'Si el segundo byte era cero
126     If String_To_Long And &H20000000 Then String_To_Long = String_To_Long Xor &H20010000

        'Si el primer byte era cero
128     If String_To_Long And &H10000000 Then String_To_Long = String_To_Long Xor &H10000000
        Exit Function
ErrorHandler:

End Function
