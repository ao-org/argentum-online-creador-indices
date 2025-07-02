VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Type ChildNode
    Key As String
    Value As String

End Type

Private Type MainNode
    Name As String
    values() As ChildNode
    numValues As Integer

End Type

Private fileData() As MainNode
Private MainNodes  As Long

Private Sub Class_Terminate()
    Dim i As Long
    If MainNodes Then

        For i = 1 To MainNodes - 1
            Erase fileData(i).values
        Next i

        Erase fileData

    End If

    MainNodes = 0

End Sub

Public Sub Initialize(ByVal File As String)
        Dim handle As Integer
        Dim Text   As String
        Dim Pos    As Long
100     Call Class_Terminate
102     handle = FreeFile
104     Open File For Input As handle

106     Do Until EOF(handle)
108         Line Input #handle, Text
110         If Len(Text) Then
112             If Left$(Text, 1) = "[" Then
114                 Pos = InStr(2, Text, "]")
116                 If Pos Then
118                     ReDim Preserve fileData(MainNodes) As MainNode
120                     fileData(MainNodes).Name = UCase$(Trim$(mid$(Text, 2, Pos - 2)))
122                     MainNodes = MainNodes + 1

                    End If

                Else
124                 Pos = InStr(2, Text, "=")
126                 If Pos Then
128                     If MainNodes Then

130                         With fileData(MainNodes - 1)
132                             ReDim Preserve .values(.numValues) As ChildNode
134                             .values(.numValues).Value = Right$(Text, Len(Text) - Pos)
136                             .values(.numValues).Key = UCase$(Left$(Text, Pos - 1))
138                             .numValues = .numValues + 1

                            End With

                        End If

                    End If

                End If

            End If

        Loop
140     Close handle
        Dim i As Long
142     If MainNodes Then
144         Call SortMainNodes(0, MainNodes - 1)

146         For i = 0 To MainNodes - 1
148             If fileData(i).numValues Then Call SortChildNodes(fileData(i), 0, fileData(i).numValues - 1)
150         Next i

        End If

End Sub

Private Sub SortChildNodes(ByRef Node As MainNode, _
                           ByVal First As Integer, _
                           ByVal Last As Integer)
        Dim min  As Integer, max As Integer
        Dim comp As String
        Dim temp As ChildNode
100     min = First: max = Last

102     With Node
104         comp = .values((min + max) \ 2).Key

106         Do While min <= max

108             Do While .values(min).Key < comp And min < Last: min = min + 1: Loop

110             Do While .values(max).Key > comp And max > First: max = max - 1: Loop
112             If min <= max Then
114                 temp = .values(min): .values(min) = .values(max): .values(max) = temp
116                 min = min + 1: max = max - 1

                End If

            Loop

        End With

118     If First < max Then SortChildNodes Node, First, max
120     If min < Last Then SortChildNodes Node, min, Last

End Sub

Private Sub SortMainNodes(ByVal First As Integer, ByVal Last As Integer)
        Dim min  As Integer, max As Integer
        Dim comp As String
        Dim temp As MainNode
100     min = First: max = Last
102     comp = fileData((min + max) \ 2).Name

104     Do While min <= max

106         Do While fileData(min).Name < comp And min < Last: min = min + 1: Loop

108         Do While fileData(max).Name > comp And max > First: max = max - 1: Loop
110         If min <= max Then
112             temp = fileData(min): fileData(min) = fileData(max): fileData(max) = temp
114             min = min + 1: max = max - 1

            End If

        Loop
116     If First < max Then SortMainNodes First, max
118     If min < Last Then SortMainNodes min, Last

End Sub

Public Function GetValue(ByVal Main As String, ByVal Key As String) As String
        Dim i As Long, j As Long
100     i = FindMain(UCase$(Main))
102     If i >= 0 Then
104         j = FindKey(fileData(i), UCase$(Key))
106         If j >= 0 Then GetValue = fileData(i).values(j).Value

        End If

End Function

Public Sub ChangeValue(ByVal Main As String, ByVal Key As String, ByVal Value As String)
        Dim i As Long, j As Long
100     i = FindMain(UCase$(Main))
102     If i < 0 Then i = AddNode(Main)
104     j = FindKey(fileData(i), UCase$(Key))
106     If j < 0 Then
108         Call AddKey(Main, Key, Value)
        Else
110         fileData(i).values(j).Value = Value

        End If

End Sub

Private Function FindKey(ByRef Node As MainNode, ByVal Key As String) As Long
        Dim min As Long, max As Long, mid As Long
100     min = 0: max = Node.numValues - 1

102     Do While min <= max
104         mid = (min + max) \ 2
106         If Node.values(mid).Key < Key Then
108             min = mid + 1
110         ElseIf Node.values(mid).Key > Key Then
112             max = max - 1
            Else
114             FindKey = mid: Exit Function

            End If

        Loop
116     FindKey = Not min

End Function

Private Function FindMain(ByVal Name As String) As Long
        Dim min As Long, max As Long, mid As Long
100     min = 0: max = MainNodes - 1

102     Do While min <= max
104         mid = (min + max) \ 2
106         If fileData(mid).Name < Name Then
108             min = mid + 1
110         ElseIf fileData(mid).Name > Name Then
112             max = max - 1
            Else
114             FindMain = mid: Exit Function

            End If

        Loop
116     FindMain = Not min

End Function

Public Function KeyExists(ByVal Name As String) As Boolean
100     KeyExists = FindMain(UCase$(Name)) >= 0

End Function

Private Function AddNode(ByRef Name As String) As Long
        Dim i As Long, MainPos As Long
100     Name = UCase$(Name)
102     ReDim Preserve fileData(MainNodes) As MainNode
104     If MainNodes Then
106         i = MainNodes - 1
108         MainPos = Not FindMain(Name)

110         Do While i >= MainPos
112             fileData(i + 1) = fileData(i)
114             i = i - 1
            Loop
116         fileData(MainPos).numValues = 0
118         Erase fileData(MainPos).values

        End If

120     fileData(MainPos).Name = Name
122     MainNodes = MainNodes + 1
124     AddNode = MainPos

End Function

Private Function AddKey(ByRef Main As String, _
                        ByRef Key As String, _
                        Optional ByRef Value As String = vbNullString) As Long
        Dim MainPos As Long, KeyPos As Long, i As Long
100     Main = UCase$(Main): Key = UCase$(Key)
102     MainPos = FindMain(Main)
104     If MainPos < 0 Then MainPos = AddNode(Main)

106     With fileData(MainPos)
108         ReDim Preserve .values(.numValues) As ChildNode
110         If .numValues > 0 Then
112             i = .numValues - 1
114             KeyPos = Not FindKey(fileData(MainPos), Key)

116             Do While i >= KeyPos
118                 .values(i + 1) = .values(i)
120                 i = i - 1
                Loop

            End If

122         .values(KeyPos).Key = Key
124         .values(KeyPos).Value = Value
126         .numValues = .numValues + 1

        End With

128     AddKey = KeyPos

End Function

Public Sub DumpFile(ByRef File As String)
        Dim hFile As Integer
        Dim i     As Long, j As Long
100     hFile = FreeFile()
102     Open File For Output As hFile

104     For i = 0 To MainNodes - 1

106         With fileData(i)
108             Print #hFile, "[" & .Name & "]"

110             For j = 0 To .numValues - 1
112                 Print #hFile, .values(j).Key & "=" & .values(j).Value
114             Next j

116             Print #hFile, vbNullString

            End With

118     Next i

120     Close hFile

End Sub

Public Function GetAllSectionNames() As Collection
        Dim col As New Collection, i As Long

100     For i = 0 To MainNodes - 1
102         col.Add fileData(i).Name
104     Next i

106     Set GetAllSectionNames = col

End Function

Public Function GetAllKeys(ByVal section As String) As Collection
        Dim col As New Collection, i As Long, j As Long
100     i = FindMain(UCase$(section))
102     If i < 0 Then Exit Function

104     For j = 0 To fileData(i).numValues - 1
106         col.Add fileData(i).values(j).Key
108     Next j

110     Set GetAllKeys = col

End Function

Public Function GetMainNodeCount() As Long
100     GetMainNodeCount = MainNodes

End Function
