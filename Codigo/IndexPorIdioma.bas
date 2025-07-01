Attribute VB_Name = "IndexPorIdioma"

Public Sub DumpLocalIndexPorIdioma(ByVal origen As clsIniReader)
        Dim i                      As Long
        Dim LangPrefix(1 To 5)     As String
        Dim LangSuffix(1 To 5)     As String
        Dim langMsgSection(1 To 5) As String
        Dim langWriters(1 To 5)    As clsIniReader
        Dim clavesGlobales         As Object
        Dim asignado               As Boolean
        Dim Valor                  As String
        Dim procesado              As Boolean
        Dim claveNorm              As Variant
        Dim normalClave            As String
100     LangPrefix(1) = "ES_": LangSuffix(1) = "sp": langMsgSection(1) = "SP_MSG"
102     LangPrefix(2) = "EN_": LangSuffix(2) = "en": langMsgSection(2) = "EN_MSG"
104     LangPrefix(3) = "PT_": LangSuffix(3) = "pt": langMsgSection(3) = "PT_MSG"
106     LangPrefix(4) = "FR_": LangSuffix(4) = "fr": langMsgSection(4) = "FR_MSG"
108     LangPrefix(5) = "IT_": LangSuffix(5) = "it": langMsgSection(5) = "IT_MSG"

110     For i = 1 To 5
112         Set langWriters(i) = New clsIniReader
114     Next i

116     Set clavesGlobales = CreateObject("Scripting.Dictionary")
        clavesGlobales.Add "BODY", 1
        clavesGlobales.Add "HEAD", 1
        clavesGlobales.Add "EXP", 1
        clavesGlobales.Add "HP", 1
        clavesGlobales.Add "MAXHIT", 1
        clavesGlobales.Add "MINHIT", 1
        clavesGlobales.Add "ORO", 1
        clavesGlobales.Add "NUMQUIZA", 1
        clavesGlobales.Add "QUIZADROPEA1", 1
        clavesGlobales.Add "QUIZADROPEA2", 1
        clavesGlobales.Add "QUIZADROPEA3", 1
        clavesGlobales.Add "QUIZADROPEA4", 1
        clavesGlobales.Add "QUIZADROPEA5", 1
        clavesGlobales.Add "QUIZADROPEA6", 1
        clavesGlobales.Add "QUIZADROPEA7", 1
        clavesGlobales.Add "QUIZADROPEA8", 1
        clavesGlobales.Add "QUIZADROPEA9", 1
        clavesGlobales.Add "QUIZADROPEA10", 1
        clavesGlobales.Add "QUIZADROPEA11", 1
        clavesGlobales.Add "QUIZADROPEA12", 1
        clavesGlobales.Add "QUIZADROPEA13", 1
        clavesGlobales.Add "QUIZADROPEA14", 1
        clavesGlobales.Add "QUIZADROPEA15", 1
        clavesGlobales.Add "QUIZADROPEA16", 1
        clavesGlobales.Add "QUIZAPROB", 1
        
        clavesGlobales.Add "MAXDEF", 1
        clavesGlobales.Add "MINDEF", 1
        clavesGlobales.Add "LINGH", 1
        clavesGlobales.Add "LINGP", 1
        clavesGlobales.Add "LINGO", 1
        clavesGlobales.Add "SKHERRERIA", 1
        clavesGlobales.Add "SKSASTRERIA", 1
        clavesGlobales.Add "PIELOSOPARDO", 1
        clavesGlobales.Add "PIELTIGREBENGALA", 1
        clavesGlobales.Add "PIELOSOPOLAR", 1
        clavesGlobales.Add "SKPOCIONES", 1
        clavesGlobales.Add "RAICES", 1
        clavesGlobales.Add "PIELTIGRE", 1
        clavesGlobales.Add "PIELLOBO", 1
        clavesGlobales.Add "LLAVE", 1
        clavesGlobales.Add "AGARRABLE", 1
        clavesGlobales.Add "GIVEEXPCLAN", 1
        
        clavesGlobales.Add "FRASCOALQ", 1
        clavesGlobales.Add "HECHIZO", 1
        clavesGlobales.Add "MORTERO", 1
        clavesGlobales.Add "SEMILLASPROS", 1
        clavesGlobales.Add "PROYECTIL", 1
        clavesGlobales.Add "MUNICIONES", 1
        clavesGlobales.Add "INFO", 1
                                
118     clavesGlobales.Add "NOMBRE", 1
120     clavesGlobales.Add "NAME", 1
122     clavesGlobales.Add "TEXTO", 1
124     clavesGlobales.Add "DESC", 1
126     clavesGlobales.Add "COOLDOWN", 1
128     clavesGlobales.Add "STAREQUERIDO", 1
130     clavesGlobales.Add "MANAREQUERIDO", 1
132     clavesGlobales.Add "MINSKILL", 1
134     clavesGlobales.Add "ICONOINDEX", 1
136     clavesGlobales.Add "TARGET", 1
138     clavesGlobales.Add "TIPO", 1
140     clavesGlobales.Add "WAV", 1
142     clavesGlobales.Add "MATERIALIZACANT", 1
144     clavesGlobales.Add "MATERIALIZAOBJ", 1
146     clavesGlobales.Add "NOWIKI", 1
148     clavesGlobales.Add "FXGRH", 1
150     clavesGlobales.Add "LOOPS", 1
152     clavesGlobales.Add "VALOR", 1
154     clavesGlobales.Add "OBJTYPE", 1
156     clavesGlobales.Add "GRHINDEX", 1
158     clavesGlobales.Add "PALABRASMAGICAS", 1
160     clavesGlobales.Add "NEXTQUEST", 1
162     clavesGlobales.Add "POSMAP", 1
164     clavesGlobales.Add "REPETIBLE", 1
166     clavesGlobales.Add "REQUIREDLEVEL", 1
168     clavesGlobales.Add "MADERA", 1
170     clavesGlobales.Add "MAPA", 1
        ' Agregadas claves de razas y sugerencias también:
172     clavesGlobales.Add "ELFOAGILIDAD", 1
174     clavesGlobales.Add "ELFOCARISMA", 1
176     clavesGlobales.Add "ELFOCONSTITUCION", 1
178     clavesGlobales.Add "ELFOFUERZA", 1
180     clavesGlobales.Add "ELFOINTELIGENCIA", 1
182     clavesGlobales.Add "ELFOOSCUROAGILIDAD", 1
184     clavesGlobales.Add "ELFOOSCUROCARISMA", 1
186     clavesGlobales.Add "ELFOOSCUROCONSTITUCION", 1
188     clavesGlobales.Add "ELFOOSCUROFUERZA", 1
190     clavesGlobales.Add "ELFOOSCUROINTELIGENCIA", 1
192     clavesGlobales.Add "ENANOAGILIDAD", 1
194     clavesGlobales.Add "ENANOCARISMA", 1
196     clavesGlobales.Add "ENANOCONSTITUCION", 1
198     clavesGlobales.Add "ENANOFUERZA", 1
200     clavesGlobales.Add "ENANOINTELIGENCIA", 1
202     clavesGlobales.Add "GNOMOAGILIDAD", 1
204     clavesGlobales.Add "GNOMOCARISMA", 1
206     clavesGlobales.Add "GNOMOCONSTITUCION", 1
208     clavesGlobales.Add "GNOMOFUERZA", 1
210     clavesGlobales.Add "GNOMOINTELIGENCIA", 1
212     clavesGlobales.Add "HUMANOAGILIDAD", 1
214     clavesGlobales.Add "HUMANOCARISMA", 1
216     clavesGlobales.Add "HUMANOCONSTITUCION", 1
218     clavesGlobales.Add "HUMANOFUERZA", 1
220     clavesGlobales.Add "HUMANOINTELIGENCIA", 1
222     clavesGlobales.Add "ORCOAGILIDAD", 1
224     clavesGlobales.Add "ORCOCARISMA", 1
226     clavesGlobales.Add "ORCOCONSTITUCION", 1
228     clavesGlobales.Add "ORCOFUERZA", 1
230     clavesGlobales.Add "ORCOINTELIGENCIA", 1
232     clavesGlobales.Add "NUMEROHECHIZO", 1
234     clavesGlobales.Add "NUMLOCALEMSG", 1
244     clavesGlobales.Add "NUMMAPAS", 1
246     clavesGlobales.Add "NUMNPCS", 1
248     clavesGlobales.Add "NUMOBJS", 1
250     clavesGlobales.Add "NUMQUESTS", 1
252     clavesGlobales.Add "NUMSUGERENCIAS", 1
254     clavesGlobales.Add "SUGERENCIA1", 1
256     clavesGlobales.Add "SUGERENCIA2", 1
258     clavesGlobales.Add "SUGERENCIA3", 1
260     clavesGlobales.Add "SUGERENCIA4", 1
262     clavesGlobales.Add "SUGERENCIA5", 1
264     clavesGlobales.Add "SUGERENCIA6", 1

        Dim secciones As Collection
266     Set secciones = origen.GetAllSectionNames()
        Dim clavesPorSeccion As Object
268     Set clavesPorSeccion = CreateObject("Scripting.Dictionary")
        Dim sec As Variant, clave As Variant, claves As Collection
270     For Each sec In secciones
272         Set claves = origen.GetAllKeys(sec)
            Dim clavesNorm As Object
274         Set clavesNorm = CreateObject("Scripting.Dictionary")
276         For Each clave In claves
                Dim baseClave As String
278             baseClave = clave

280             For i = 1 To 5
282                 If LCase(Left(clave, Len(LangPrefix(i)))) = LCase(LangPrefix(i)) Then
284                     baseClave = mid$(clave, Len(LangPrefix(i)) + 1)
                        Exit For

                    End If

286             Next i

288             If Not clavesNorm.Exists(baseClave) Then clavesNorm.Add baseClave, 1
290         Next clave

292         clavesPorSeccion.Add sec, clavesNorm
294     Next sec

296     For Each sec In secciones
298         Set claves = origen.GetAllKeys(sec)
300         procesado = False
302         If EsSeccionTraducible(sec) Or Left$(sec, 4) = "NAME" Or sec = "SUGERENCIAS" Or sec = "INIT" Or sec = "MODRAZA" Then
304             For Each clave In claves
306                 Valor = origen.GetValue(sec, clave)
308                 asignado = False

310                 For i = 1 To 5
312                     If LCase(Left$(clave, Len(LangPrefix(i)))) = LCase(LangPrefix(i)) Then
314                         normalClave = mid$(clave, Len(LangPrefix(i)) + 1)
316                         If LCase(normalClave) = "nombre" Then normalClave = "Nombre"
318                         langWriters(i).ChangeValue sec, FormatoClave(normalClave), Valor
320                         asignado = True
                            Exit For

                        End If

322                 Next i

324                 If Not asignado Then
326                     If clavesGlobales.Exists(UCase(clave)) Or Left$(UCase(clave), 4) = "MAPA" Then

328                         For i = 1 To 5
330                             If langWriters(i).GetValue(sec, FormatoClave(clave)) = "" Then
332                                 langWriters(i).ChangeValue sec, FormatoClave(clave), Valor

                                End If

334                         Next i

                        Else
336                         langWriters(1).ChangeValue sec, FormatoClave(clave), Valor

                        End If

                    End If

338             Next clave

340             Set clavesNorm = clavesPorSeccion(sec)

342             For i = 1 To 5
344                 For Each claveNorm In clavesNorm
346                     If langWriters(i).GetValue(sec, FormatoClave(claveNorm)) = "" Then
348                         langWriters(i).ChangeValue sec, FormatoClave(claveNorm), ""

                        End If
350                 Next claveNorm
362             Next i

364             GoTo SiguienteSeccion

            End If

366         For i = 1 To 5
368             If LCase(Left(sec, Len(LangPrefix(i)))) = LCase(LangPrefix(i)) Then
370                 For Each clave In claves
372                     Valor = origen.GetValue(sec, clave)
374                     langWriters(i).ChangeValue sec, FormatoClave(clave), Valor
376                 Next clave

378                 GoTo SiguienteSeccion

                End If

380             If UCase(sec) = UCase(langMsgSection(i)) Then
382                 For Each clave In claves
384                     Valor = origen.GetValue(sec, clave)
386                     langWriters(i).ChangeValue sec, FormatoClave(clave), Valor
388                 Next clave

390                 GoTo SiguienteSeccion

                End If

392         Next i

394         For Each clave In claves
396             Valor = origen.GetValue(sec, clave)
398             asignado = False

400             For i = 1 To 5
402                 If LCase(Left$(clave, Len(LangPrefix(i)))) = LCase(LangPrefix(i)) Then
404                     normalClave = mid$(clave, Len(LangPrefix(i)) + 1)
406                     If LCase(normalClave) = "nombre" Then normalClave = "Nombre"
408                     langWriters(i).ChangeValue sec, FormatoClave(normalClave), Valor
410                     asignado = True
                        Exit For

                    End If

412             Next i

414             If Not asignado Then
416                 If clavesGlobales.Exists(UCase(clave)) Or Left$(UCase(clave), 4) = "MAPA" Then

418                     For i = 1 To 5
420                         If langWriters(i).GetValue(sec, FormatoClave(clave)) = "" Then
422                             langWriters(i).ChangeValue sec, FormatoClave(clave), Valor

                            End If

424                     Next i

                    Else
426                     langWriters(1).ChangeValue sec, FormatoClave(clave), Valor

                    End If

                End If

428         Next clave

SiguienteSeccion:
430     Next sec

432     For i = 1 To 5
434         langWriters(i).DumpFile App.Path & "\..\Recursos\init\" & LangSuffix(i) & "_localindex.dat"
436         Set langWriters(i) = Nothing
438     Next i

440     MsgBox "Archivos localindex por idioma creados con éxito.", vbInformation

End Sub

Private Function EsSeccionTraducible(ByVal sec As String) As Boolean
        Dim s As String
100     s = LCase(sec)
102     EsSeccionTraducible = (Left(s, 5) = "quest" Or _
           Left(s, 7) = "hechizo" Or _
           Left(s, 3) = "npc" Or _
           Left(s, 3) = "obj")

End Function

Private Function FormatoClave(ByVal clave As String) As String

100     Select Case LCase(clave)

            Case "name": FormatoClave = "Name"
102         Case "texto": FormatoClave = "Texto"
104         Case "desc": FormatoClave = "Desc"
106         Case "grhindex": FormatoClave = "GrhIndex"
108         Case "objtype": FormatoClave = "ObjType"
110         Case Else: FormatoClave = clave

        End Select

End Function


