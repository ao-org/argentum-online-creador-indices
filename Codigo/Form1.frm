VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Creador de indices"
   ClientHeight    =   2850
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Mensajes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear archivo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Preparado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Argentum Online"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
        Dim Obj     As Integer
        Dim Npc     As Integer
        Dim Hechizo As Integer
        Dim Raza    As Integer
        Dim numobjs As Long
' === BLOQUE MULTILENGUAJE ===
' === Escritura multilenguaje ===
Dim idiomas() As String
idiomas = Split("es,en,pt,fr,it", ",")

Dim ManagerES As New clsIniReader
Dim ManagerEN As New clsIniReader
Dim ManagerPT As New clsIniReader
Dim ManagerFR As New clsIniReader
Dim ManagerIT As New clsIniReader

Call ManagerES.Initialize(App.Path & "\..\Recursos\init\es_localindex.dat")
Call ManagerEN.Initialize(App.Path & "\..\Recursos\init\en_localindex.dat")
Call ManagerPT.Initialize(App.Path & "\..\Recursos\init\pt_localindex.dat")
Call ManagerFR.Initialize(App.Path & "\..\Recursos\init\fr_localindex.dat")
Call ManagerIT.Initialize(App.Path & "\..\Recursos\init\it_localindex.dat")

' Guardar claves en los managers según idioma
Private Sub GuardarClaveMultilenguaje(ByVal seccion As String, ByVal clave As String, ByVal valor As String)
    If Len(valor) = 0 Then Exit Sub

    Dim idioma As String
    idioma = LCase(Left(clave, 3))

    Select Case idioma
        Case "en_"
            Call ManagerEN.ChangeValue(seccion, clave, valor)
        Case "pt_"
            Call ManagerPT.ChangeValue(seccion, clave, valor)
        Case "fr_"
            Call ManagerFR.ChangeValue(seccion, clave, valor)
        Case "it_"
            Call ManagerIT.ChangeValue(seccion, clave, valor)
        Case "es_"
            Call ManagerES.ChangeValue(seccion, clave, valor)
        Case Else
            ' Clave sin prefijo de idioma -> va en todos
            Call ManagerES.ChangeValue(seccion, clave, valor)
            Call ManagerEN.ChangeValue(seccion, clave, valor)
            Call ManagerPT.ChangeValue(seccion, clave, valor)
            Call ManagerFR.ChangeValue(seccion, clave, valor)
            Call ManagerIT.ChangeValue(seccion, clave, valor)
    End Select
End Sub

100     If FileExist(OutputFile, vbNormal) Then
102         Clean_File OutputFile

        End If

104     If FileExist(App.Path & "\..\Recursos\Dat\obj.dat", vbNormal) Then
106         ObjFile = App.Path & "\..\Recursos\Dat\obj.dat"
108         numobjs = Val(GetVar(ObjFile, "INIT", "NumOBJs"))
110         Label3.Caption = "0/" & numobjs
112         ReDim ObjData(1 To numobjs) As ObjDatas
            Dim Leer As New clsIniReader
114         Call Leer.Initialize(ObjFile)

116         For Obj = 1 To numobjs
118             DoEvents
                ' Gráficos y nombres (multi idioma)
120             ObjData(Obj).grhindex = Val(Leer.GetValue("OBJ" & Obj, "grhindex"))
122             ObjData(Obj).Name = Leer.GetValue("OBJ" & Obj, "Name")
124             ObjData(Obj).en_name = Leer.GetValue("OBJ" & Obj, "en_Name")
126             ObjData(Obj).pt_name = Leer.GetValue("OBJ" & Obj, "pt_name")
128             ObjData(Obj).fr_name = Leer.GetValue("OBJ" & Obj, "fr_name")
130             ObjData(Obj).it_name = Leer.GetValue("OBJ" & Obj, "it_name")
                ' Texto descriptivo (multi idioma)
132             ObjData(Obj).texto = Leer.GetValue("OBJ" & Obj, "Texto")
134             ObjData(Obj).en_texto = Leer.GetValue("OBJ" & Obj, "en_Texto")
136             ObjData(Obj).pt_texto = Leer.GetValue("OBJ" & Obj, "pt_texto")
138             ObjData(Obj).fr_texto = Leer.GetValue("OBJ" & Obj, "fr_texto")
140             ObjData(Obj).it_texto = Leer.GetValue("OBJ" & Obj, "it_texto")
                ' Info extendida (multi idioma)
142             ObjData(Obj).Info = Leer.GetValue("OBJ" & Obj, "Info")
144             ObjData(Obj).en_Info = Leer.GetValue("OBJ" & Obj, "en_Info")
146             ObjData(Obj).pt_Info = Leer.GetValue("OBJ" & Obj, "pt_Info")
148             ObjData(Obj).fr_Info = Leer.GetValue("OBJ" & Obj, "fr_Info")
150             ObjData(Obj).it_Info = Leer.GetValue("OBJ" & Obj, "it_Info")
                ' Estadísticas
152             ObjData(Obj).MINDEF = Val(Leer.GetValue("OBJ" & Obj, "MinDef"))
154             ObjData(Obj).MaxDEF = Val(Leer.GetValue("OBJ" & Obj, "MaxDef"))
156             ObjData(Obj).MinHit = Val(Leer.GetValue("OBJ" & Obj, "MinHit"))
158             ObjData(Obj).MaxHit = Val(Leer.GetValue("OBJ" & Obj, "MaxHit"))
160             ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
                ' Efectos visuales
162             ObjData(Obj).CreaGRH = Leer.GetValue("OBJ" & Obj, "CreaGRH")
164             ObjData(Obj).CreaLuz = Leer.GetValue("OBJ" & Obj, "CreaLuz")
166             ObjData(Obj).CreaParticulaPiso = Val(Leer.GetValue("OBJ" & Obj, "CreaParticulaPiso"))
                ' Otros efectos
168             ObjData(Obj).Proyectil = Val(Leer.GetValue("OBJ" & Obj, "Proyectil"))
170             ObjData(Obj).Municiones = Val(Leer.GetValue("OBJ" & Obj, "Municiones"))
172             ObjData(Obj).Hechizo = Val(Leer.GetValue("OBJ" & Obj, "Hechizo"))
                ' Componentes alquimia/herbolaria
174             ObjData(Obj).Raices = Val(Leer.GetValue("OBJ" & Obj, "Raices"))
176             ObjData(Obj).Cuchara = Val(Leer.GetValue("OBJ" & Obj, "Cuchara"))
178             ObjData(Obj).Botella = Val(Leer.GetValue("OBJ" & Obj, "Botella"))
180             ObjData(Obj).Mortero = Val(Leer.GetValue("OBJ" & Obj, "Mortero"))
182             ObjData(Obj).FrascoAlq = Val(Leer.GetValue("OBJ" & Obj, "FrascoAlq"))
184             ObjData(Obj).FrascoElixir = Val(Leer.GetValue("OBJ" & Obj, "FrascoElixir"))
186             ObjData(Obj).Dosificador = Val(Leer.GetValue("OBJ" & Obj, "Dosificador"))
                ' Ingredientes mágicos y naturales
188             ObjData(Obj).Orquidea = Val(Leer.GetValue("OBJ" & Obj, "Orquidea"))
190             ObjData(Obj).Carmesi = Val(Leer.GetValue("OBJ" & Obj, "Carmesi"))
192             ObjData(Obj).HongoDeLuz = Val(Leer.GetValue("OBJ" & Obj, "HongoDeLuz"))
194             ObjData(Obj).Esporas = Val(Leer.GetValue("OBJ" & Obj, "Esporas"))
196             ObjData(Obj).Tuna = Val(Leer.GetValue("OBJ" & Obj, "Tuna"))
198             ObjData(Obj).Cala = Val(Leer.GetValue("OBJ" & Obj, "Cala"))
200             ObjData(Obj).ColaDeZorro = Val(Leer.GetValue("OBJ" & Obj, "ColaDeZorro"))
202             ObjData(Obj).FlorOceano = Val(Leer.GetValue("OBJ" & Obj, "FlorOceano"))
204             ObjData(Obj).FlorRoja = Val(Leer.GetValue("OBJ" & Obj, "FlorRoja"))
206             ObjData(Obj).Hierva = Val(Leer.GetValue("OBJ" & Obj, "Hierva"))
208             ObjData(Obj).HojasDeRin = Val(Leer.GetValue("OBJ" & Obj, "HojasDeRin"))
210             ObjData(Obj).HojasRojas = Val(Leer.GetValue("OBJ" & Obj, "HojasRojas"))
212             ObjData(Obj).SemillasPros = Val(Leer.GetValue("OBJ" & Obj, "SemillasPros"))
214             ObjData(Obj).Pimiento = Val(Leer.GetValue("OBJ" & Obj, "Pimiento"))
                ' Materiales
216             ObjData(Obj).Madera = Val(Leer.GetValue("OBJ" & Obj, "Madera"))
218             ObjData(Obj).MaderaElfica = Val(Leer.GetValue("OBJ" & Obj, "MaderaElfica"))
220             ObjData(Obj).PielLobo = Val(Leer.GetValue("OBJ" & Obj, "PielLobo"))
222             ObjData(Obj).PielLoboNegro = Val(Leer.GetValue("OBJ" & Obj, "PielLoboNegro"))
224             ObjData(Obj).PielTigre = Val(Leer.GetValue("OBJ" & Obj, "PielTigre"))
226             ObjData(Obj).PielTigreBengala = Val(Leer.GetValue("OBJ" & Obj, "PielTigreBengala"))
228             ObjData(Obj).PielOsoPardo = Val(Leer.GetValue("OBJ" & Obj, "PielOsoPardo"))
230             ObjData(Obj).PielOsoPolar = Val(Leer.GetValue("OBJ" & Obj, "PielOsoPolar"))
232             ObjData(Obj).LingH = Val(Leer.GetValue("OBJ" & Obj, "LingH"))
234             ObjData(Obj).LingP = Val(Leer.GetValue("OBJ" & Obj, "LingP"))
236             ObjData(Obj).LingO = Val(Leer.GetValue("OBJ" & Obj, "LingO"))
238             ObjData(Obj).Coal = Val(Leer.GetValue("OBJ" & Obj, "Coal"))
                ' Otros
240             ObjData(Obj).Destruye = Val(Leer.GetValue("OBJ" & Obj, "Destruye"))
242             ObjData(Obj).SkHerreria = Val(Leer.GetValue("OBJ" & Obj, "SkHerreria"))
244             ObjData(Obj).SkPociones = Val(Leer.GetValue("OBJ" & Obj, "SkPociones"))
246             ObjData(Obj).Sksastreria = Val(Leer.GetValue("OBJ" & Obj, "Sksastreria"))
248             ObjData(Obj).Valor = Val(Leer.GetValue("OBJ" & Obj, "Valor"))
250             ObjData(Obj).Agarrable = Val(Leer.GetValue("OBJ" & Obj, "Agarrable"))
252             ObjData(Obj).Llave = Val(Leer.GetValue("OBJ" & Obj, "Llave"))
254             ObjData(Obj).Cooldown = Val(Leer.GetValue("OBJ" & Obj, "CD"))
256             ObjData(Obj).CdType = Val(Leer.GetValue("OBJ" & Obj, "CDType"))
258             ObjData(Obj).SpellIndex = Val(Leer.GetValue("OBJ" & Obj, "HechizoIndex"))
260             Label3.ForeColor = vbRed
262             Label3.Caption = "Leyendo objetos: " & Obj & "/" & numobjs
264         Next Obj


266         Obj = 1
            Dim Manager As clsIniReader
268         Set Manager = New clsIniReader
270         Call Manager.Initialize(OutputFile)
272         Call GuardarClaveMultilenguaje("INIT", "NumOBJs", numobjs)

274         For Obj = 1 To numobjs
276             DoEvents
278             Call GuardarClaveMultilenguaje("OBJ" & Obj, "GrhIndex", ObjData(Obj).grhindex)

280             If Len(ObjData(Obj).Name) <> 0 Then
282                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Name", ObjData(Obj).Name)

                End If

284             If Len(ObjData(Obj).texto) <> 0 Then
286                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Texto", ObjData(Obj).texto)

                End If

288             If Len(ObjData(Obj).Info) <> 0 Then
290                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Info", ObjData(Obj).Info)

                End If

                'English
292             If Len(ObjData(Obj).en_name) <> 0 Then
294                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "en_Name", ObjData(Obj).en_name)


                End If

296             If Len(ObjData(Obj).en_texto) <> 0 Then
298                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "en_Texto", ObjData(Obj).en_texto)

                End If

300             If Len(ObjData(Obj).en_Info) <> 0 Then
302                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "en_Info", ObjData(Obj).en_Info)

                End If

304             If ObjData(Obj).MINDEF > 0 Then
306                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "MINDEF", ObjData(Obj).MINDEF)


                End If

308             If ObjData(Obj).MaxDEF > 0 Then
310                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "MaxDEF", ObjData(Obj).MaxDEF)

                End If

312             If ObjData(Obj).MinHit > 0 Then
314                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "MinHIt", ObjData(Obj).MinHit)

                End If

316             If ObjData(Obj).MaxHit > 0 Then
318                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "maxhit", ObjData(Obj).MaxHit)

                End If

320             If ObjData(Obj).ObjType > 0 Then
322                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "ObjType", ObjData(Obj).ObjType)

                End If

324             If Len(ObjData(Obj).CreaLuz) <> 0 Then
326                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "CreaLuz", ObjData(Obj).CreaLuz)

                End If

328             If Len(ObjData(Obj).CreaGRH) <> 0 Then
330                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "CreaGRH", ObjData(Obj).CreaGRH)

                End If

332             If ObjData(Obj).Hechizo <> 0 Then
334                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Hechizo", ObjData(Obj).Hechizo)

                End If

336             If ObjData(Obj).Raices <> 0 Then
338                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Raices", ObjData(Obj).Raices)

                End If

340             If ObjData(Obj).Cuchara <> 0 Then
342                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Cuchara", ObjData(Obj).Cuchara)

                End If

344             If ObjData(Obj).Botella <> 0 Then
346                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Botella", ObjData(Obj).Botella)

                End If

348             If ObjData(Obj).Mortero <> 0 Then
350                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Mortero", ObjData(Obj).Mortero)

                End If

352             If ObjData(Obj).FrascoAlq <> 0 Then
354                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "FrascoAlq", ObjData(Obj).FrascoAlq)

                End If

356             If ObjData(Obj).FrascoElixir <> 0 Then
358                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "FrascoElixir", ObjData(Obj).FrascoElixir)

                End If

360             If ObjData(Obj).Dosificador <> 0 Then
362                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Dosificador", ObjData(Obj).Dosificador)

                End If

364             If ObjData(Obj).Orquidea <> 0 Then
366                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Orquidea", ObjData(Obj).Orquidea)

                End If

368             If ObjData(Obj).Carmesi <> 0 Then
370                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Carmesi", ObjData(Obj).Carmesi)

                End If

372             If ObjData(Obj).HongoDeLuz <> 0 Then
374                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "HongoDeLuz", ObjData(Obj).HongoDeLuz)

                End If

376             If ObjData(Obj).Esporas <> 0 Then
378                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Esporas", ObjData(Obj).Esporas)

                End If

380             If ObjData(Obj).Tuna <> 0 Then
382                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Tuna", ObjData(Obj).Tuna)

                End If

384             If ObjData(Obj).Cala <> 0 Then
386                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Cala", ObjData(Obj).Cala)

                End If

388             If ObjData(Obj).ColaDeZorro <> 0 Then
390                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "ColaDeZorro", ObjData(Obj).ColaDeZorro)

                End If

392             If ObjData(Obj).FlorOceano <> 0 Then
394                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "FlorOceano", ObjData(Obj).FlorOceano)

                End If

396             If ObjData(Obj).FlorRoja <> 0 Then
398                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "FlorRoja", ObjData(Obj).FlorRoja)

                End If

400             If ObjData(Obj).Hierva <> 0 Then
402                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Hierva", ObjData(Obj).Hierva)

                End If

404             If ObjData(Obj).HojasDeRin <> 0 Then
406                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "HojasDeRin", ObjData(Obj).HojasDeRin)

                End If

408             If ObjData(Obj).HojasRojas <> 0 Then
410                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "HojasRojas", ObjData(Obj).HojasRojas)

                End If

412             If ObjData(Obj).SemillasPros <> 0 Then
414                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "SemillasPros", ObjData(Obj).SemillasPros)

                End If

416             If ObjData(Obj).Pimiento <> 0 Then
418                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Pimiento", ObjData(Obj).Pimiento)

                End If

420             If ObjData(Obj).Madera <> 0 Then
422                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Madera", ObjData(Obj).Madera)

                End If

424             If ObjData(Obj).MaderaElfica <> 0 Then
426                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "MaderaElfica", ObjData(Obj).MaderaElfica)

                End If

428             If ObjData(Obj).PielLobo <> 0 Then
430                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "PielLobo", ObjData(Obj).PielLobo)

                End If

432             If ObjData(Obj).PielLoboNegro <> 0 Then
434                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "PielLoboNegro", ObjData(Obj).PielLoboNegro)

                End If

436             If ObjData(Obj).PielOsoPardo <> 0 Then
438                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "PielOsoPardo", ObjData(Obj).PielOsoPardo)

                End If

440             If ObjData(Obj).PielTigre <> 0 Then
442                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "PielTigre", ObjData(Obj).PielTigre)

                End If

444             If ObjData(Obj).PielTigreBengala <> 0 Then
446                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "PielTigreBengala", ObjData(Obj).PielTigreBengala)

                End If

448             If ObjData(Obj).PielOsoPolar <> 0 Then
450                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "PielOsoPolar", ObjData(Obj).PielOsoPolar)

                End If

452             If ObjData(Obj).LingH <> 0 Then
454                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "LingH", ObjData(Obj).LingH)

                End If

456             If ObjData(Obj).LingP <> 0 Then
458                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "LingP", ObjData(Obj).LingP)

                End If

460             If ObjData(Obj).LingO <> 0 Then
462                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "LingO", ObjData(Obj).LingO)

                End If

464             If ObjData(Obj).Coal <> 0 Then
466                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Coal", ObjData(Obj).Coal)

                End If

468             If ObjData(Obj).Destruye <> 0 Then
470                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Destruye", ObjData(Obj).Destruye)

                End If

472             If ObjData(Obj).SkHerreria <> 0 Then
474                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "SkHerreria", ObjData(Obj).SkHerreria)

                End If

476             If ObjData(Obj).SkPociones <> 0 Then
478                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "SkPociones", ObjData(Obj).SkPociones)

                End If

480             If ObjData(Obj).Sksastreria <> 0 Then
482                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Sksastreria", ObjData(Obj).Sksastreria)

                End If

484             If ObjData(Obj).Valor <> 0 Then
486                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Valor", ObjData(Obj).Valor)

                End If

488             If ObjData(Obj).Agarrable Then
490                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Agarrable", 1)

                End If

492             If ObjData(Obj).CreaParticulaPiso > 0 Then
494                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "CreaParticulaPiso", ObjData(Obj).CreaParticulaPiso)

                End If

496             If ObjData(Obj).Proyectil > 0 Then
498                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Proyectil", ObjData(Obj).Proyectil)

                End If

500             If ObjData(Obj).Municiones > 0 Then
502                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Municiones", ObjData(Obj).Municiones)

                End If

504             If ObjData(Obj).Llave > 0 Then
506                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "Llave", ObjData(Obj).Llave)

                End If

508             If ObjData(Obj).Cooldown > 0 Then
510                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "CD", ObjData(Obj).Cooldown)

                End If

512             If ObjData(Obj).CdType > 0 Then
514                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "CDType", ObjData(Obj).CdType)

                End If

516             If ObjData(Obj).SpellIndex > 0 Then
518                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "SpellIndex", ObjData(Obj).SpellIndex)

                End If

520             Label3.Caption = "Grabando: " & Obj & "/" & numobjs
522             Label3.ForeColor = &HC0C0&

                ' Portugués
524             If Len(ObjData(Obj).pt_name) <> 0 Then
526                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "pt_name", ObjData(Obj).pt_name)

                End If

528             If Len(ObjData(Obj).pt_texto) <> 0 Then
530                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "pt_texto", ObjData(Obj).pt_texto)

                End If

                ' Francés
532             If Len(ObjData(Obj).fr_name) <> 0 Then
534                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "fr_name", ObjData(Obj).fr_name)

                End If

536             If Len(ObjData(Obj).fr_texto) <> 0 Then
538                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "fr_texto", ObjData(Obj).fr_texto)

                End If

                ' Italiano
540             If Len(ObjData(Obj).it_name) <> 0 Then
542                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "it_name", ObjData(Obj).it_name)

                End If

544             If Len(ObjData(Obj).it_texto) <> 0 Then
546                 Call GuardarClaveMultilenguaje("OBJ" & Obj, "it_texto", ObjData(Obj).it_texto)

                End If

548         Next Obj

550         Label3.ForeColor = vbGreen
552         Label3.Caption = "Creado objindex.dat"
        Else
554         MsgBox "Falta el archivo obj.dat dentro de la carpeta INIT."

        End If

556     If FileExist(App.Path & "\..\Recursos\Dat\npcs.dat", vbNormal) Then
558         NpcFile = App.Path & "\..\Recursos\Dat\npcs.dat"
560         Call Leer.Initialize(NpcFile)
            Dim numnpcs As Long
562         numnpcs = Val(GetVar(NpcFile, "INIT", "NumNPCs"))
564         Label3.Caption = "0/" & numnpcs
566         ReDim NpcData(1 To numnpcs) As NpcDatas
            Dim aux As String

568         For Npc = 1 To numnpcs
570             DoEvents
                ' Nombres y descripciones (multi idioma)
572             NpcData(Npc).Name = Leer.GetValue("npc" & Npc, "Name")
574             NpcData(Npc).en_name = Leer.GetValue("npc" & Npc, "en_Name")
576             NpcData(Npc).pt_name = Leer.GetValue("npc" & Npc, "pt_Name")
578             NpcData(Npc).fr_name = Leer.GetValue("npc" & Npc, "fr_Name")
580             NpcData(Npc).it_name = Leer.GetValue("npc" & Npc, "it_Name")
582             NpcData(Npc).Desc = Leer.GetValue("npc" & Npc, "desc")
584             NpcData(Npc).en_Desc = Leer.GetValue("npc" & Npc, "en_desc")
586             NpcData(Npc).pt_Desc = Leer.GetValue("npc" & Npc, "pt_desc")
588             NpcData(Npc).fr_Desc = Leer.GetValue("npc" & Npc, "fr_desc")
590             NpcData(Npc).it_Desc = Leer.GetValue("npc" & Npc, "it_desc")
                ' Atributos generales
592             NpcData(Npc).Body = Val(Leer.GetValue("npc" & Npc, "Body"))
594             NpcData(Npc).Head = Val(Leer.GetValue("npc" & Npc, "Head"))
596             NpcData(Npc).Hp = Val(Leer.GetValue("npc" & Npc, "MaxHP"))
598             NpcData(Npc).Exp = Val(Leer.GetValue("npc" & Npc, "GiveEXP"))
600             NpcData(Npc).ExpClan = Val(Leer.GetValue("npc" & Npc, "GiveEXPClan"))
602             NpcData(Npc).Oro = Val(Leer.GetValue("npc" & Npc, "GiveGLD"))
                ' Combate
604             NpcData(Npc).MinHit = Val(Leer.GetValue("npc" & Npc, "MinHit"))
606             NpcData(Npc).MaxHit = Val(Leer.GetValue("npc" & Npc, "MaxHit"))
                ' Flags y control
608             NpcData(Npc).PuedeInvocar = Val(Leer.GetValue("npc" & Npc, "PuedeInvocar"))
610             NpcData(Npc).NoMapInfo = Val(Leer.GetValue("npc" & Npc, "NoMapInfo"))
612             NpcData(Npc).QuizaProb = Val(Leer.GetValue("npc" & Npc, "QuizaProb"))
                ' QuizaDropea
614             aux = Val(GetVar(NpcFile, "Npc" & Npc, "NumQuiza"))


616             If aux = 0 Then
618                 NpcData(Npc).NumQuiza = 0
                Else
620                 NpcData(Npc).NumQuiza = Val(aux)
622                 ReDim NpcData(Npc).QuizaDropea(1 To NpcData(Npc).NumQuiza) As Integer
                    Dim LoopC As Integer
624                 For LoopC = 1 To NpcData(Npc).NumQuiza
626                     NpcData(Npc).QuizaDropea(LoopC) = Val(Leer.GetValue("npc" & Npc, "QuizaDropea" & LoopC))
628                 Next LoopC

                End If

630             Label3.ForeColor = vbRed
632             Label3.Caption = "Leyendo NPCs: " & Npc & "/" & numnpcs
634         Next Npc

636         Npc = 1
638         Call GuardarClaveMultilenguaje("INIT", "NumNPCs", numnpcs)

640         For Npc = 1 To numnpcs
642             DoEvents

644             If Len(NpcData(Npc).Name) <> 0 Then
646                 Call GuardarClaveMultilenguaje("Npc" & Npc, "Name", NpcData(Npc).Name)

                End If

648             If Len(NpcData(Npc).en_name) <> 0 Then
650                 Call GuardarClaveMultilenguaje("Npc" & Npc, "en_Name", NpcData(Npc).en_name)

                    ' Name multilenguaje
652                 If Len(NpcData(Npc).pt_name) <> 0 Then
654                     Call GuardarClaveMultilenguaje("Npc" & Npc, "pt_Name", NpcData(Npc).pt_name)

                    End If

656                 If Len(NpcData(Npc).fr_name) <> 0 Then
658                     Call GuardarClaveMultilenguaje("Npc" & Npc, "fr_Name", NpcData(Npc).fr_name)

                    End If

660                 If Len(NpcData(Npc).it_name) <> 0 Then
662                     Call GuardarClaveMultilenguaje("Npc" & Npc, "it_Name", NpcData(Npc).it_name)

                    End If

                End If

664             If Len(NpcData(Npc).en_Desc) <> 0 Then
666                 Call GuardarClaveMultilenguaje("Npc" & Npc, "en_desc", NpcData(Npc).en_Desc)

                    ' Desc multilenguaje
668                 If Len(NpcData(Npc).pt_Desc) <> 0 Then
670                     Call GuardarClaveMultilenguaje("Npc" & Npc, "pt_desc", NpcData(Npc).pt_Desc)

                    End If

672                 If Len(NpcData(Npc).fr_Desc) <> 0 Then
674                     Call GuardarClaveMultilenguaje("Npc" & Npc, "fr_desc", NpcData(Npc).fr_Desc)

                    End If

676                 If Len(NpcData(Npc).it_Desc) <> 0 Then
678                     Call GuardarClaveMultilenguaje("Npc" & Npc, "it_desc", NpcData(Npc).it_Desc)

                    End If


                End If

680             If Len(NpcData(Npc).Desc) <> 0 Then
682                 Call GuardarClaveMultilenguaje("Npc" & Npc, "Desc", NpcData(Npc).Desc)

                End If

684             If NpcData(Npc).Body <> 0 Then
686                 Call GuardarClaveMultilenguaje("Npc" & Npc, "Body", NpcData(Npc).Body)

                End If

688             If NpcData(Npc).Head <> 0 Then
690                 Call GuardarClaveMultilenguaje("Npc" & Npc, "Head", NpcData(Npc).Head)

                End If

692             If NpcData(Npc).Exp <> 0 Then
694                 Call GuardarClaveMultilenguaje("Npc" & Npc, "Exp", NpcData(Npc).Exp)

                End If

696             If NpcData(Npc).Hp <> 0 Then
698                 Call GuardarClaveMultilenguaje("Npc" & Npc, "Hp", NpcData(Npc).Hp)

                End If

700             If NpcData(Npc).MaxHit <> 0 Then
702                 Call GuardarClaveMultilenguaje("Npc" & Npc, "MaxHit", NpcData(Npc).MaxHit)

                End If

704             If NpcData(Npc).MinHit <> 0 Then
706                 Call GuardarClaveMultilenguaje("Npc" & Npc, "MinHit", NpcData(Npc).MinHit)

                End If

708             If NpcData(Npc).Oro <> 0 Then
710                 Call GuardarClaveMultilenguaje("Npc" & Npc, "Oro", NpcData(Npc).Oro)

                End If

712             If NpcData(Npc).ExpClan <> 0 Then
714                 Call GuardarClaveMultilenguaje("Npc" & Npc, "GiveEXPClan", NpcData(Npc).ExpClan)

                End If

716             If NpcData(Npc).NumQuiza <> 0 Then
718                 Call GuardarClaveMultilenguaje("Npc" & Npc, "NumQuiza", NpcData(Npc).NumQuiza)

720                 For LoopC = 1 To NpcData(Npc).NumQuiza
722                     Call Manager.ChangeValue("Npc" & Npc, "QuizaDropea" & LoopC, NpcData(Npc).QuizaDropea(LoopC))
724                 Next LoopC

                End If

726             If NpcData(Npc).QuizaProb <> 0 Then
728                 Call GuardarClaveMultilenguaje("Npc" & Npc, "QuizaProb", NpcData(Npc).QuizaProb)

                End If

730             If NpcData(Npc).NoMapInfo <> 0 Then
732                 Call GuardarClaveMultilenguaje("Npc" & Npc, "NoMapInfo", NpcData(Npc).NoMapInfo)

                End If

734             If NpcData(Npc).PuedeInvocar <> 0 Then
736                 Call GuardarClaveMultilenguaje("Npc" & Npc, "PuedeInvocar", NpcData(Npc).PuedeInvocar)

                End If

738             Label3.Caption = "Grabando NPCs: " & Npc & "/" & numnpcs
740             Label3.ForeColor = &HC0C0&
742         Next Npc

        Else
744         MsgBox "Falta el archivo npcs.dat dentro de la carpeta dats."

        End If

746     If FileExist(App.Path & "\..\Recursos\Dat\hechizos.dat", vbNormal) Then
            Dim hechizosFile As String, numhechizos As Long
748         hechizosFile = App.Path & "\..\Recursos\Dat\hechizos.dat"
750         numhechizos = Val(GetVar(hechizosFile, "INIT", "NumeroHechizos"))
            Dim hechic As New clsIniReader
752         Call hechic.Initialize(hechizosFile)
754         Label3.Caption = "Leyendo Hechizos: " & "0/" & numhechizos
756         ReDim HechizoData(1 To numhechizos) As HechizoDatas

758         For Hechizo = 1 To numhechizos
760             DoEvents
                ' Nombre del hechizo
762             HechizoData(Hechizo).Nombre = hechic.GetValue("Hechizo" & Hechizo, "Nombre")
764             HechizoData(Hechizo).en_name = hechic.GetValue("Hechizo" & Hechizo, "en_name")
766             HechizoData(Hechizo).pt_name = hechic.GetValue("Hechizo" & Hechizo, "pt_Nombre")
768             HechizoData(Hechizo).fr_name = hechic.GetValue("Hechizo" & Hechizo, "fr_Nombre")
770             HechizoData(Hechizo).it_name = hechic.GetValue("Hechizo" & Hechizo, "it_Nombre")
                ' Descripción
772             HechizoData(Hechizo).Desc = hechic.GetValue("Hechizo" & Hechizo, "desc")
774             HechizoData(Hechizo).en_Desc = hechic.GetValue("Hechizo" & Hechizo, "en_Desc")
776             HechizoData(Hechizo).pt_Desc = hechic.GetValue("Hechizo" & Hechizo, "pt_Desc")
778             HechizoData(Hechizo).fr_Desc = hechic.GetValue("Hechizo" & Hechizo, "fr_Desc")
780             HechizoData(Hechizo).it_Desc = hechic.GetValue("Hechizo" & Hechizo, "it_Desc")
                ' Palabras mágicas
782             HechizoData(Hechizo).PalabrasMagicas = hechic.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
                ' Mensajes del lanzador
784             HechizoData(Hechizo).HechizeroMsg = hechic.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
786             HechizoData(Hechizo).en_HechizeroMsg = hechic.GetValue("Hechizo" & Hechizo, "en_HechizeroMsg")
788             HechizoData(Hechizo).pt_HechizeroMsg = hechic.GetValue("Hechizo" & Hechizo, "pt_HechizeroMsg")
790             HechizoData(Hechizo).fr_HechizeroMsg = hechic.GetValue("Hechizo" & Hechizo, "fr_HechizeroMsg")
792             HechizoData(Hechizo).it_HechizeroMsg = hechic.GetValue("Hechizo" & Hechizo, "it_HechizeroMsg")
                ' Mensajes para target
794             HechizoData(Hechizo).TargetMsg = hechic.GetValue("Hechizo" & Hechizo, "TargetMsg")
796             HechizoData(Hechizo).en_TargetMsg = hechic.GetValue("Hechizo" & Hechizo, "en_TargetMsg")
798             HechizoData(Hechizo).pt_TargetMsg = hechic.GetValue("Hechizo" & Hechizo, "pt_TargetMsg")
800             HechizoData(Hechizo).fr_TargetMsg = hechic.GetValue("Hechizo" & Hechizo, "fr_TargetMsg")
802             HechizoData(Hechizo).it_TargetMsg = hechic.GetValue("Hechizo" & Hechizo, "it_TargetMsg")
                ' Mensajes para caster
804             HechizoData(Hechizo).PropioMsg = hechic.GetValue("Hechizo" & Hechizo, "PropioMsg")
806             HechizoData(Hechizo).en_PropioMsg = hechic.GetValue("Hechizo" & Hechizo, "en_PropioMsg")
808             HechizoData(Hechizo).pt_PropioMsg = hechic.GetValue("Hechizo" & Hechizo, "pt_PropioMsg")
810             HechizoData(Hechizo).fr_PropioMsg = hechic.GetValue("Hechizo" & Hechizo, "fr_PropioMsg")
812             HechizoData(Hechizo).it_PropioMsg = hechic.GetValue("Hechizo" & Hechizo, "it_PropioMsg")
                ' Requerimientos
814             HechizoData(Hechizo).ManaRequerido = Val(hechic.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
816             HechizoData(Hechizo).StaRequerido = Val(hechic.GetValue("Hechizo" & Hechizo, "StaRequerido"))
818             HechizoData(Hechizo).MinSkill = Val(hechic.GetValue("Hechizo" & Hechizo, "MinSkill"))
                ' Gráfica
820             HechizoData(Hechizo).IconoIndex = Val(hechic.GetValue("Hechizo" & Hechizo, "IconoIndex"))
822             HechizoData(Hechizo).Cooldown = Val(hechic.GetValue("Hechizo" & Hechizo, "Cooldown"))
824             Label3.ForeColor = vbRed
826             Label3.Caption = "Leyendo: " & Hechizo & "/" & numhechizos
828         Next Hechizo

830         Call GuardarClaveMultilenguaje("INIT", "NumeroHechizo", numhechizos)

832         For Hechizo = 1 To numhechizos
834             DoEvents
                ' Español
836             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "Nombre", HechizoData(Hechizo).Nombre)
838             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "Desc", HechizoData(Hechizo).Desc)
840             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "PalabrasMagicas", HechizoData(Hechizo).PalabrasMagicas)
842             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "HechizeroMsg", HechizoData(Hechizo).HechizeroMsg)
844             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "TargetMsg", HechizoData(Hechizo).TargetMsg)
846             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "PropioMsg", HechizoData(Hechizo).PropioMsg)
                ' Inglés
848             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "en_Name", HechizoData(Hechizo).en_name)
850             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "en_Desc", HechizoData(Hechizo).en_Desc)
852             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "en_HechizeroMsg", HechizoData(Hechizo).en_HechizeroMsg)
854             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "en_TargetMsg", HechizoData(Hechizo).en_TargetMsg)
856             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "en_PropioMsg", HechizoData(Hechizo).en_PropioMsg)
                ' Portugués
858             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "pt_Nombre", HechizoData(Hechizo).pt_name)
860             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "pt_Desc", HechizoData(Hechizo).pt_Desc)
862             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "pt_HechizeroMsg", HechizoData(Hechizo).pt_HechizeroMsg)
864             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "pt_TargetMsg", HechizoData(Hechizo).pt_TargetMsg)
866             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "pt_PropioMsg", HechizoData(Hechizo).pt_PropioMsg)
                ' Francés
868             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "fr_Nombre", HechizoData(Hechizo).fr_name)
870             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "fr_Desc", HechizoData(Hechizo).fr_Desc)
872             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "fr_HechizeroMsg", HechizoData(Hechizo).fr_HechizeroMsg)
874             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "fr_TargetMsg", HechizoData(Hechizo).fr_TargetMsg)
876             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "fr_PropioMsg", HechizoData(Hechizo).fr_PropioMsg)
                ' Italiano
878             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "it_Nombre", HechizoData(Hechizo).it_name)
880             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "it_Desc", HechizoData(Hechizo).it_Desc)
882             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "it_HechizeroMsg", HechizoData(Hechizo).it_HechizeroMsg)
884             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "it_TargetMsg", HechizoData(Hechizo).it_TargetMsg)
886             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "it_PropioMsg", HechizoData(Hechizo).it_PropioMsg)
                ' Otros datos
888             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "ManaRequerido", HechizoData(Hechizo).ManaRequerido)
890             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "StaRequerido", HechizoData(Hechizo).StaRequerido)
892             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "MinSkill", HechizoData(Hechizo).MinSkill)
894             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "IconoIndex", HechizoData(Hechizo).IconoIndex)
896             Call GuardarClaveMultilenguaje("Hechizo" & Hechizo, "Cooldown", HechizoData(Hechizo).Cooldown)
898             Label3.Caption = "Grabando Hechizos: " & Hechizo & "/" & numhechizos
900             Label3.ForeColor = &HC0C0&
902         Next Hechizo

        End If

904     Call CargarMensajesIdioma("SP", Manager, Label3)
906     Call CargarMensajesIdioma("EN", Manager, Label3)
908     Call CargarMensajesIdioma("PT", Manager, Label3)
910     Call CargarMensajesIdioma("FR", Manager, Label3)
912     Call CargarMensajesIdioma("IT", Manager, Label3)
914     'Call CargarMensajesIdioma("ES", Manager, Label3)

            Dim MsgFile As String
            Dim NumLocaleSP_Msg As Long
            Dim arrLocale_SP_SMG() As String
            Dim SP_MSG As Integer
            Dim EN_MSG As Integer
            Dim Msgsss As New clsIniReader
            
        If FileExist(App.Path & "\..\Recursos\init\EN_LocalMsg.dat", vbNormal) Then
            Dim NumLocaleEN_Msg As Long
            Dim arrLocale_EN_SMG() As String
        
            MsgFile = App.Path & "\..\Recursos\init\EN_LocalMsg.dat"
            Call Msgsss.Initialize(MsgFile)
            NumLocaleEN_Msg = Val(Msgsss.GetValue("INIT", "NumLocaleEN_Msg"))
            Label3.Caption = "0/" & CStr(NumLocaleEN_Msg)
            ReDim arrLocale_EN_SMG(1 To NumLocaleEN_Msg) As String

            Call GuardarClaveMultilenguaje("INIT", "NumLocaleMsg", NumLocaleEN_Msg)

        Else
            MsgBox "Falta el archivo EN_LocalMsg.dat dentro de la carpeta dats."
        End If


916     If FileExist(App.Path & "\..\Recursos\init\NameMapa.dat", vbNormal) Then
            Dim MapFile As String
918         MapFile = App.Path & "\..\Recursos\init\NameMapa.dat"
            Dim Mapa As New clsIniReader
920         Call Mapa.Initialize(MapFile)
922         Label3.Caption = "0/" & 750
924         ReDim MapName(1 To 750) As String
926         ReDim MapDesc(1 To 750) As String

928         For Npc = 1 To 750
930             DoEvents
932             MapName(Npc) = Mapa.GetValue("NameMapa", "mapa" & Npc)
934             MapDesc(Npc) = Mapa.GetValue("NameMapa", "mapa" & Npc & "desc")
936             Label3.ForeColor = vbRed
938             Label3.Caption = "Leyendo Mapas: " & Npc & "/" & 750
940         Next Npc

942         Npc = 1
944         Call GuardarClaveMultilenguaje("INIT", "NumMapas", 750)

946         For Npc = 1 To 750
948             DoEvents
950             Call Manager.ChangeValue("NAMEMAPA", "Mapa" & Npc, MapName(Npc))
952             Call Manager.ChangeValue("NAMEMAPA", "Mapa" & Npc & "Desc", MapDesc(Npc))
954             Label3.Caption = "Grabando Mapas: " & Npc & "/" & 750
956             Label3.ForeColor = &HC0C0&
958         Next Npc

        Else
960         MsgBox "Falta el archivo NameMapa.dat dentro de la carpeta dats."

        End If

    ' Quests - Soporte multilenguaje
    If FileExist(App.Path & "\..\Recursos\Dat\Quests.DAT", vbNormal) Then
        MapFile = App.Path & "\..\Recursos\Dat\Quests.DAT"
        Call Mapa.Initialize(MapFile)

        Dim nunquest As Integer
        nunquest = Mapa.GetValue("INIT", "NumQuests")
        Label3.Caption = "0/" & nunquest

        ReDim QuestName(1 To nunquest) As String
        ReDim QuestDesc(1 To nunquest) As String
        ReDim QuestFin(1 To nunquest) As String
        ReDim QuestNameEN(1 To nunquest) As String
        ReDim QuestDescEN(1 To nunquest) As String
        ReDim QuestFinEN(1 To nunquest) As String
        ReDim QuestNamePT(1 To nunquest) As String
        ReDim QuestDescPT(1 To nunquest) As String
        ReDim QuestFinPT(1 To nunquest) As String
        ReDim QuestNameFR(1 To nunquest) As String
        ReDim QuestDescFR(1 To nunquest) As String
        ReDim QuestFinFR(1 To nunquest) As String
        ReDim QuestNameIT(1 To nunquest) As String
        ReDim QuestDescIT(1 To nunquest) As String
        ReDim QuestFinIT(1 To nunquest) As String
        ReDim QuestNext(1 To nunquest) As String
        ReDim QuestPos(1 To nunquest) As Integer
        ReDim QuestRepetible(1 To nunquest) As Byte
        ReDim RequiredLevel(1 To nunquest) As Integer

        For Npc = 1 To nunquest
            DoEvents
            Label3.ForeColor = vbRed
            Label3.Caption = "Leyendo Quest: " & Npc & "/" & nunquest

            QuestName(Npc) = Mapa.GetValue("QUEST" & Npc, "Nombre")
            QuestNameEN(Npc) = Mapa.GetValue("QUEST" & Npc, "en_Nombre")
            QuestNamePT(Npc) = Mapa.GetValue("QUEST" & Npc, "pt_Nombre")
            QuestNameFR(Npc) = Mapa.GetValue("QUEST" & Npc, "fr_Nombre")
            QuestNameIT(Npc) = Mapa.GetValue("QUEST" & Npc, "it_Nombre")

            QuestDesc(Npc) = Mapa.GetValue("QUEST" & Npc, "Desc")
            QuestFin(Npc) = Mapa.GetValue("QUEST" & Npc, "DescFinal")
            QuestDescEN(Npc) = Mapa.GetValue("QUEST" & Npc, "en_Desc")
            QuestFinEN(Npc) = Mapa.GetValue("QUEST" & Npc, "en_DescFinal")
            QuestDescPT(Npc) = Mapa.GetValue("QUEST" & Npc, "pt_Desc")
            QuestFinPT(Npc) = Mapa.GetValue("QUEST" & Npc, "pt_DescFinal")
            QuestDescFR(Npc) = Mapa.GetValue("QUEST" & Npc, "fr_Desc")
            QuestFinFR(Npc) = Mapa.GetValue("QUEST" & Npc, "fr_DescFinal")
            QuestDescIT(Npc) = Mapa.GetValue("QUEST" & Npc, "it_Desc")
            QuestFinIT(Npc) = Mapa.GetValue("QUEST" & Npc, "it_DescFinal")

            QuestNext(Npc) = Mapa.GetValue("QUEST" & Npc, "NextQuest")
            QuestRepetible(Npc) = Val(Mapa.GetValue("QUEST" & Npc, "Repetible"))
            QuestPos(Npc) = Val(Mapa.GetValue("QUEST" & Npc, "PosMap"))
            RequiredLevel(Npc) = Val(Mapa.GetValue("QUEST" & Npc, "RequiredLevel"))
        Next Npc

        Npc = 1
        Call GuardarClaveMultilenguaje("INIT", "NumQuests", nunquest)

        For Npc = 1 To nunquest
            DoEvents
            Label3.ForeColor = &HC0C0&
            Label3.Caption = "Grabando Quest: " & Npc & "/" & nunquest

            Call GuardarClaveMultilenguaje("QUEST" & Npc, "Nombre", QuestName(Npc))
            Call GuardarClaveMultilenguaje("QUEST" & Npc, "en_Nombre", QuestNameEN(Npc))
            Call GuardarClaveMultilenguaje("QUEST" & Npc, "pt_Nombre", QuestNamePT(Npc))
            Call GuardarClaveMultilenguaje("QUEST" & Npc, "fr_Nombre", QuestNameFR(Npc))
            Call GuardarClaveMultilenguaje("QUEST" & Npc, "it_Nombre", QuestNameIT(Npc))

            Call GuardarClaveMultilenguaje("QUEST" & Npc, "Desc", QuestDesc(Npc))
            Call GuardarClaveMultilenguaje("QUEST" & Npc, "DescFinal", QuestFin(Npc))
            Call GuardarClaveMultilenguaje("QUEST" & Npc, "en_Desc", QuestDescEN(Npc))
            Call GuardarClaveMultilenguaje("QUEST" & Npc, "en_DescFinal", QuestFinEN(Npc))
            Call GuardarClaveMultilenguaje("QUEST" & Npc, "pt_Desc", QuestDescPT(Npc))
            Call GuardarClaveMultilenguaje("QUEST" & Npc, "pt_DescFinal", QuestFinPT(Npc))
            Call GuardarClaveMultilenguaje("QUEST" & Npc, "fr_Desc", QuestDescFR(Npc))
            Call GuardarClaveMultilenguaje("QUEST" & Npc, "fr_DescFinal", QuestFinFR(Npc))
            Call GuardarClaveMultilenguaje("QUEST" & Npc, "it_Desc", QuestDescIT(Npc))
            Call GuardarClaveMultilenguaje("QUEST" & Npc, "it_DescFinal", QuestFinIT(Npc))

            Call GuardarClaveMultilenguaje("QUEST" & Npc, "NextQuest", QuestNext(Npc))
            Call GuardarClaveMultilenguaje("QUEST" & Npc, "Repetible", QuestRepetible(Npc))
            Call GuardarClaveMultilenguaje("QUEST" & Npc, "RequiredLevel", RequiredLevel(Npc))
            Call GuardarClaveMultilenguaje("QUEST" & Npc, "PosMap", QuestPos(Npc))
        Next Npc

    Else
        MsgBox "Falta el archivo Quests.DAT dentro de la carpeta dats."
    End If


1058     If FileExist(App.Path & "\..\Recursos\init\sugerencias.ini", vbNormal) Then
1060         MapFile = App.Path & "\..\Recursos\init\sugerencias.ini"
1062         Call Mapa.Initialize(MapFile)
             Dim NumSug As Integer
1064         NumSug = Val(Mapa.GetValue("Sugerencias", "NumSugerencias"))
1066         Label3.Caption = "0/" & CStr(NumSug)
1068         ReDim Sugerencia(1 To NumSug) As String

1070         For Npc = 1 To NumSug
1072             DoEvents
1074             Sugerencia(Npc) = Mapa.GetValue("Sugerencias", "Sugerencia" & Npc)
1076             Label3.ForeColor = vbRed
1078             Label3.Caption = "Leyendo: " & Npc & "/" & nunquest
1080         Next Npc

1082         Npc = 1
1084         Call GuardarClaveMultilenguaje("INIT", "NumSugerencias", NumSug)

1086         For Npc = 1 To NumSug
1088             DoEvents
1090             Call Manager.ChangeValue("Sugerencias", "Sugerencia" & Npc, Sugerencia(Npc))
1092             Label3.Caption = "Grabando: " & Npc & "/" & NumSug
1094             Label3.ForeColor = &HC0C0&
1096         Next Npc

         Else
1098         MsgBox "Falta el archivo Sugerencias.ini dentro de la carpeta init."

         End If

         Dim ListaRazas(1 To NUMRAZAS) As String
1100     ListaRazas(1) = "Humano"
1102     ListaRazas(2) = "Elfo"
1104     ListaRazas(3) = "Elfo Oscuro"
1106     ListaRazas(4) = "Gnomo"
1108     ListaRazas(5) = "Enano"
1110     ListaRazas(6) = "Orco"
1112     Call Leer.Initialize(App.Path & "\..\Recursos\Dat\Balance.dat")
         Dim SearchVar As String

1114     For Raza = 1 To NUMRAZAS

1116         With ModRaza(Raza)
1118             SearchVar = Replace(ListaRazas(Raza), " ", vbNullString)
1120             .Fuerza = Val(Leer.GetValue("MODRAZA", SearchVar + "Fuerza"))
1122             .Agilidad = Val(Leer.GetValue("MODRAZA", SearchVar + "Agilidad"))
1124             .Inteligencia = Val(Leer.GetValue("MODRAZA", SearchVar + "Inteligencia"))
1126             .Constitucion = Val(Leer.GetValue("MODRAZA", SearchVar + "Constitucion"))
1128             .Carisma = Val(Leer.GetValue("MODRAZA", SearchVar + "Carisma"))
1130             Call Manager.ChangeValue("MODRAZA", SearchVar + "Fuerza", .Fuerza)
1132             Call Manager.ChangeValue("MODRAZA", SearchVar + "Agilidad", .Agilidad)
1134             Call Manager.ChangeValue("MODRAZA", SearchVar + "Inteligencia", .Inteligencia)
1136             Call Manager.ChangeValue("MODRAZA", SearchVar + "Constitucion", .Constitucion)
1138             Call Manager.ChangeValue("MODRAZA", SearchVar + "Carisma", .Carisma)

             End With

1140     Next Raza

1142     Set Leer = Nothing
1144     Call Manager.DumpFile(OutputFile)
1146     Set Manager = Nothing
1148     Label3.ForeColor = vbGreen
1150     Label3.Caption = "Creado localindex.dat"

End Sub

Private Sub Command2_Click()
100     Form2.Show

End Sub

Public Sub LeerLineaComandos()
        Dim rdata As String
100     rdata = Command
        Dim FileTypeName As String
102     FileTypeName = ReadField(1, rdata, Asc("*")) ' File Type Name

104     If Len(FileTypeName) > 0 Then
106         FileTypeName = UCase(FileTypeName)

108         Select Case FileTypeName

                Case Is = "CREAR_ARCHIVO"
110                 Call Command1_Click

            End Select

112         End

        End If

End Sub

Private Sub Form_Load()
100     Form1.Visible = True
102     OutputFile = App.Path & "\..\Recursos\init\localindex.dat"
        ' Leer argumentos
104     Call LeerLineaComandos

End Sub

Private Sub CargarMensajesIdioma(ByVal codigoIdioma As String, _
                                 ByRef Manager As Object, _
                                 ByRef Label3 As Object)
        Dim MsgFile   As String
        Dim Msgsss    As New clsIniReader
        Dim NumMsgs   As Long
        Dim arrMsgs() As String
        Dim i         As Long
        Dim keyInit   As String, keySeccion As String
100     MsgFile = App.Path & "\..\Recursos\init\" & codigoIdioma & "_LocalMsg.dat"

102     If FileExist(MsgFile, vbNormal) Then
104         Msgsss.Initialize MsgFile
106         keyInit = "NumLocale" & codigoIdioma & "_Msg"
108         keySeccion = codigoIdioma & "_MSG"
110         NumMsgs = Val(Msgsss.GetValue("INIT", keyInit))
112         ReDim arrMsgs(1 To NumMsgs)

            ' Leer mensajes
114         For i = 1 To NumMsgs
116             DoEvents
118             arrMsgs(i) = Msgsss.GetValue(keySeccion, "Msg" & i)
120             Label3.ForeColor = vbRed
122             Label3.Caption = "Leyendo MSG " & codigoIdioma & ": " & i & "/" & NumMsgs
124         Next i

            ' Escribir mensajes
126         For i = 1 To NumMsgs
128             DoEvents
130             Manager.ChangeValue codigoIdioma & "_MSG", "Msg" & i, arrMsgs(i)
132             Label3.Caption = "Grabando MSG " & codigoIdioma & ": " & i & "/" & NumMsgs
134             Label3.ForeColor = &HC0C0&
136         Next i

        Else
138         MsgBox "Falta el archivo " & codigoIdioma & "_LocalMsg.dat dentro de la carpeta init.", vbExclamation

        End If

End Sub
