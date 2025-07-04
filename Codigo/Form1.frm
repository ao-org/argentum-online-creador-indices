VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Creador de indices Multilenguaje"
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
120             ObjData(Obj).grhindex = Val(Leer.GetValue("OBJ" & Obj, "grhindex"))
                ' Español
122             ObjData(Obj).Name = Leer.GetValue("OBJ" & Obj, "Name")
124             ObjData(Obj).texto = Leer.GetValue("OBJ" & Obj, "Texto")
126             ObjData(Obj).Info = Leer.GetValue("OBJ" & Obj, "Info")
                ' Inglés
128             ObjData(Obj).en_name = Leer.GetValue("OBJ" & Obj, "en_Name")
130             ObjData(Obj).en_texto = Leer.GetValue("OBJ" & Obj, "en_Texto")
132             ObjData(Obj).en_Info = Leer.GetValue("OBJ" & Obj, "en_Info")
                ' Portugués
134             ObjData(Obj).pt_name = Leer.GetValue("OBJ" & Obj, "pt_Name")
136             ObjData(Obj).pt_texto = Leer.GetValue("OBJ" & Obj, "pt_Texto")
138             ObjData(Obj).pt_Info = Leer.GetValue("OBJ" & Obj, "pt_Info")
                ' Francés
140             ObjData(Obj).fr_name = Leer.GetValue("OBJ" & Obj, "fr_Name")
142             ObjData(Obj).fr_texto = Leer.GetValue("OBJ" & Obj, "fr_Texto")
144             ObjData(Obj).fr_Info = Leer.GetValue("OBJ" & Obj, "fr_Info")
                ' Italiano
146             ObjData(Obj).it_name = Leer.GetValue("OBJ" & Obj, "it_Name")
148             ObjData(Obj).it_texto = Leer.GetValue("OBJ" & Obj, "it_Texto")
150             ObjData(Obj).it_Info = Leer.GetValue("OBJ" & Obj, "it_Info")
                ' Resto de propiedades
152             ObjData(Obj).MINDEF = Val(Leer.GetValue("OBJ" & Obj, "MinDef"))
154             ObjData(Obj).MaxDEF = Val(Leer.GetValue("OBJ" & Obj, "MaxDef"))
156             ObjData(Obj).MinHit = Val(Leer.GetValue("OBJ" & Obj, "MinHit"))
158             ObjData(Obj).MaxHit = Val(Leer.GetValue("OBJ" & Obj, "MaxHit"))
160             ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
162             ObjData(Obj).CreaGRH = Leer.GetValue("OBJ" & Obj, "CreaGRH")
164             ObjData(Obj).CreaLuz = Leer.GetValue("OBJ" & Obj, "CreaLuz")
166             ObjData(Obj).CreaParticulaPiso = Val(Leer.GetValue("OBJ" & Obj, "CreaParticulaPiso"))
168             ObjData(Obj).Proyectil = Val(Leer.GetValue("OBJ" & Obj, "Proyectil"))
170             ObjData(Obj).Hechizo = Val(Leer.GetValue("OBJ" & Obj, "Hechizo"))
172             ObjData(Obj).Raices = Val(Leer.GetValue("OBJ" & Obj, "Raices"))
174             ObjData(Obj).Cuchara = Val(Leer.GetValue("OBJ" & Obj, "Cuchara"))
176             ObjData(Obj).Botella = Val(Leer.GetValue("OBJ" & Obj, "Botella"))
178             ObjData(Obj).Mortero = Val(Leer.GetValue("OBJ" & Obj, "Mortero"))
180             ObjData(Obj).FrascoAlq = Val(Leer.GetValue("OBJ" & Obj, "FrascoAlq"))
182             ObjData(Obj).FrascoElixir = Val(Leer.GetValue("OBJ" & Obj, "FrascoElixir"))
184             ObjData(Obj).Dosificador = Val(Leer.GetValue("OBJ" & Obj, "Dosificador"))
186             ObjData(Obj).Orquidea = Val(Leer.GetValue("OBJ" & Obj, "Orquidea"))
188             ObjData(Obj).Carmesi = Val(Leer.GetValue("OBJ" & Obj, "Carmesi"))
190             ObjData(Obj).HongoDeLuz = Val(Leer.GetValue("OBJ" & Obj, "HongoDeLuz"))
192             ObjData(Obj).Esporas = Val(Leer.GetValue("OBJ" & Obj, "Esporas"))
194             ObjData(Obj).Tuna = Val(Leer.GetValue("OBJ" & Obj, "Tuna"))
196             ObjData(Obj).Cala = Val(Leer.GetValue("OBJ" & Obj, "Cala"))
198             ObjData(Obj).ColaDeZorro = Val(Leer.GetValue("OBJ" & Obj, "ColaDeZorro"))
200             ObjData(Obj).FlorOceano = Val(Leer.GetValue("OBJ" & Obj, "FlorOceano"))
202             ObjData(Obj).FlorRoja = Val(Leer.GetValue("OBJ" & Obj, "FlorRoja"))
204             ObjData(Obj).Hierva = Val(Leer.GetValue("OBJ" & Obj, "Hierva"))
206             ObjData(Obj).HojasDeRin = Val(Leer.GetValue("OBJ" & Obj, "HojasDeRin"))
208             ObjData(Obj).HojasRojas = Val(Leer.GetValue("OBJ" & Obj, "HojasRojas"))
210             ObjData(Obj).SemillasPros = Val(Leer.GetValue("OBJ" & Obj, "SemillasPros"))
212             ObjData(Obj).Pimiento = Val(Leer.GetValue("OBJ" & Obj, "Pimiento"))
214             ObjData(Obj).Madera = Val(Leer.GetValue("OBJ" & Obj, "Madera"))
216             ObjData(Obj).MaderaElfica = Val(Leer.GetValue("OBJ" & Obj, "MaderaElfica"))
218             ObjData(Obj).PielLobo = Val(Leer.GetValue("OBJ" & Obj, "PielLobo"))
220             ObjData(Obj).PielLoboNegro = Val(Leer.GetValue("OBJ" & Obj, "PielLoboNegro"))
222             ObjData(Obj).PielTigre = Val(Leer.GetValue("OBJ" & Obj, "PielTigre"))
224             ObjData(Obj).PielOsoPardo = Val(Leer.GetValue("OBJ" & Obj, "PielOsoPardo"))
226             ObjData(Obj).PielTigreBengala = Val(Leer.GetValue("OBJ" & Obj, "PielTigreBengala"))
228             ObjData(Obj).PielOsoPolar = Val(Leer.GetValue("OBJ" & Obj, "PielOsoPolar"))
230             ObjData(Obj).LingH = Val(Leer.GetValue("OBJ" & Obj, "LingH"))
232             ObjData(Obj).LingP = Val(Leer.GetValue("OBJ" & Obj, "LingP"))
234             ObjData(Obj).LingO = Val(Leer.GetValue("OBJ" & Obj, "LingO"))
236             ObjData(Obj).Coal = Val(Leer.GetValue("OBJ" & Obj, "Coal"))
238             ObjData(Obj).Destruye = Val(Leer.GetValue("OBJ" & Obj, "Destruye"))
240             ObjData(Obj).SkHerreria = Val(Leer.GetValue("OBJ" & Obj, "SkHerreria"))
242             ObjData(Obj).SkPociones = Val(Leer.GetValue("OBJ" & Obj, "SkPociones"))
244             ObjData(Obj).Sksastreria = Val(Leer.GetValue("OBJ" & Obj, "Sksastreria"))
246             ObjData(Obj).Valor = Val(Leer.GetValue("OBJ" & Obj, "Valor"))
248             ObjData(Obj).Agarrable = Val(Leer.GetValue("OBJ" & Obj, "Agarrable"))
250             ObjData(Obj).Llave = Val(Leer.GetValue("OBJ" & Obj, "Llave"))
252             ObjData(Obj).Municiones = Val(Leer.GetValue("OBJ" & Obj, "Municiones"))
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
272         Call Manager.ChangeValue("INIT", "NumOBJs", numobjs)

274         For Obj = 1 To numobjs
276             DoEvents
278             Call Manager.ChangeValue("OBJ" & Obj, "GrhIndex", ObjData(Obj).grhindex)
                ' Español
280             If Len(ObjData(Obj).Name) <> 0 Then
282                 Call Manager.ChangeValue("OBJ" & Obj, "Name", ObjData(Obj).Name)

                End If

284             If Len(ObjData(Obj).texto) <> 0 Then
286                 Call Manager.ChangeValue("OBJ" & Obj, "Texto", ObjData(Obj).texto)

                End If

288             If Len(ObjData(Obj).Info) <> 0 Then
290                 Call Manager.ChangeValue("OBJ" & Obj, "Info", ObjData(Obj).Info)

                End If

                ' Inglés
292             If Len(ObjData(Obj).en_name) <> 0 Then
294                 Call Manager.ChangeValue("OBJ" & Obj, "en_Name", ObjData(Obj).en_name)

                End If

296             If Len(ObjData(Obj).en_texto) <> 0 Then
298                 Call Manager.ChangeValue("OBJ" & Obj, "en_Texto", ObjData(Obj).en_texto)

                End If

300             If Len(ObjData(Obj).en_Info) <> 0 Then
302                 Call Manager.ChangeValue("OBJ" & Obj, "en_Info", ObjData(Obj).en_Info)

                End If

                ' Portugués
304             If Len(ObjData(Obj).pt_name) <> 0 Then
306                 Call Manager.ChangeValue("OBJ" & Obj, "pt_Name", ObjData(Obj).pt_name)

                End If

308             If Len(ObjData(Obj).pt_texto) <> 0 Then
310                 Call Manager.ChangeValue("OBJ" & Obj, "pt_Texto", ObjData(Obj).pt_texto)

                End If

312             If Len(ObjData(Obj).pt_Info) <> 0 Then
314                 Call Manager.ChangeValue("OBJ" & Obj, "pt_Info", ObjData(Obj).pt_Info)

                End If

                ' Francés
316             If Len(ObjData(Obj).fr_name) <> 0 Then
318                 Call Manager.ChangeValue("OBJ" & Obj, "fr_Name", ObjData(Obj).fr_name)

                End If

320             If Len(ObjData(Obj).fr_texto) <> 0 Then
322                 Call Manager.ChangeValue("OBJ" & Obj, "fr_Texto", ObjData(Obj).fr_texto)

                End If

324             If Len(ObjData(Obj).fr_Info) <> 0 Then
326                 Call Manager.ChangeValue("OBJ" & Obj, "fr_Info", ObjData(Obj).fr_Info)

                End If

                ' Italiano
328             If Len(ObjData(Obj).it_name) <> 0 Then
330                 Call Manager.ChangeValue("OBJ" & Obj, "it_Name", ObjData(Obj).it_name)

                End If

332             If Len(ObjData(Obj).it_texto) <> 0 Then
334                 Call Manager.ChangeValue("OBJ" & Obj, "it_Texto", ObjData(Obj).it_texto)

                End If

336             If Len(ObjData(Obj).it_Info) <> 0 Then
338                 Call Manager.ChangeValue("OBJ" & Obj, "it_Info", ObjData(Obj).it_Info)

                End If

340             If Len(ObjData(Obj).en_Info) <> 0 Then
342                 Call Manager.ChangeValue("OBJ" & Obj, "en_Info", ObjData(Obj).en_Info)

                End If

344             If ObjData(Obj).MINDEF > 0 Then
346                 Call Manager.ChangeValue("OBJ" & Obj, "MINDEF", ObjData(Obj).MINDEF)

                End If

348             If ObjData(Obj).MaxDEF > 0 Then
350                 Call Manager.ChangeValue("OBJ" & Obj, "MaxDEF", ObjData(Obj).MaxDEF)

                End If

352             If ObjData(Obj).MinHit > 0 Then
354                 Call Manager.ChangeValue("OBJ" & Obj, "MinHIt", ObjData(Obj).MinHit)

                End If

356             If ObjData(Obj).MaxHit > 0 Then
358                 Call Manager.ChangeValue("OBJ" & Obj, "maxhit", ObjData(Obj).MaxHit)

                End If

360             If ObjData(Obj).ObjType > 0 Then
362                 Call Manager.ChangeValue("OBJ" & Obj, "ObjType", ObjData(Obj).ObjType)

                End If

364             If Len(ObjData(Obj).CreaLuz) <> 0 Then
366                 Call Manager.ChangeValue("OBJ" & Obj, "CreaLuz", ObjData(Obj).CreaLuz)

                End If

368             If Len(ObjData(Obj).CreaGRH) <> 0 Then
370                 Call Manager.ChangeValue("OBJ" & Obj, "CreaGRH", ObjData(Obj).CreaGRH)

                End If

372             If ObjData(Obj).Hechizo <> 0 Then
374                 Call Manager.ChangeValue("OBJ" & Obj, "Hechizo", ObjData(Obj).Hechizo)

                End If

376             If ObjData(Obj).Raices <> 0 Then
378                 Call Manager.ChangeValue("OBJ" & Obj, "Raices", ObjData(Obj).Raices)

                End If

380             If ObjData(Obj).Cuchara <> 0 Then
382                 Call Manager.ChangeValue("OBJ" & Obj, "Cuchara", ObjData(Obj).Cuchara)

                End If

384             If ObjData(Obj).Botella <> 0 Then
386                 Call Manager.ChangeValue("OBJ" & Obj, "Botella", ObjData(Obj).Botella)

                End If

388             If ObjData(Obj).Mortero <> 0 Then
390                 Call Manager.ChangeValue("OBJ" & Obj, "Mortero", ObjData(Obj).Mortero)

                End If

392             If ObjData(Obj).FrascoAlq <> 0 Then
394                 Call Manager.ChangeValue("OBJ" & Obj, "FrascoAlq", ObjData(Obj).FrascoAlq)

                End If

396             If ObjData(Obj).FrascoElixir <> 0 Then
398                 Call Manager.ChangeValue("OBJ" & Obj, "FrascoElixir", ObjData(Obj).FrascoElixir)

                End If

400             If ObjData(Obj).Dosificador <> 0 Then
402                 Call Manager.ChangeValue("OBJ" & Obj, "Dosificador", ObjData(Obj).Dosificador)

                End If

404             If ObjData(Obj).Orquidea <> 0 Then
406                 Call Manager.ChangeValue("OBJ" & Obj, "Orquidea", ObjData(Obj).Orquidea)

                End If

408             If ObjData(Obj).Carmesi <> 0 Then
410                 Call Manager.ChangeValue("OBJ" & Obj, "Carmesi", ObjData(Obj).Carmesi)

                End If

412             If ObjData(Obj).HongoDeLuz <> 0 Then
414                 Call Manager.ChangeValue("OBJ" & Obj, "HongoDeLuz", ObjData(Obj).HongoDeLuz)

                End If

416             If ObjData(Obj).Esporas <> 0 Then
418                 Call Manager.ChangeValue("OBJ" & Obj, "Esporas", ObjData(Obj).Esporas)

                End If

420             If ObjData(Obj).Tuna <> 0 Then
422                 Call Manager.ChangeValue("OBJ" & Obj, "Tuna", ObjData(Obj).Tuna)

                End If

424             If ObjData(Obj).Cala <> 0 Then
426                 Call Manager.ChangeValue("OBJ" & Obj, "Cala", ObjData(Obj).Cala)

                End If

428             If ObjData(Obj).ColaDeZorro <> 0 Then
430                 Call Manager.ChangeValue("OBJ" & Obj, "ColaDeZorro", ObjData(Obj).ColaDeZorro)

                End If

432             If ObjData(Obj).FlorOceano <> 0 Then
434                 Call Manager.ChangeValue("OBJ" & Obj, "FlorOceano", ObjData(Obj).FlorOceano)

                End If

436             If ObjData(Obj).FlorRoja <> 0 Then
438                 Call Manager.ChangeValue("OBJ" & Obj, "FlorRoja", ObjData(Obj).FlorRoja)

                End If

440             If ObjData(Obj).Hierva <> 0 Then
442                 Call Manager.ChangeValue("OBJ" & Obj, "Hierva", ObjData(Obj).Hierva)

                End If

444             If ObjData(Obj).HojasDeRin <> 0 Then
446                 Call Manager.ChangeValue("OBJ" & Obj, "HojasDeRin", ObjData(Obj).HojasDeRin)

                End If

448             If ObjData(Obj).HojasRojas <> 0 Then
450                 Call Manager.ChangeValue("OBJ" & Obj, "HojasRojas", ObjData(Obj).HojasRojas)

                End If

452             If ObjData(Obj).SemillasPros <> 0 Then
454                 Call Manager.ChangeValue("OBJ" & Obj, "SemillasPros", ObjData(Obj).SemillasPros)

                End If

456             If ObjData(Obj).Pimiento <> 0 Then
458                 Call Manager.ChangeValue("OBJ" & Obj, "Pimiento", ObjData(Obj).Pimiento)

                End If

460             If ObjData(Obj).Madera <> 0 Then
462                 Call Manager.ChangeValue("OBJ" & Obj, "Madera", ObjData(Obj).Madera)

                End If

464             If ObjData(Obj).MaderaElfica <> 0 Then
466                 Call Manager.ChangeValue("OBJ" & Obj, "MaderaElfica", ObjData(Obj).MaderaElfica)

                End If

468             If ObjData(Obj).PielLobo <> 0 Then
470                 Call Manager.ChangeValue("OBJ" & Obj, "PielLobo", ObjData(Obj).PielLobo)

                End If

472             If ObjData(Obj).PielLoboNegro <> 0 Then
474                 Call Manager.ChangeValue("OBJ" & Obj, "PielLoboNegro", ObjData(Obj).PielLoboNegro)

                End If

476             If ObjData(Obj).PielOsoPardo <> 0 Then
478                 Call Manager.ChangeValue("OBJ" & Obj, "PielOsoPardo", ObjData(Obj).PielOsoPardo)

                End If

480             If ObjData(Obj).PielTigre <> 0 Then
482                 Call Manager.ChangeValue("OBJ" & Obj, "PielTigre", ObjData(Obj).PielTigre)

                End If

484             If ObjData(Obj).PielTigreBengala <> 0 Then
486                 Call Manager.ChangeValue("OBJ" & Obj, "PielTigreBengala", ObjData(Obj).PielTigreBengala)

                End If

488             If ObjData(Obj).PielOsoPolar <> 0 Then
490                 Call Manager.ChangeValue("OBJ" & Obj, "PielOsoPolar", ObjData(Obj).PielOsoPolar)

                End If

492             If ObjData(Obj).LingH <> 0 Then
494                 Call Manager.ChangeValue("OBJ" & Obj, "LingH", ObjData(Obj).LingH)

                End If

496             If ObjData(Obj).LingP <> 0 Then
498                 Call Manager.ChangeValue("OBJ" & Obj, "LingP", ObjData(Obj).LingP)

                End If

500             If ObjData(Obj).LingO <> 0 Then
502                 Call Manager.ChangeValue("OBJ" & Obj, "LingO", ObjData(Obj).LingO)

                End If

504             If ObjData(Obj).Coal <> 0 Then
506                 Call Manager.ChangeValue("OBJ" & Obj, "Coal", ObjData(Obj).Coal)

                End If

508             If ObjData(Obj).Destruye <> 0 Then
510                 Call Manager.ChangeValue("OBJ" & Obj, "Destruye", ObjData(Obj).Destruye)

                End If

512             If ObjData(Obj).SkHerreria <> 0 Then
514                 Call Manager.ChangeValue("OBJ" & Obj, "SkHerreria", ObjData(Obj).SkHerreria)

                End If

516             If ObjData(Obj).SkPociones <> 0 Then
518                 Call Manager.ChangeValue("OBJ" & Obj, "SkPociones", ObjData(Obj).SkPociones)

                End If

520             If ObjData(Obj).Sksastreria <> 0 Then
522                 Call Manager.ChangeValue("OBJ" & Obj, "Sksastreria", ObjData(Obj).Sksastreria)

                End If

524             If ObjData(Obj).Valor <> 0 Then
526                 Call Manager.ChangeValue("OBJ" & Obj, "Valor", ObjData(Obj).Valor)

                End If

528             If ObjData(Obj).Agarrable Then
530                 Call Manager.ChangeValue("OBJ" & Obj, "Agarrable", 1)

                End If

532             If ObjData(Obj).CreaParticulaPiso > 0 Then
534                 Call Manager.ChangeValue("OBJ" & Obj, "CreaParticulaPiso", ObjData(Obj).CreaParticulaPiso)

                End If

536             If ObjData(Obj).Proyectil > 0 Then
538                 Call Manager.ChangeValue("OBJ" & Obj, "Proyectil", ObjData(Obj).Proyectil)

                End If

540             If ObjData(Obj).Municiones > 0 Then
542                 Call Manager.ChangeValue("OBJ" & Obj, "Municiones", ObjData(Obj).Municiones)

                End If

544             If ObjData(Obj).Llave > 0 Then
546                 Call Manager.ChangeValue("OBJ" & Obj, "Llave", ObjData(Obj).Llave)

                End If

548             If ObjData(Obj).Cooldown > 0 Then
550                 Call Manager.ChangeValue("OBJ" & Obj, "CD", ObjData(Obj).Cooldown)

                End If

552             If ObjData(Obj).CdType > 0 Then
554                 Call Manager.ChangeValue("OBJ" & Obj, "CDType", ObjData(Obj).CdType)

                End If

556             If ObjData(Obj).SpellIndex > 0 Then
558                 Call Manager.ChangeValue("OBJ" & Obj, "SpellIndex", ObjData(Obj).SpellIndex)

                End If

560             Label3.Caption = "Grabando: " & Obj & "/" & numobjs
562             Label3.ForeColor = vbGreen
564         Next Obj

566         Label3.ForeColor = vbGreen
568         Label3.Caption = "Creado objindex.dat"
        Else
570         MsgBox "Falta el archivo obj.dat dentro de la carpeta INIT."

        End If

572     If FileExist(App.Path & "\..\Recursos\Dat\npcs.dat", vbNormal) Then
574         NpcFile = App.Path & "\..\Recursos\Dat\npcs.dat"
576         Call Leer.Initialize(NpcFile)
            Dim numnpcs As Long
578         numnpcs = Val(GetVar(NpcFile, "INIT", "NumNPCs"))
580         Label3.Caption = "0/" & numnpcs
582         ReDim NpcData(1 To numnpcs) As NpcDatas
            Dim aux As String

584         For Npc = 1 To numnpcs
586             DoEvents
588             NpcData(Npc).Name = Leer.GetValue("npc" & Npc, "Name")
590             NpcData(Npc).en_name = Leer.GetValue("npc" & Npc, "en_Name")
                ' Portugués
592             NpcData(Npc).pt_name = Leer.GetValue("npc" & Npc, "pt_Name")
                ' Francés
594             NpcData(Npc).fr_name = Leer.GetValue("npc" & Npc, "fr_Name")
                ' Italiano
596             NpcData(Npc).it_name = Leer.GetValue("npc" & Npc, "it_Name")
598             NpcData(Npc).desc = Leer.GetValue("npc" & Npc, "desc")
600             NpcData(Npc).en_Desc = Leer.GetValue("npc" & Npc, "en_desc")
                ' Portugués
602             NpcData(Npc).pt_Desc = Leer.GetValue("npc" & Npc, "pt_desc")
                ' Francés
604             NpcData(Npc).fr_Desc = Leer.GetValue("npc" & Npc, "fr_desc")
                ' Italiano
606             NpcData(Npc).it_Desc = Leer.GetValue("npc" & Npc, "it_desc")
608             NpcData(Npc).Body = Val(Leer.GetValue("npc" & Npc, "Body"))
610             NpcData(Npc).Exp = Val(Leer.GetValue("npc" & Npc, "GiveEXP"))
612             NpcData(Npc).Head = Val(Leer.GetValue("npc" & Npc, "Head"))
614             NpcData(Npc).Hp = Val(Leer.GetValue("npc" & Npc, "MaxHP"))
616             NpcData(Npc).MaxHit = Val(Leer.GetValue("npc" & Npc, "MaxHit"))
618             NpcData(Npc).MinHit = Val(Leer.GetValue("npc" & Npc, "MinHit"))
620             NpcData(Npc).Oro = Val(Leer.GetValue("npc" & Npc, "GiveGLD"))
622             NpcData(Npc).ExpClan = Val(Leer.GetValue("npc" & Npc, "GiveEXPClan"))
624             NpcData(Npc).PuedeInvocar = Val(Leer.GetValue("npc" & Npc, "PuedeInvocar"))
626             NpcData(Npc).NoMapInfo = Val(Leer.GetValue("npc" & Npc, "NoMapInfo"))
628             NpcData(Npc).QuizaProb = Val(Leer.GetValue("npc" & Npc, "QuizaProb"))
630             aux = Val(GetVar(NpcFile, "Npc" & Npc, "NumQuiza"))
632             If aux = 0 Then
634                 NpcData(Npc).NumQuiza = 0
                Else
636                 NpcData(Npc).NumQuiza = Val(aux)
638                 ReDim NpcData(Npc).QuizaDropea(1 To NpcData(Npc).NumQuiza) As Integer
                    Dim LoopC As Long

640                 For LoopC = 1 To NpcData(Npc).NumQuiza
642                     NpcData(Npc).QuizaDropea(LoopC) = Val(Leer.GetValue("npc" & Npc, "QuizaDropea" & LoopC))
644                 Next LoopC

                End If

646             Label3.ForeColor = vbRed
648             Label3.Caption = "Leyendo NPCs: " & Npc & "/" & numnpcs
650         Next Npc

652         Npc = 1
654         Call Manager.ChangeValue("INIT", "NumNPCs", numnpcs)

656         For Npc = 1 To numnpcs
658             DoEvents
660             If Len(NpcData(Npc).Name) <> 0 Then
662                 Call Manager.ChangeValue("Npc" & Npc, "Name", NpcData(Npc).Name)

                End If

664             If Len(NpcData(Npc).en_name) <> 0 Then
666                 Call Manager.ChangeValue("Npc" & Npc, "en_Name", NpcData(Npc).en_name)

                End If

668             If Len(NpcData(Npc).pt_name) <> 0 Then
670                 Call Manager.ChangeValue("Npc" & Npc, "pt_Name", NpcData(Npc).pt_name)

                End If

672             If Len(NpcData(Npc).fr_name) <> 0 Then
674                 Call Manager.ChangeValue("Npc" & Npc, "fr_Name", NpcData(Npc).fr_name)

                End If

676             If Len(NpcData(Npc).it_name) <> 0 Then
678                 Call Manager.ChangeValue("Npc" & Npc, "it_Name", NpcData(Npc).it_name)

                End If

680             If Len(NpcData(Npc).en_Desc) <> 0 Then
682                 Call Manager.ChangeValue("Npc" & Npc, "en_desc", NpcData(Npc).en_Desc)

                End If

684             If Len(NpcData(Npc).pt_Desc) <> 0 Then
686                 Call Manager.ChangeValue("Npc" & Npc, "pt_desc", NpcData(Npc).pt_Desc)

                End If

688             If Len(NpcData(Npc).fr_Desc) <> 0 Then
690                 Call Manager.ChangeValue("Npc" & Npc, "fr_desc", NpcData(Npc).fr_Desc)

                End If

692             If Len(NpcData(Npc).it_Desc) <> 0 Then
694                 Call Manager.ChangeValue("Npc" & Npc, "it_desc", NpcData(Npc).it_Desc)

                End If

696             If Len(NpcData(Npc).desc) <> 0 Then
698                 Call Manager.ChangeValue("Npc" & Npc, "Desc", NpcData(Npc).desc)

                End If

700             If NpcData(Npc).Body <> 0 Then
702                 Call Manager.ChangeValue("Npc" & Npc, "Body", NpcData(Npc).Body)

                End If

704             If NpcData(Npc).Head <> 0 Then
706                 Call Manager.ChangeValue("Npc" & Npc, "Head", NpcData(Npc).Head)

                End If

708             If NpcData(Npc).Exp <> 0 Then
710                 Call Manager.ChangeValue("Npc" & Npc, "Exp", NpcData(Npc).Exp)

                End If

712             If NpcData(Npc).Hp <> 0 Then
714                 Call Manager.ChangeValue("Npc" & Npc, "Hp", NpcData(Npc).Hp)

                End If

716             If NpcData(Npc).MaxHit <> 0 Then
718                 Call Manager.ChangeValue("Npc" & Npc, "MaxHit", NpcData(Npc).MaxHit)

                End If

720             If NpcData(Npc).MinHit <> 0 Then
722                 Call Manager.ChangeValue("Npc" & Npc, "MinHit", NpcData(Npc).MinHit)

                End If

724             If NpcData(Npc).Oro <> 0 Then
726                 Call Manager.ChangeValue("Npc" & Npc, "Oro", NpcData(Npc).Oro)

                End If

728             If NpcData(Npc).ExpClan <> 0 Then
730                 Call Manager.ChangeValue("Npc" & Npc, "GiveEXPClan", NpcData(Npc).ExpClan)

                End If

732             If NpcData(Npc).NumQuiza <> 0 Then
734                 Call Manager.ChangeValue("Npc" & Npc, "NumQuiza", NpcData(Npc).NumQuiza)

736                 For LoopC = 1 To NpcData(Npc).NumQuiza
738                     Call Manager.ChangeValue("Npc" & Npc, "QuizaDropea" & LoopC, NpcData(Npc).QuizaDropea(LoopC))
740                 Next LoopC

                End If

742             If NpcData(Npc).QuizaProb <> 0 Then
744                 Call Manager.ChangeValue("Npc" & Npc, "QuizaProb", NpcData(Npc).QuizaProb)

                End If

746             If NpcData(Npc).NoMapInfo <> 0 Then
748                 Call Manager.ChangeValue("Npc" & Npc, "NoMapInfo", NpcData(Npc).NoMapInfo)

                End If

750             If NpcData(Npc).PuedeInvocar <> 0 Then
752                 Call Manager.ChangeValue("Npc" & Npc, "PuedeInvocar", NpcData(Npc).PuedeInvocar)

                End If

754             Label3.Caption = "Grabando NPCs: " & Npc & "/" & numnpcs
756             Label3.ForeColor = vbGreen
758         Next Npc

        Else
760         MsgBox "Falta el archivo npcs.dat dentro de la carpeta dats."
        End If

762     If FileExist(App.Path & "\..\Recursos\Dat\hechizos.dat", vbNormal) Then
            Dim hechizosFile As String, numhechizos As Long
764         hechizosFile = App.Path & "\..\Recursos\Dat\hechizos.dat"
766         numhechizos = Val(GetVar(hechizosFile, "INIT", "NumeroHechizos"))
            Dim hechic As New clsIniReader
768         Call hechic.Initialize(hechizosFile)
770         Label3.Caption = "Leyendo Hechizos: " & "0/" & numhechizos
772         ReDim HechizoData(1 To numhechizos) As HechizoDatas

774         For Hechizo = 1 To numhechizos
776             DoEvents
778             HechizoData(Hechizo).Nombre = hechic.GetValue("Hechizo" & Hechizo, "Nombre")
780             HechizoData(Hechizo).en_Nombre = hechic.GetValue("Hechizo" & Hechizo, "en_Nombre")
782             HechizoData(Hechizo).pt_Nombre = hechic.GetValue("Hechizo" & Hechizo, "pt_Nombre")
784             HechizoData(Hechizo).fr_Nombre = hechic.GetValue("Hechizo" & Hechizo, "fr_Nombre")
786             HechizoData(Hechizo).it_Nombre = hechic.GetValue("Hechizo" & Hechizo, "it_Nombre")
788             HechizoData(Hechizo).desc = hechic.GetValue("Hechizo" & Hechizo, "desc")
790             HechizoData(Hechizo).en_Desc = hechic.GetValue("Hechizo" & Hechizo, "en_Desc")
792             HechizoData(Hechizo).pt_Desc = hechic.GetValue("Hechizo" & Hechizo, "pt_Desc")
794             HechizoData(Hechizo).fr_Desc = hechic.GetValue("Hechizo" & Hechizo, "fr_Desc")
796             HechizoData(Hechizo).it_Desc = hechic.GetValue("Hechizo" & Hechizo, "it_Desc")
798             HechizoData(Hechizo).PalabrasMagicas = hechic.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
800             HechizoData(Hechizo).HechizeroMsg = hechic.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
802             HechizoData(Hechizo).en_HechizeroMsg = hechic.GetValue("Hechizo" & Hechizo, "en_HechizeroMsg")
804             HechizoData(Hechizo).pt_HechizeroMsg = hechic.GetValue("Hechizo" & Hechizo, "pt_HechizeroMsg")
806             HechizoData(Hechizo).fr_HechizeroMsg = hechic.GetValue("Hechizo" & Hechizo, "fr_HechizeroMsg")
808             HechizoData(Hechizo).it_HechizeroMsg = hechic.GetValue("Hechizo" & Hechizo, "it_HechizeroMsg")
810             HechizoData(Hechizo).TargetMsg = hechic.GetValue("Hechizo" & Hechizo, "TargetMsg")
812             HechizoData(Hechizo).en_TargetMsg = hechic.GetValue("Hechizo" & Hechizo, "en_TargetMsg")
814             HechizoData(Hechizo).pt_TargetMsg = hechic.GetValue("Hechizo" & Hechizo, "pt_TargetMsg")
816             HechizoData(Hechizo).fr_TargetMsg = hechic.GetValue("Hechizo" & Hechizo, "fr_TargetMsg")
818             HechizoData(Hechizo).it_TargetMsg = hechic.GetValue("Hechizo" & Hechizo, "it_TargetMsg")
820             HechizoData(Hechizo).PropioMsg = hechic.GetValue("Hechizo" & Hechizo, "PropioMsg")
822             HechizoData(Hechizo).en_PropioMsg = hechic.GetValue("Hechizo" & Hechizo, "en_PropioMsg")
824             HechizoData(Hechizo).pt_PropioMsg = hechic.GetValue("Hechizo" & Hechizo, "pt_PropioMsg")
826             HechizoData(Hechizo).fr_PropioMsg = hechic.GetValue("Hechizo" & Hechizo, "fr_PropioMsg")
828             HechizoData(Hechizo).it_PropioMsg = hechic.GetValue("Hechizo" & Hechizo, "it_PropioMsg")
830             HechizoData(Hechizo).ManaRequerido = Val(hechic.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
832             HechizoData(Hechizo).StaRequerido = Val(hechic.GetValue("Hechizo" & Hechizo, "StaRequerido"))
834             HechizoData(Hechizo).MinSkill = Val(hechic.GetValue("Hechizo" & Hechizo, "MinSkill"))
836             HechizoData(Hechizo).StaRequerido = Val(hechic.GetValue("Hechizo" & Hechizo, "StaRequerido"))
838             HechizoData(Hechizo).IconoIndex = Val(hechic.GetValue("Hechizo" & Hechizo, "IconoIndex"))
840             HechizoData(Hechizo).Cooldown = Val(hechic.GetValue("Hechizo" & Hechizo, "Cooldown"))
842             Label3.ForeColor = vbRed
844             Label3.Caption = "Leyendo: " & Hechizo & "/" & numhechizos
846         Next Hechizo

848         Call Manager.ChangeValue("INIT", "NumeroHechizo", numhechizos)

850         For Hechizo = 1 To numhechizos
852             DoEvents
                ' Español
854             Call Manager.ChangeValue("Hechizo" & Hechizo, "Nombre", HechizoData(Hechizo).Nombre)
856             Call Manager.ChangeValue("Hechizo" & Hechizo, "Desc", HechizoData(Hechizo).desc)
858             Call Manager.ChangeValue("Hechizo" & Hechizo, "PalabrasMagicas", HechizoData(Hechizo).PalabrasMagicas)
860             Call Manager.ChangeValue("Hechizo" & Hechizo, "HechizeroMsg", HechizoData(Hechizo).HechizeroMsg)
862             Call Manager.ChangeValue("Hechizo" & Hechizo, "TargetMsg", HechizoData(Hechizo).TargetMsg)
864             Call Manager.ChangeValue("Hechizo" & Hechizo, "PropioMsg", HechizoData(Hechizo).PropioMsg)
                ' Inglés
866             Call Manager.ChangeValue("Hechizo" & Hechizo, "en_Name", HechizoData(Hechizo).en_Nombre)
868             Call Manager.ChangeValue("Hechizo" & Hechizo, "en_Desc", HechizoData(Hechizo).en_Desc)
870             Call Manager.ChangeValue("Hechizo" & Hechizo, "en_HechizeroMsg", HechizoData(Hechizo).en_HechizeroMsg)
872             Call Manager.ChangeValue("Hechizo" & Hechizo, "en_TargetMsg", HechizoData(Hechizo).en_TargetMsg)
874             Call Manager.ChangeValue("Hechizo" & Hechizo, "en_PropioMsg", HechizoData(Hechizo).en_PropioMsg)
                ' Portugués
876             Call Manager.ChangeValue("Hechizo" & Hechizo, "pt_Nombre", HechizoData(Hechizo).pt_Nombre)
878             Call Manager.ChangeValue("Hechizo" & Hechizo, "pt_Desc", HechizoData(Hechizo).pt_Desc)
880             Call Manager.ChangeValue("Hechizo" & Hechizo, "pt_HechizeroMsg", HechizoData(Hechizo).pt_HechizeroMsg)
882             Call Manager.ChangeValue("Hechizo" & Hechizo, "pt_TargetMsg", HechizoData(Hechizo).pt_TargetMsg)
884             Call Manager.ChangeValue("Hechizo" & Hechizo, "pt_PropioMsg", HechizoData(Hechizo).pt_PropioMsg)
                ' Francés
886             Call Manager.ChangeValue("Hechizo" & Hechizo, "fr_Nombre", HechizoData(Hechizo).fr_Nombre)
888             Call Manager.ChangeValue("Hechizo" & Hechizo, "fr_Desc", HechizoData(Hechizo).fr_Desc)
890             Call Manager.ChangeValue("Hechizo" & Hechizo, "fr_HechizeroMsg", HechizoData(Hechizo).fr_HechizeroMsg)
892             Call Manager.ChangeValue("Hechizo" & Hechizo, "fr_TargetMsg", HechizoData(Hechizo).fr_TargetMsg)
894             Call Manager.ChangeValue("Hechizo" & Hechizo, "fr_PropioMsg", HechizoData(Hechizo).fr_PropioMsg)
                ' Italiano
896             Call Manager.ChangeValue("Hechizo" & Hechizo, "it_Nombre", HechizoData(Hechizo).it_Nombre)
898             Call Manager.ChangeValue("Hechizo" & Hechizo, "it_Desc", HechizoData(Hechizo).it_Desc)
900             Call Manager.ChangeValue("Hechizo" & Hechizo, "it_HechizeroMsg", HechizoData(Hechizo).it_HechizeroMsg)
902             Call Manager.ChangeValue("Hechizo" & Hechizo, "it_TargetMsg", HechizoData(Hechizo).it_TargetMsg)
904             Call Manager.ChangeValue("Hechizo" & Hechizo, "it_PropioMsg", HechizoData(Hechizo).it_PropioMsg)
                ' Stats
906             Call Manager.ChangeValue("Hechizo" & Hechizo, "ManaRequerido", HechizoData(Hechizo).ManaRequerido)
908             Call Manager.ChangeValue("Hechizo" & Hechizo, "StaRequerido", HechizoData(Hechizo).StaRequerido)
910             Call Manager.ChangeValue("Hechizo" & Hechizo, "MinSkill", HechizoData(Hechizo).MinSkill)
912             Call Manager.ChangeValue("Hechizo" & Hechizo, "IconoIndex", HechizoData(Hechizo).IconoIndex)
914             Call Manager.ChangeValue("Hechizo" & Hechizo, "Cooldown", HechizoData(Hechizo).Cooldown)
916             Label3.Caption = "Grabando Hechizos: " & Hechizo & "/" & numhechizos
918             Label3.ForeColor = vbGreen
920         Next Hechizo

        End If

        Dim MsgFile            As String
        Dim Msgsss             As New clsIniReader
        Dim NumLocaleSP_Msg    As Long, NumLocaleEN_Msg As Long
        Dim NumLocalePT_Msg    As Long, NumLocaleFR_Msg As Long, NumLocaleIT_Msg As Long
        Dim SP_MSG             As Integer, EN_MSG As Integer
        Dim PT_MSG             As Integer, FR_MSG As Integer, IT_MSG As Integer
        Dim arrLocale_SP_SMG() As String, arrLocale_EN_SMG() As String
        Dim arrLocale_PT_SMG() As String, arrLocale_FR_SMG() As String, arrLocale_IT_SMG() As String
        ' Español
922     If FileExist(App.Path & "\..\Recursos\init\SP_LocalMsg.dat", vbNormal) Then
924         MsgFile = App.Path & "\..\Recursos\init\SP_LocalMsg.dat"
926         Call Msgsss.Initialize(MsgFile)
928         NumLocaleSP_Msg = Val(Msgsss.GetValue("INIT", "NumLocaleSP_Msg"))
930         Label3.Caption = "0/" & CStr(NumLocaleSP_Msg)
932         ReDim arrLocale_SP_SMG(1 To NumLocaleSP_Msg)

934         For SP_MSG = 1 To NumLocaleSP_Msg
936             DoEvents
938             arrLocale_SP_SMG(SP_MSG) = Msgsss.GetValue("SP_MSG", "Msg" & SP_MSG)
940             Label3.ForeColor = vbRed
942             Label3.Caption = "Leyendo MSG SP: " & SP_MSG & "/" & NumLocaleSP_Msg
944         Next SP_MSG

946         Call Manager.ChangeValue("INIT", "NumLocaleMsg", NumLocaleSP_Msg)

948         For SP_MSG = 1 To NumLocaleSP_Msg
950             DoEvents
952             Call Manager.ChangeValue("SP_Msg", "Msg" & SP_MSG, arrLocale_SP_SMG(SP_MSG))
954             Label3.Caption = "Grabando MSG SP: " & SP_MSG & "/" & NumLocaleSP_Msg
956             Label3.ForeColor = vbGreen
958         Next SP_MSG

        Else
960         MsgBox "Falta el archivo SP_LocalMsg.dat dentro de la carpeta INIT."

        End If

        ' Inglés
962     If FileExist(App.Path & "\..\Recursos\init\EN_LocalMsg.dat", vbNormal) Then
964         MsgFile = App.Path & "\..\Recursos\init\EN_LocalMsg.dat"
966         Call Msgsss.Initialize(MsgFile)
968         NumLocaleEN_Msg = Val(Msgsss.GetValue("INIT", "NumLocaleEN_Msg"))
970         Label3.Caption = "0/" & CStr(NumLocaleEN_Msg)
972         ReDim arrLocale_EN_SMG(1 To NumLocaleEN_Msg)

974         For EN_MSG = 1 To NumLocaleEN_Msg
976             DoEvents
978             arrLocale_EN_SMG(EN_MSG) = Msgsss.GetValue("EN_MSG", "Msg" & EN_MSG)
980             Label3.ForeColor = vbRed
982             Label3.Caption = "Leyendo MSG EN: " & EN_MSG & "/" & NumLocaleEN_Msg
984         Next EN_MSG

986         Call Manager.ChangeValue("INIT", "NumLocaleMsg", NumLocaleEN_Msg)

988         For EN_MSG = 1 To NumLocaleEN_Msg
990             DoEvents
992             Call Manager.ChangeValue("EN_Msg", "Msg" & EN_MSG, arrLocale_EN_SMG(EN_MSG))
994             Label3.Caption = "Grabando MSG EN: " & EN_MSG & "/" & NumLocaleEN_Msg
996             Label3.ForeColor = vbGreen
998         Next EN_MSG

         Else
1000         MsgBox "Falta el archivo EN_LocalMsg.dat dentro de la carpeta INIT."

         End If

         ' Portugués
1002     If FileExist(App.Path & "\..\Recursos\init\PT_LocalMsg.dat", vbNormal) Then
1004         MsgFile = App.Path & "\..\Recursos\init\PT_LocalMsg.dat"
1006         Call Msgsss.Initialize(MsgFile)
1008         NumLocalePT_Msg = Val(Msgsss.GetValue("INIT", "NumLocalePT_Msg"))
1010         Label3.Caption = "0/" & CStr(NumLocalePT_Msg)
1012         ReDim arrLocale_PT_SMG(1 To NumLocalePT_Msg)

1014         For PT_MSG = 1 To NumLocalePT_Msg
1016             DoEvents
1018             arrLocale_PT_SMG(PT_MSG) = Msgsss.GetValue("PT_MSG", "Msg" & PT_MSG)
1020             Label3.ForeColor = vbRed
1022             Label3.Caption = "Leyendo MSG PT: " & PT_MSG & "/" & NumLocalePT_Msg
1024         Next PT_MSG

1026         Call Manager.ChangeValue("INIT", "NumLocaleMsg", NumLocalePT_Msg)

1028         For PT_MSG = 1 To NumLocalePT_Msg
1030             DoEvents
1032             Call Manager.ChangeValue("PT_Msg", "Msg" & PT_MSG, arrLocale_PT_SMG(PT_MSG))
1034             Label3.Caption = "Grabando MSG PT: " & PT_MSG & "/" & NumLocalePT_Msg
1036             Label3.ForeColor = vbGreen
1038         Next PT_MSG

         Else
1040         MsgBox "Falta el archivo PT_LocalMsg.dat dentro de la carpeta INIT."

         End If

         ' Francés
1042     If FileExist(App.Path & "\..\Recursos\init\FR_LocalMsg.dat", vbNormal) Then
1044         MsgFile = App.Path & "\..\Recursos\init\FR_LocalMsg.dat"
1046         Call Msgsss.Initialize(MsgFile)
1048         NumLocaleFR_Msg = Val(Msgsss.GetValue("INIT", "NumLocaleFR_Msg"))
1050         Label3.Caption = "0/" & CStr(NumLocaleFR_Msg)
1052         ReDim arrLocale_FR_SMG(1 To NumLocaleFR_Msg)

1054         For FR_MSG = 1 To NumLocaleFR_Msg
1056             DoEvents
1058             arrLocale_FR_SMG(FR_MSG) = Msgsss.GetValue("FR_MSG", "Msg" & FR_MSG)
1060             Label3.ForeColor = vbRed
1062             Label3.Caption = "Leyendo MSG FR: " & FR_MSG & "/" & NumLocaleFR_Msg
1064         Next FR_MSG

1066         Call Manager.ChangeValue("INIT", "NumLocaleMsg", NumLocaleFR_Msg)

1068         For FR_MSG = 1 To NumLocaleFR_Msg
1070             DoEvents
1072             Call Manager.ChangeValue("FR_Msg", "Msg" & FR_MSG, arrLocale_FR_SMG(FR_MSG))
1074             Label3.Caption = "Grabando MSG FR: " & FR_MSG & "/" & NumLocaleFR_Msg
1076             Label3.ForeColor = vbGreen
1078         Next FR_MSG

         Else
1080         MsgBox "Falta el archivo FR_LocalMsg.dat dentro de la carpeta INIT."

         End If

         ' Italiano
1082     If FileExist(App.Path & "\..\Recursos\init\IT_LocalMsg.dat", vbNormal) Then
1084         MsgFile = App.Path & "\..\Recursos\init\IT_LocalMsg.dat"
1086         Call Msgsss.Initialize(MsgFile)
1088         NumLocaleIT_Msg = Val(Msgsss.GetValue("INIT", "NumLocaleIT_Msg"))
1090         Label3.Caption = "0/" & CStr(NumLocaleIT_Msg)
1092         ReDim arrLocale_IT_SMG(1 To NumLocaleIT_Msg)

1094         For IT_MSG = 1 To NumLocaleIT_Msg
1096             DoEvents
1098             arrLocale_IT_SMG(IT_MSG) = Msgsss.GetValue("IT_MSG", "Msg" & IT_MSG)
1100             Label3.ForeColor = vbRed
1102             Label3.Caption = "Leyendo MSG IT: " & IT_MSG & "/" & NumLocaleIT_Msg
1104         Next IT_MSG

1106         Call Manager.ChangeValue("INIT", "NumLocaleMsg", NumLocaleIT_Msg)

1108         For IT_MSG = 1 To NumLocaleIT_Msg
1110             DoEvents
1112             Call Manager.ChangeValue("IT_Msg", "Msg" & IT_MSG, arrLocale_IT_SMG(IT_MSG))
1114             Label3.Caption = "Grabando MSG IT: " & IT_MSG & "/" & NumLocaleIT_Msg
1116             Label3.ForeColor = vbGreen
1118         Next IT_MSG

         Else
1120         MsgBox "Falta el archivo IT_LocalMsg.dat dentro de la carpeta INIT."

         End If

1122     If FileExist(App.Path & "\..\Recursos\init\NameMapa.dat", vbNormal) Then
             Dim MapFile As String
1124         MapFile = App.Path & "\..\Recursos\init\NameMapa.dat"
             Dim Mapa As New clsIniReader
1126         Call Mapa.Initialize(MapFile)
1128         Label3.Caption = "0/" & 750
1130         ReDim MapName(1 To 750) As String
1132         ReDim MapDesc(1 To 750) As String

1134         For Npc = 1 To 750
1136             DoEvents
1138             MapName(Npc) = Mapa.GetValue("NameMapa", "mapa" & Npc)
1140             MapDesc(Npc) = Mapa.GetValue("NameMapa", "mapa" & Npc & "desc")
1142             Label3.ForeColor = vbRed
1144             Label3.Caption = "Leyendo Mapas: " & Npc & "/" & 750
1146         Next Npc

1148         Npc = 1
1150         Call Manager.ChangeValue("INIT", "NumMapas", 750)

1152         For Npc = 1 To 750
1154             DoEvents
1156             Call Manager.ChangeValue("NAMEMAPA", "Mapa" & Npc, MapName(Npc))
1158             Call Manager.ChangeValue("NAMEMAPA", "Mapa" & Npc & "Desc", MapDesc(Npc))
1160             Label3.Caption = "Grabando Mapas: " & Npc & "/" & 750
1162             Label3.ForeColor = vbGreen
1164         Next Npc

         Else
1166         MsgBox "Falta el archivo NameMapa.dat dentro de la carpeta dats."

         End If

         'quest
1168     If FileExist(App.Path & "\..\Recursos\Dat\Quests.DAT", vbNormal) Then
1170         MapFile = App.Path & "\..\Recursos\Dat\Quests.DAT"
1172         Call Mapa.Initialize(MapFile)
             Dim nunquest As Integer
             Dim QUESTNUM As Integer
1174         nunquest = Mapa.GetValue("INIT", "NumQuests")
1176         Label3.Caption = "0/" & nunquest
1178         ReDim QuestNombre(1 To nunquest) As String
1180         ReDim QuestDesc(1 To nunquest) As String
1182         ReDim QuestFin(1 To nunquest) As String
1184         ReDim QuestNameEN(1 To nunquest) As String
1186         ReDim QuestDescEN(1 To nunquest) As String
1188         ReDim QuestFinEN(1 To nunquest) As String
1190         ReDim QuestNamePT(1 To nunquest) As String
1192         ReDim QuestDescPT(1 To nunquest) As String
1194         ReDim QuestFinPT(1 To nunquest) As String
1196         ReDim QuestNameFR(1 To nunquest) As String
1198         ReDim QuestDescFR(1 To nunquest) As String
1200         ReDim QuestFinFR(1 To nunquest) As String
1202         ReDim QuestNameIT(1 To nunquest) As String
1204         ReDim QuestDescIT(1 To nunquest) As String
1206         ReDim QuestFinIT(1 To nunquest) As String
1208         ReDim QuestNext(1 To nunquest) As String
1210         ReDim QuestPos(1 To nunquest) As Integer
1212         ReDim QuestRepetible(1 To nunquest) As Byte
1214         ReDim RequiredLevel(1 To nunquest) As Integer

1216         For QUESTNUM = 1 To nunquest
1218             DoEvents
1220             QuestNombre([QUESTNUM]) = Mapa.GetValue("QUEST" & QUESTNUM, "Nombre")
1222             QuestNameEN([QUESTNUM]) = Mapa.GetValue("QUEST" & QUESTNUM, "en_Nombre")
1224             QuestNamePT([QUESTNUM]) = Mapa.GetValue("QUEST" & QUESTNUM, "pt_Nombre")
1226             QuestNameFR([QUESTNUM]) = Mapa.GetValue("QUEST" & QUESTNUM, "fr_Nombre")
1228             QuestNameIT([QUESTNUM]) = Mapa.GetValue("QUEST" & QUESTNUM, "it_Nombre")
1230             QuestDesc([QUESTNUM]) = Mapa.GetValue("QUEST" & QUESTNUM, "Desc")
1232             QuestDescEN([QUESTNUM]) = Mapa.GetValue("QUEST" & QUESTNUM, "en_Desc")
1234             QuestDescPT([QUESTNUM]) = Mapa.GetValue("QUEST" & QUESTNUM, "pt_Desc")
1236             QuestDescFR([QUESTNUM]) = Mapa.GetValue("QUEST" & QUESTNUM, "fr_Desc")
1238             QuestDescIT(QUESTNUM) = Mapa.GetValue("QUEST" & QUESTNUM, "it_Desc")
1240             QuestFin(QUESTNUM) = Mapa.GetValue("QUEST" & QUESTNUM, "DescFinal")
1242             QuestFinEN(QUESTNUM) = Mapa.GetValue("QUEST" & QUESTNUM, "en_DescFinal")
1244             QuestFinPT(QUESTNUM) = Mapa.GetValue("QUEST" & QUESTNUM, "pt_DescFinal")
1246             QuestFinFR(QUESTNUM) = Mapa.GetValue("QUEST" & QUESTNUM, "fr_DescFinal")
1248             QuestFinIT(QUESTNUM) = Mapa.GetValue("QUEST" & QUESTNUM, "it_DescFinal")
1250             QuestNext(QUESTNUM) = Mapa.GetValue("QUEST" & QUESTNUM, "NextQuest")
1252             QuestRepetible(QUESTNUM) = Val(Mapa.GetValue("QUEST" & QUESTNUM, "Repetible"))
1254             QuestPos(QUESTNUM) = Val(Mapa.GetValue("QUEST" & QUESTNUM, "PosMap"))
1256             RequiredLevel(QUESTNUM) = Val(Mapa.GetValue("QUEST" & QUESTNUM, "RequiredLevel"))
1258             Label3.ForeColor = vbRed
1260             Label3.Caption = "Leyendo Quest: " & QUESTNUM & "/" & nunquest
1262         Next QUESTNUM

1264         Npc = 1
1266         Call Manager.ChangeValue("INIT", "NumQuests", nunquest)

1268         For QUESTNUM = 1 To nunquest
1270             DoEvents
1272             Call Manager.ChangeValue("QUEST" & QUESTNUM, "Nombre", QuestNombre(QUESTNUM))
1274             Call Manager.ChangeValue("QUEST" & QUESTNUM, "en_Nombre", QuestNameEN(QUESTNUM))
1276             Call Manager.ChangeValue("QUEST" & QUESTNUM, "pt_Nombre", QuestNamePT(QUESTNUM))
1278             Call Manager.ChangeValue("QUEST" & QUESTNUM, "fr_Nombre", QuestNameFR(QUESTNUM))
1280             Call Manager.ChangeValue("QUEST" & QUESTNUM, "it_Nombre", QuestNameIT(QUESTNUM))
1282             Call Manager.ChangeValue("QUEST" & QUESTNUM, "Desc", QuestDesc(QUESTNUM))
1284             Call Manager.ChangeValue("QUEST" & QUESTNUM, "en_Desc", QuestDescEN(QUESTNUM))
1286             Call Manager.ChangeValue("QUEST" & QUESTNUM, "pt_Desc", QuestDescPT(QUESTNUM))
1288             Call Manager.ChangeValue("QUEST" & QUESTNUM, "fr_Desc", QuestDescFR(QUESTNUM))
1290             Call Manager.ChangeValue("QUEST" & QUESTNUM, "it_Desc", QuestDescIT(QUESTNUM))
1292             Call Manager.ChangeValue("QUEST" & QUESTNUM, "DescFinal", QuestFin(QUESTNUM))
1294             Call Manager.ChangeValue("QUEST" & QUESTNUM, "en_DescFinal", QuestFinEN(QUESTNUM))
1296             Call Manager.ChangeValue("QUEST" & QUESTNUM, "pt_DescFinal", QuestFinPT(QUESTNUM))
1298             Call Manager.ChangeValue("QUEST" & QUESTNUM, "fr_DescFinal", QuestFinFR(QUESTNUM))
1300             Call Manager.ChangeValue("QUEST" & QUESTNUM, "it_DescFinal", QuestFinIT(QUESTNUM))
1302             Call Manager.ChangeValue("QUEST" & QUESTNUM, "NextQuest", QuestNext(QUESTNUM))
1304             Call Manager.ChangeValue("QUEST" & QUESTNUM, "Repetible", QuestRepetible(QUESTNUM))
1306             Call Manager.ChangeValue("QUEST" & QUESTNUM, "RequiredLevel", RequiredLevel(QUESTNUM))
1308             Call Manager.ChangeValue("QUEST" & QUESTNUM, "PosMap", QuestPos(QUESTNUM))
1310             Label3.Caption = "Grabando Quest: " & QUESTNUM & "/" & nunquest
1312             Label3.ForeColor = vbGreen
1314         Next QUESTNUM

         Else
1316         MsgBox "Falta el archivo Quests.DAT dentro de la carpeta dats."

         End If


Dim idiomas() As String
Dim prefijos() As String
Dim secciones() As String
Dim iIdioma As Integer

idiomas = Split("SP,EN,PT,FR,IT", ",")
prefijos = Split("sp,en,pt,fr,it", ",")
secciones = Split("SP_SUGERENCIAS,EN_SUGERENCIAS,PT_SUGERENCIAS,FR_SUGERENCIAS,IT_SUGERENCIAS", ",")

For iIdioma = 0 To UBound(idiomas)
    Dim SugFile As String
    Dim SugerenciasReader As New clsIniReader
    Dim NumSugs As Integer
    Dim j As Integer

    SugFile = App.Path & "\..\Recursos\init\" & prefijos(iIdioma) & "_sugerencias.ini"
    
    If FileExist(SugFile, vbNormal) Then
        Call SugerenciasReader.Initialize(SugFile)
        NumSugs = Val(SugerenciasReader.GetValue(secciones(iIdioma), "NumSugerencias"))
        
        ' Solo actualizamos NumSugerencias en INIT si es español
        If idiomas(iIdioma) = "SP" Then
            Call Manager.ChangeValue("INIT", "NumSugerencias", NumSugs)
        End If
        
        For j = 1 To NumSugs
            DoEvents
            Call Manager.ChangeValue(secciones(iIdioma), "Sugerencia" & j, SugerenciasReader.GetValue(secciones(iIdioma), "Sugerencia" & j))
            Label3.Caption = "Grabando " & idiomas(iIdioma) & ": " & j & "/" & NumSugs
            Label3.ForeColor = vbGreen
        Next j
    Else
        MsgBox "Falta el archivo " & prefijos(iIdioma) & "_sugerencias.ini dentro de la carpeta init."
    End If
Next iIdioma



         Dim ListaRazas(1 To NUMRAZAS) As String
1360     ListaRazas(1) = "Humano"
1362     ListaRazas(2) = "Elfo"
1364     ListaRazas(3) = "Elfo Oscuro"
1366     ListaRazas(4) = "Gnomo"
1368     ListaRazas(5) = "Enano"
1370     ListaRazas(6) = "Orco"
1372     Call Leer.Initialize(App.Path & "\..\Recursos\Dat\Balance.dat")
         Dim SearchVar As String

1374     For Raza = 1 To NUMRAZAS

1376         With ModRaza(Raza)
1378             SearchVar = Replace(ListaRazas(Raza), " ", vbNullString)
1380             .Fuerza = Val(Leer.GetValue("MODRAZA", SearchVar + "Fuerza"))
1382             .Agilidad = Val(Leer.GetValue("MODRAZA", SearchVar + "Agilidad"))
1384             .Inteligencia = Val(Leer.GetValue("MODRAZA", SearchVar + "Inteligencia"))
1386             .Constitucion = Val(Leer.GetValue("MODRAZA", SearchVar + "Constitucion"))
1388             .Carisma = Val(Leer.GetValue("MODRAZA", SearchVar + "Carisma"))
1390             Call Manager.ChangeValue("MODRAZA", SearchVar + "Fuerza", .Fuerza)
1392             Call Manager.ChangeValue("MODRAZA", SearchVar + "Agilidad", .Agilidad)
1394             Call Manager.ChangeValue("MODRAZA", SearchVar + "Inteligencia", .Inteligencia)
1396             Call Manager.ChangeValue("MODRAZA", SearchVar + "Constitucion", .Constitucion)
1398             Call Manager.ChangeValue("MODRAZA", SearchVar + "Carisma", .Carisma)

             End With

1400     Next Raza

1402     Set Leer = Nothing
1404     Call Manager.DumpFile(OutputFile)
1406     Set Manager = Nothing
1408     Label3.ForeColor = vbGreen
1410     Label3.Caption = "Creado localindex.dat"
         Dim origen As New clsIniReader
1412     Call origen.Initialize(App.Path & "\..\Recursos\Init\localindex.dat")
1414     Call DumpLocalIndexPorIdioma(origen)

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
