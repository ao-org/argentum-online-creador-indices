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
122             ObjData(Obj).Name = Leer.GetValue("OBJ" & Obj, "Name")
124             ObjData(Obj).en_Name = Leer.GetValue("OBJ" & Obj, "en_Name")
126             ObjData(Obj).texto = Leer.GetValue("OBJ" & Obj, "Texto")
128             ObjData(Obj).en_texto = Leer.GetValue("OBJ" & Obj, "en_Texto")
130             ObjData(Obj).Info = Leer.GetValue("OBJ" & Obj, "Info")
132             ObjData(Obj).en_Info = Leer.GetValue("OBJ" & Obj, "en_Info")
134             ObjData(Obj).MINDEF = Val(Leer.GetValue("OBJ" & Obj, "MinDef"))
136             ObjData(Obj).MaxDEF = Val(Leer.GetValue("OBJ" & Obj, "MaxDef"))
138             ObjData(Obj).MinHit = Val(Leer.GetValue("OBJ" & Obj, "MinHit"))
140             ObjData(Obj).MaxHit = Val(Leer.GetValue("OBJ" & Obj, "MaxHit"))
142             ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
144             ObjData(Obj).CreaGRH = Leer.GetValue("OBJ" & Obj, "CreaGRH")
146             ObjData(Obj).CreaLuz = Leer.GetValue("OBJ" & Obj, "CreaLuz")
148             ObjData(Obj).CreaParticulaPiso = Val(Leer.GetValue("OBJ" & Obj, "CreaParticulaPiso"))
150             ObjData(Obj).Proyectil = Val(Leer.GetValue("OBJ" & Obj, "Proyectil"))
152             ObjData(Obj).Hechizo = Val(Leer.GetValue("OBJ" & Obj, "Hechizo"))
154             ObjData(Obj).Raices = Val(Leer.GetValue("OBJ" & Obj, "Raices"))
156             ObjData(Obj).Cuchara = Val(Leer.GetValue("OBJ" & Obj, "Cuchara"))
158             ObjData(Obj).Botella = Val(Leer.GetValue("OBJ" & Obj, "Botella"))
160             ObjData(Obj).Mortero = Val(Leer.GetValue("OBJ" & Obj, "Mortero"))
162             ObjData(Obj).FrascoAlq = Val(Leer.GetValue("OBJ" & Obj, "FrascoAlq"))
164             ObjData(Obj).FrascoElixir = Val(Leer.GetValue("OBJ" & Obj, "FrascoElixir"))
166             ObjData(Obj).Dosificador = Val(Leer.GetValue("OBJ" & Obj, "Dosificador"))
168             ObjData(Obj).Orquidea = Val(Leer.GetValue("OBJ" & Obj, "Orquidea"))
170             ObjData(Obj).Carmesi = Val(Leer.GetValue("OBJ" & Obj, "Carmesi"))
172             ObjData(Obj).HongoDeLuz = Val(Leer.GetValue("OBJ" & Obj, "HongoDeLuz"))
174             ObjData(Obj).Esporas = Val(Leer.GetValue("OBJ" & Obj, "Esporas"))
176             ObjData(Obj).Tuna = Val(Leer.GetValue("OBJ" & Obj, "Tuna"))
178             ObjData(Obj).Cala = Val(Leer.GetValue("OBJ" & Obj, "Cala"))
180             ObjData(Obj).ColaDeZorro = Val(Leer.GetValue("OBJ" & Obj, "ColaDeZorro"))
182             ObjData(Obj).FlorOceano = Val(Leer.GetValue("OBJ" & Obj, "FlorOceano"))
184             ObjData(Obj).FlorRoja = Val(Leer.GetValue("OBJ" & Obj, "FlorRoja"))
186             ObjData(Obj).Hierva = Val(Leer.GetValue("OBJ" & Obj, "Hierva"))
188             ObjData(Obj).HojasDeRin = Val(Leer.GetValue("OBJ" & Obj, "HojasDeRin"))
190             ObjData(Obj).HojasRojas = Val(Leer.GetValue("OBJ" & Obj, "HojasRojas"))
192             ObjData(Obj).SemillasPros = Val(Leer.GetValue("OBJ" & Obj, "SemillasPros"))
194             ObjData(Obj).Pimiento = Val(Leer.GetValue("OBJ" & Obj, "Pimiento"))
196             ObjData(Obj).Madera = Val(Leer.GetValue("OBJ" & Obj, "Madera"))
198             ObjData(Obj).MaderaElfica = Val(Leer.GetValue("OBJ" & Obj, "MaderaElfica"))
200             ObjData(Obj).PielLobo = Val(Leer.GetValue("OBJ" & Obj, "PielLobo"))
202             ObjData(Obj).PielLoboNegro = Val(Leer.GetValue("OBJ" & Obj, "PielLoboNegro"))
203             ObjData(Obj).PielTigre = Val(Leer.GetValue("OBJ" & Obj, "PielTigre"))
204             ObjData(Obj).PielOsoPardo = Val(Leer.GetValue("OBJ" & Obj, "PielOsoPardo"))
205             ObjData(Obj).PielTigreBengala = Val(Leer.GetValue("OBJ" & Obj, "PielTigreBengala"))
206             ObjData(Obj).PielOsoPolar = Val(Leer.GetValue("OBJ" & Obj, "PielOsoPolar"))
208             ObjData(Obj).LingH = Val(Leer.GetValue("OBJ" & Obj, "LingH"))
210             ObjData(Obj).LingP = Val(Leer.GetValue("OBJ" & Obj, "LingP"))
212             ObjData(Obj).LingO = Val(Leer.GetValue("OBJ" & Obj, "LingO"))
214             ObjData(Obj).Coal = Val(Leer.GetValue("OBJ" & Obj, "Coal"))
216             ObjData(Obj).Destruye = Val(Leer.GetValue("OBJ" & Obj, "Destruye"))
218             ObjData(Obj).SkHerreria = Val(Leer.GetValue("OBJ" & Obj, "SkHerreria"))
220             ObjData(Obj).SkPociones = Val(Leer.GetValue("OBJ" & Obj, "SkPociones"))
222             ObjData(Obj).Sksastreria = Val(Leer.GetValue("OBJ" & Obj, "Sksastreria"))
224             ObjData(Obj).Valor = Val(Leer.GetValue("OBJ" & Obj, "Valor"))
226             ObjData(Obj).Agarrable = Val(Leer.GetValue("OBJ" & Obj, "Agarrable"))
228             ObjData(Obj).Llave = Val(Leer.GetValue("OBJ" & Obj, "Llave"))
230             ObjData(Obj).Municiones = Val(Leer.GetValue("OBJ" & Obj, "Municiones"))
232             ObjData(Obj).Cooldown = Val(Leer.GetValue("OBJ" & Obj, "CD"))
234             ObjData(Obj).CdType = Val(Leer.GetValue("OBJ" & Obj, "CDType"))
236             ObjData(Obj).SpellIndex = Val(Leer.GetValue("OBJ" & Obj, "HechizoIndex"))
238             Label3.ForeColor = vbRed
240             Label3.Caption = "Leyendo objetos: " & Obj & "/" & numobjs
242         Next Obj

244         Obj = 1

            Dim Manager As clsIniReader

246         Set Manager = New clsIniReader
248         Call Manager.Initialize(OutputFile)
250         Call Manager.ChangeValue("INIT", "NumOBJs", numobjs)

252         For Obj = 1 To numobjs
254             DoEvents
256             Call Manager.ChangeValue("OBJ" & Obj, "GrhIndex", ObjData(Obj).grhindex)

258             If Len(ObjData(Obj).Name) <> 0 Then
260                 Call Manager.ChangeValue("OBJ" & Obj, "Name", ObjData(Obj).Name)

                End If

262             If Len(ObjData(Obj).texto) <> 0 Then
264                 Call Manager.ChangeValue("OBJ" & Obj, "Texto", ObjData(Obj).texto)

                End If

266             If Len(ObjData(Obj).Info) <> 0 Then
268                 Call Manager.ChangeValue("OBJ" & Obj, "Info", ObjData(Obj).Info)

                End If

                'English
270             If Len(ObjData(Obj).en_Name) <> 0 Then
272                 Call Manager.ChangeValue("OBJ" & Obj, "en_Name", ObjData(Obj).en_Name)

                End If

274             If Len(ObjData(Obj).en_texto) <> 0 Then
276                 Call Manager.ChangeValue("OBJ" & Obj, "en_Texto", ObjData(Obj).en_texto)

                End If

278             If Len(ObjData(Obj).en_Info) <> 0 Then
280                 Call Manager.ChangeValue("OBJ" & Obj, "en_Info", ObjData(Obj).en_Info)

                End If

282             If ObjData(Obj).MINDEF > 0 Then
284                 Call Manager.ChangeValue("OBJ" & Obj, "MINDEF", ObjData(Obj).MINDEF)

                End If

286             If ObjData(Obj).MaxDEF > 0 Then
288                 Call Manager.ChangeValue("OBJ" & Obj, "MaxDEF", ObjData(Obj).MaxDEF)

                End If

290             If ObjData(Obj).MinHit > 0 Then
292                 Call Manager.ChangeValue("OBJ" & Obj, "MinHIt", ObjData(Obj).MinHit)

                End If

294             If ObjData(Obj).MaxHit > 0 Then
296                 Call Manager.ChangeValue("OBJ" & Obj, "maxhit", ObjData(Obj).MaxHit)

                End If

298             If ObjData(Obj).ObjType > 0 Then
300                 Call Manager.ChangeValue("OBJ" & Obj, "ObjType", ObjData(Obj).ObjType)

                End If

302             If Len(ObjData(Obj).CreaLuz) <> 0 Then
304                 Call Manager.ChangeValue("OBJ" & Obj, "CreaLuz", ObjData(Obj).CreaLuz)

                End If

306             If Len(ObjData(Obj).CreaGRH) <> 0 Then
308                 Call Manager.ChangeValue("OBJ" & Obj, "CreaGRH", ObjData(Obj).CreaGRH)

                End If

310             If ObjData(Obj).Hechizo <> 0 Then
312                 Call Manager.ChangeValue("OBJ" & Obj, "Hechizo", ObjData(Obj).Hechizo)

                End If

314             If ObjData(Obj).Raices <> 0 Then
316                 Call Manager.ChangeValue("OBJ" & Obj, "Raices", ObjData(Obj).Raices)

                End If

318             If ObjData(Obj).Cuchara <> 0 Then
320                 Call Manager.ChangeValue("OBJ" & Obj, "Cuchara", ObjData(Obj).Cuchara)

                End If

322             If ObjData(Obj).Botella <> 0 Then
324                 Call Manager.ChangeValue("OBJ" & Obj, "Botella", ObjData(Obj).Botella)

                End If

326             If ObjData(Obj).Mortero <> 0 Then
328                 Call Manager.ChangeValue("OBJ" & Obj, "Mortero", ObjData(Obj).Mortero)

                End If

330             If ObjData(Obj).FrascoAlq <> 0 Then
332                 Call Manager.ChangeValue("OBJ" & Obj, "FrascoAlq", ObjData(Obj).FrascoAlq)

                End If

334             If ObjData(Obj).FrascoElixir <> 0 Then
336                 Call Manager.ChangeValue("OBJ" & Obj, "FrascoElixir", ObjData(Obj).FrascoElixir)

                End If

338             If ObjData(Obj).Dosificador <> 0 Then
340                 Call Manager.ChangeValue("OBJ" & Obj, "Dosificador", ObjData(Obj).Dosificador)

                End If

342             If ObjData(Obj).Orquidea <> 0 Then
344                 Call Manager.ChangeValue("OBJ" & Obj, "Orquidea", ObjData(Obj).Orquidea)

                End If

346             If ObjData(Obj).Carmesi <> 0 Then
348                 Call Manager.ChangeValue("OBJ" & Obj, "Carmesi", ObjData(Obj).Carmesi)

                End If

350             If ObjData(Obj).HongoDeLuz <> 0 Then
352                 Call Manager.ChangeValue("OBJ" & Obj, "HongoDeLuz", ObjData(Obj).HongoDeLuz)

                End If

354             If ObjData(Obj).Esporas <> 0 Then
356                 Call Manager.ChangeValue("OBJ" & Obj, "Esporas", ObjData(Obj).Esporas)

                End If

358             If ObjData(Obj).Tuna <> 0 Then
360                 Call Manager.ChangeValue("OBJ" & Obj, "Tuna", ObjData(Obj).Tuna)

                End If

362             If ObjData(Obj).Cala <> 0 Then
364                 Call Manager.ChangeValue("OBJ" & Obj, "Cala", ObjData(Obj).Cala)

                End If

366             If ObjData(Obj).ColaDeZorro <> 0 Then
368                 Call Manager.ChangeValue("OBJ" & Obj, "ColaDeZorro", ObjData(Obj).ColaDeZorro)

                End If

370             If ObjData(Obj).FlorOceano <> 0 Then
372                 Call Manager.ChangeValue("OBJ" & Obj, "FlorOceano", ObjData(Obj).FlorOceano)

                End If

374             If ObjData(Obj).FlorRoja <> 0 Then
376                 Call Manager.ChangeValue("OBJ" & Obj, "FlorRoja", ObjData(Obj).FlorRoja)

                End If

378             If ObjData(Obj).Hierva <> 0 Then
380                 Call Manager.ChangeValue("OBJ" & Obj, "Hierva", ObjData(Obj).Hierva)

                End If

382             If ObjData(Obj).HojasDeRin <> 0 Then
384                 Call Manager.ChangeValue("OBJ" & Obj, "HojasDeRin", ObjData(Obj).HojasDeRin)

                End If

386             If ObjData(Obj).HojasRojas <> 0 Then
388                 Call Manager.ChangeValue("OBJ" & Obj, "HojasRojas", ObjData(Obj).HojasRojas)

                End If

390             If ObjData(Obj).SemillasPros <> 0 Then
392                 Call Manager.ChangeValue("OBJ" & Obj, "SemillasPros", ObjData(Obj).SemillasPros)

                End If

394             If ObjData(Obj).Pimiento <> 0 Then
396                 Call Manager.ChangeValue("OBJ" & Obj, "Pimiento", ObjData(Obj).Pimiento)

                End If

398             If ObjData(Obj).Madera <> 0 Then
400                 Call Manager.ChangeValue("OBJ" & Obj, "Madera", ObjData(Obj).Madera)

                End If

402             If ObjData(Obj).MaderaElfica <> 0 Then
404                 Call Manager.ChangeValue("OBJ" & Obj, "MaderaElfica", ObjData(Obj).MaderaElfica)

                End If

406             If ObjData(Obj).PielLobo <> 0 Then
408                 Call Manager.ChangeValue("OBJ" & Obj, "PielLobo", ObjData(Obj).PielLobo)

                End If

410             If ObjData(Obj).PielLoboNegro <> 0 Then
412                 Call Manager.ChangeValue("OBJ" & Obj, "PielLoboNegro", ObjData(Obj).PielLoboNegro)

                End If

414             If ObjData(Obj).PielOsoPardo <> 0 Then
416                 Call Manager.ChangeValue("OBJ" & Obj, "PielOsoPardo", ObjData(Obj).PielOsoPardo)

                End If
                
                
                If ObjData(Obj).PielTigre <> 0 Then
                 Call Manager.ChangeValue("OBJ" & Obj, "PielTigre", ObjData(Obj).PielTigre)

                End If
                
                
         
                If ObjData(Obj).PielTigreBengala <> 0 Then
                 Call Manager.ChangeValue("OBJ" & Obj, "PielTigreBengala", ObjData(Obj).PielTigreBengala)

                End If
                

418             If ObjData(Obj).PielOsoPolar <> 0 Then
420                 Call Manager.ChangeValue("OBJ" & Obj, "PielOsoPolar", ObjData(Obj).PielOsoPolar)

                End If

422             If ObjData(Obj).LingH <> 0 Then
424                 Call Manager.ChangeValue("OBJ" & Obj, "LingH", ObjData(Obj).LingH)

                End If

426             If ObjData(Obj).LingP <> 0 Then
428                 Call Manager.ChangeValue("OBJ" & Obj, "LingP", ObjData(Obj).LingP)

                End If

430             If ObjData(Obj).LingO <> 0 Then
432                 Call Manager.ChangeValue("OBJ" & Obj, "LingO", ObjData(Obj).LingO)

                End If

434             If ObjData(Obj).Coal <> 0 Then
436                 Call Manager.ChangeValue("OBJ" & Obj, "Coal", ObjData(Obj).Coal)

                End If

438             If ObjData(Obj).Destruye <> 0 Then
440                 Call Manager.ChangeValue("OBJ" & Obj, "Destruye", ObjData(Obj).Destruye)

                End If

442             If ObjData(Obj).SkHerreria <> 0 Then
444                 Call Manager.ChangeValue("OBJ" & Obj, "SkHerreria", ObjData(Obj).SkHerreria)

                End If

446             If ObjData(Obj).SkPociones <> 0 Then
448                 Call Manager.ChangeValue("OBJ" & Obj, "SkPociones", ObjData(Obj).SkPociones)

                End If

450             If ObjData(Obj).Sksastreria <> 0 Then
452                 Call Manager.ChangeValue("OBJ" & Obj, "Sksastreria", ObjData(Obj).Sksastreria)

                End If

454             If ObjData(Obj).Valor <> 0 Then
456                 Call Manager.ChangeValue("OBJ" & Obj, "Valor", ObjData(Obj).Valor)

                End If

458             If ObjData(Obj).Agarrable Then
460                 Call Manager.ChangeValue("OBJ" & Obj, "Agarrable", 1)

                End If

462             If ObjData(Obj).CreaParticulaPiso > 0 Then
464                 Call Manager.ChangeValue("OBJ" & Obj, "CreaParticulaPiso", ObjData(Obj).CreaParticulaPiso)

                End If

466             If ObjData(Obj).Proyectil > 0 Then
468                 Call Manager.ChangeValue("OBJ" & Obj, "Proyectil", ObjData(Obj).Proyectil)

                End If

470             If ObjData(Obj).Municiones > 0 Then
472                 Call Manager.ChangeValue("OBJ" & Obj, "Municiones", ObjData(Obj).Municiones)

                End If

474             If ObjData(Obj).Llave > 0 Then
476                 Call Manager.ChangeValue("OBJ" & Obj, "Llave", ObjData(Obj).Llave)

                End If

478             If ObjData(Obj).Cooldown > 0 Then
480                 Call Manager.ChangeValue("OBJ" & Obj, "CD", ObjData(Obj).Cooldown)

                End If

482             If ObjData(Obj).CdType > 0 Then
484                 Call Manager.ChangeValue("OBJ" & Obj, "CDType", ObjData(Obj).CdType)

                End If

486             If ObjData(Obj).SpellIndex > 0 Then
488                 Call Manager.ChangeValue("OBJ" & Obj, "SpellIndex", ObjData(Obj).SpellIndex)

                End If

490             Label3.Caption = "Grabando: " & Obj & "/" & numobjs
492             Label3.ForeColor = &HC0C0&
494         Next Obj

496         Label3.ForeColor = vbGreen
498         Label3.Caption = "Creado objindex.dat"
        Else
500         MsgBox "Falta el archivo obj.dat dentro de la carpeta INIT."

        End If

502     If FileExist(App.Path & "\..\Recursos\Dat\npcs.dat", vbNormal) Then
504         NpcFile = App.Path & "\..\Recursos\Dat\npcs.dat"
506         Call Leer.Initialize(NpcFile)

            Dim numnpcs As Long

508         numnpcs = Val(GetVar(NpcFile, "INIT", "NumNPCs"))
510         Label3.Caption = "0/" & numnpcs
512         ReDim NpcData(1 To numnpcs) As NpcDatas

            Dim aux As String

514         For Npc = 1 To numnpcs
516             DoEvents
518             NpcData(Npc).Name = Leer.GetValue("npc" & Npc, "Name")
520             NpcData(Npc).en_Name = Leer.GetValue("npc" & Npc, "en_Name")
522             NpcData(Npc).desc = Leer.GetValue("npc" & Npc, "desc")
524             NpcData(Npc).en_desc = Leer.GetValue("npc" & Npc, "en_desc")
526             NpcData(Npc).Body = Val(Leer.GetValue("npc" & Npc, "Body"))
528             NpcData(Npc).Exp = Val(Leer.GetValue("npc" & Npc, "GiveEXP"))
530             NpcData(Npc).Head = Val(Leer.GetValue("npc" & Npc, "Head"))
532             NpcData(Npc).Hp = Val(Leer.GetValue("npc" & Npc, "MaxHP"))
534             NpcData(Npc).MaxHit = Val(Leer.GetValue("npc" & Npc, "MaxHit"))
536             NpcData(Npc).MinHit = Val(Leer.GetValue("npc" & Npc, "MinHit"))
538             NpcData(Npc).Oro = Val(Leer.GetValue("npc" & Npc, "GiveGLD"))
540             NpcData(Npc).ExpClan = Val(Leer.GetValue("npc" & Npc, "GiveEXPClan"))
542             NpcData(Npc).PuedeInvocar = Val(Leer.GetValue("npc" & Npc, "PuedeInvocar"))
                NpcData(Npc).NoMapInfo = Val(Leer.GetValue("npc" & Npc, "NoMapInfo"))
544             NpcData(Npc).QuizaProb = Val(Leer.GetValue("npc" & Npc, "QuizaProb"))
546             aux = Val(GetVar(NpcFile, "Npc" & Npc, "NumQuiza"))

548             If aux = 0 Then
550                 NpcData(Npc).NumQuiza = 0
                Else
552                 NpcData(Npc).NumQuiza = Val(aux)
554                 ReDim NpcData(Npc).QuizaDropea(1 To NpcData(Npc).NumQuiza) As Integer

                    Dim LoopC As Long

556                 For LoopC = 1 To NpcData(Npc).NumQuiza
558                     NpcData(Npc).QuizaDropea(LoopC) = Val(Leer.GetValue("npc" & Npc, "QuizaDropea" & LoopC))
560                 Next LoopC

                End If

562             Label3.ForeColor = vbRed
564             Label3.Caption = "Leyendo NPCs: " & Npc & "/" & numnpcs
566         Next Npc

568         Npc = 1
570         Call Manager.ChangeValue("INIT", "NumNPCs", numnpcs)

572         For Npc = 1 To numnpcs
574             DoEvents

576             If Len(NpcData(Npc).Name) <> 0 Then
578                 Call Manager.ChangeValue("Npc" & Npc, "Name", NpcData(Npc).Name)

                End If

580             If Len(NpcData(Npc).en_Name) <> 0 Then
582                 Call Manager.ChangeValue("Npc" & Npc, "en_Name", NpcData(Npc).en_Name)

                End If

584             If Len(NpcData(Npc).en_desc) <> 0 Then
586                 Call Manager.ChangeValue("Npc" & Npc, "en_desc", NpcData(Npc).en_desc)

                End If

588             If Len(NpcData(Npc).desc) <> 0 Then
590                 Call Manager.ChangeValue("Npc" & Npc, "Desc", NpcData(Npc).desc)

                End If

592             If NpcData(Npc).Body <> 0 Then
594                 Call Manager.ChangeValue("Npc" & Npc, "Body", NpcData(Npc).Body)

                End If

596             If NpcData(Npc).Head <> 0 Then
598                 Call Manager.ChangeValue("Npc" & Npc, "Head", NpcData(Npc).Head)

                End If

600             If NpcData(Npc).Exp <> 0 Then
602                 Call Manager.ChangeValue("Npc" & Npc, "Exp", NpcData(Npc).Exp)

                End If

604             If NpcData(Npc).Hp <> 0 Then
606                 Call Manager.ChangeValue("Npc" & Npc, "Hp", NpcData(Npc).Hp)

                End If

608             If NpcData(Npc).MaxHit <> 0 Then
610                 Call Manager.ChangeValue("Npc" & Npc, "MaxHit", NpcData(Npc).MaxHit)

                End If

612             If NpcData(Npc).MinHit <> 0 Then
614                 Call Manager.ChangeValue("Npc" & Npc, "MinHit", NpcData(Npc).MinHit)

                End If

616             If NpcData(Npc).Oro <> 0 Then
618                 Call Manager.ChangeValue("Npc" & Npc, "Oro", NpcData(Npc).Oro)

                End If

620             If NpcData(Npc).ExpClan <> 0 Then
622                 Call Manager.ChangeValue("Npc" & Npc, "GiveEXPClan", NpcData(Npc).ExpClan)

                End If

624             If NpcData(Npc).NumQuiza <> 0 Then
626                 Call Manager.ChangeValue("Npc" & Npc, "NumQuiza", NpcData(Npc).NumQuiza)

628                 For LoopC = 1 To NpcData(Npc).NumQuiza
630                     Call Manager.ChangeValue("Npc" & Npc, "QuizaDropea" & LoopC, NpcData(Npc).QuizaDropea(LoopC))
632                 Next LoopC

                End If

634             If NpcData(Npc).QuizaProb <> 0 Then
636                 Call Manager.ChangeValue("Npc" & Npc, "QuizaProb", NpcData(Npc).QuizaProb)
                End If
                
                If NpcData(Npc).NoMapInfo <> 0 Then
                     Call Manager.ChangeValue("Npc" & Npc, "NoMapInfo", NpcData(Npc).NoMapInfo)
                End If

638             If NpcData(Npc).PuedeInvocar <> 0 Then
640                 Call Manager.ChangeValue("Npc" & Npc, "PuedeInvocar", NpcData(Npc).PuedeInvocar)

                End If

642             Label3.Caption = "Grabando NPCs: " & Npc & "/" & numnpcs
644             Label3.ForeColor = &HC0C0&
646         Next Npc

        Else
648         MsgBox "Falta el archivo npcs.dat dentro de la carpeta dats."

        End If

650     If FileExist(App.Path & "\..\Recursos\Dat\hechizos.dat", vbNormal) Then

            Dim hechizosFile As String, numhechizos As Long

652         hechizosFile = App.Path & "\..\Recursos\Dat\hechizos.dat"
654         numhechizos = Val(GetVar(hechizosFile, "INIT", "NumeroHechizos"))

            Dim hechic As New clsIniReader

656         Call hechic.Initialize(hechizosFile)
658         Label3.Caption = "Leyendo Hechizos: " & "0/" & numhechizos
660         ReDim HechizoData(1 To numhechizos) As HechizoDatas

662         For Hechizo = 1 To numhechizos
664             DoEvents
666             HechizoData(Hechizo).Nombre = hechic.GetValue("Hechizo" & Hechizo, "Nombre")
668             HechizoData(Hechizo).desc = hechic.GetValue("Hechizo" & Hechizo, "desc")
670             HechizoData(Hechizo).PalabrasMagicas = hechic.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
672             HechizoData(Hechizo).HechizeroMsg = hechic.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
674             HechizoData(Hechizo).TargetMsg = hechic.GetValue("Hechizo" & Hechizo, "TargetMsg")
676             HechizoData(Hechizo).PropioMsg = hechic.GetValue("Hechizo" & Hechizo, "PropioMsg")
678             HechizoData(Hechizo).ManaRequerido = Val(hechic.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
680             HechizoData(Hechizo).StaRequerido = Val(hechic.GetValue("Hechizo" & Hechizo, "StaRequerido"))
682             HechizoData(Hechizo).MinSkill = Val(hechic.GetValue("Hechizo" & Hechizo, "MinSkill"))
684             HechizoData(Hechizo).StaRequerido = Val(hechic.GetValue("Hechizo" & Hechizo, "StaRequerido"))
686             HechizoData(Hechizo).IconoIndex = Val(hechic.GetValue("Hechizo" & Hechizo, "IconoIndex"))
688             HechizoData(Hechizo).Cooldown = Val(hechic.GetValue("Hechizo" & Hechizo, "Cooldown"))
690             Label3.ForeColor = vbRed
692             Label3.Caption = "Leyendo: " & Hechizo & "/" & numhechizos
694         Next Hechizo

696         Call Manager.ChangeValue("INIT", "NumeroHechizo", numhechizos)

698         For Hechizo = 1 To numhechizos
700             DoEvents
702             Call Manager.ChangeValue("Hechizo" & Hechizo, "Nombre", HechizoData(Hechizo).Nombre)
704             Call Manager.ChangeValue("Hechizo" & Hechizo, "Desc", HechizoData(Hechizo).desc)
706             Call Manager.ChangeValue("Hechizo" & Hechizo, "PalabrasMagicas", HechizoData(Hechizo).PalabrasMagicas)
708             Call Manager.ChangeValue("Hechizo" & Hechizo, "HechizeroMsg", HechizoData(Hechizo).HechizeroMsg)
710             Call Manager.ChangeValue("Hechizo" & Hechizo, "TargetMsg", HechizoData(Hechizo).TargetMsg)
712             Call Manager.ChangeValue("Hechizo" & Hechizo, "PropioMsg", HechizoData(Hechizo).PropioMsg)
714             Call Manager.ChangeValue("Hechizo" & Hechizo, "ManaRequerido", HechizoData(Hechizo).ManaRequerido)
716             Call Manager.ChangeValue("Hechizo" & Hechizo, "StaRequerido", HechizoData(Hechizo).StaRequerido)
718             Call Manager.ChangeValue("Hechizo" & Hechizo, "MinSkill", HechizoData(Hechizo).MinSkill)
720             Call Manager.ChangeValue("Hechizo" & Hechizo, "IconoIndex", HechizoData(Hechizo).IconoIndex)
722             Call Manager.ChangeValue("Hechizo" & Hechizo, "Cooldown", HechizoData(Hechizo).Cooldown)
724             Label3.Caption = "Grabando Hechizos: " & Hechizo & "/" & numhechizos
726             Label3.ForeColor = &HC0C0&
728         Next Hechizo

        End If

730     If FileExist(App.Path & "\..\Recursos\init\LocalMsg.dat", vbNormal) Then

            Dim MsgFile As String

732         MsgFile = App.Path & "\..\Recursos\init\LocalMsg.dat"

            Dim Msgsss As New clsIniReader

734         Call Msgsss.Initialize(MsgFile)
736         numnpcs = Val(Msgsss.GetValue("INIT", "NumLocaleMsg"))
738         Label3.Caption = "0/" & CStr(numnpcs)
740         ReDim arrLocale_SMG(1 To numnpcs) As String

742         For Npc = 1 To numnpcs
744             DoEvents
746             arrLocale_SMG(Npc) = Msgsss.GetValue("msg", "Msg" & Npc)
748             Label3.ForeColor = vbRed
750             Label3.Caption = "Leyendo NPCs: " & Npc & "/" & numnpcs
752         Next Npc

754         Npc = 1
756         Call Manager.ChangeValue("INIT", "NumLocaleMsg", numnpcs)

758         For Npc = 1 To numnpcs
760             DoEvents
762             Call Manager.ChangeValue("Msg", "Msg" & Npc, arrLocale_SMG(Npc))
764             Label3.Caption = "Grabando NPCs: " & Npc & "/" & numnpcs
766             Label3.ForeColor = &HC0C0&
768         Next Npc

        Else
770         MsgBox "Falta el archivo LocalMsg.dat dentro de la carpeta dats."

        End If

772     If FileExist(App.Path & "\..\Recursos\init\NameMapa.dat", vbNormal) Then

            Dim MapFile As String

774         MapFile = App.Path & "\..\Recursos\init\NameMapa.dat"

            Dim Mapa As New clsIniReader

776         Call Mapa.Initialize(MapFile)
778         Label3.Caption = "0/" & 750
780         ReDim MapName(1 To 750) As String
782         ReDim MapDesc(1 To 750) As String

784         For Npc = 1 To 750
786             DoEvents
788             MapName(Npc) = Mapa.GetValue("NameMapa", "mapa" & Npc)
790             MapDesc(Npc) = Mapa.GetValue("NameMapa", "mapa" & Npc & "desc")
792             Label3.ForeColor = vbRed
794             Label3.Caption = "Leyendo Mapas: " & Npc & "/" & 750
796         Next Npc

798         Npc = 1
800         Call Manager.ChangeValue("INIT", "NumMapas", 750)

802         For Npc = 1 To 750
804             DoEvents
806             Call Manager.ChangeValue("NAMEMAPA", "Mapa" & Npc, MapName(Npc))
808             Call Manager.ChangeValue("NAMEMAPA", "Mapa" & Npc & "Desc", MapDesc(Npc))
810             Label3.Caption = "Grabando Mapas: " & Npc & "/" & 750
812             Label3.ForeColor = &HC0C0&
814         Next Npc

        Else
816         MsgBox "Falta el archivo NameMapa.dat dentro de la carpeta dats."

        End If

        'quest
818     If FileExist(App.Path & "\..\Recursos\Dat\Quests.DAT", vbNormal) Then
820         MapFile = App.Path & "\..\Recursos\Dat\Quests.DAT"
822         Call Mapa.Initialize(MapFile)

            Dim nunquest As Integer

824         nunquest = Mapa.GetValue("INIT", "NumQuests")
826         Label3.Caption = "0/" & nunquest
828         ReDim QuestName(1 To nunquest) As String
830         ReDim QuestDesc(1 To nunquest) As String
832         ReDim QuestFin(1 To nunquest) As String
834         ReDim QuestNext(1 To nunquest) As String
836         ReDim QuestPos(1 To nunquest) As Integer
838         ReDim QuestRepetible(1 To nunquest) As Byte
840         ReDim RequiredLevel(1 To nunquest) As Integer

842         For Npc = 1 To nunquest
844             DoEvents
846             QuestName(Npc) = Mapa.GetValue("QUEST" & Npc, "Nombre")
848             QuestDesc(Npc) = Mapa.GetValue("QUEST" & Npc, "Desc")
850             QuestFin(Npc) = Mapa.GetValue("QUEST" & Npc, "DescFinal")
852             QuestNext(Npc) = Mapa.GetValue("QUEST" & Npc, "NextQuest")
854             QuestRepetible(Npc) = Val(Mapa.GetValue("QUEST" & Npc, "Repetible"))
856             QuestPos(Npc) = Val(Mapa.GetValue("QUEST" & Npc, "PosMap"))
858             RequiredLevel(Npc) = Val(Mapa.GetValue("QUEST" & Npc, "RequiredLevel"))
860             Label3.ForeColor = vbRed
862             Label3.Caption = "Leyendo Quest: " & Npc & "/" & nunquest
864         Next Npc

866         Npc = 1
868         Call Manager.ChangeValue("INIT", "NumQuests", nunquest)

870         For Npc = 1 To nunquest
872             DoEvents
874             Call Manager.ChangeValue("QUEST" & Npc, "Nombre", QuestName(Npc))
876             Call Manager.ChangeValue("QUEST" & Npc, "Desc", QuestDesc(Npc))
878             Call Manager.ChangeValue("QUEST" & Npc, "DescFinal", QuestFin(Npc))
880             Call Manager.ChangeValue("QUEST" & Npc, "NextQuest", QuestNext(Npc))
882             Call Manager.ChangeValue("QUEST" & Npc, "Repetible", QuestRepetible(Npc))
884             Call Manager.ChangeValue("QUEST" & Npc, "RequiredLevel", RequiredLevel(Npc))
886             Call Manager.ChangeValue("QUEST" & Npc, "PosMap", QuestPos(Npc))
888             Label3.Caption = "Grabando Quest: " & Npc & "/" & nunquest
890             Label3.ForeColor = &HC0C0&
892         Next Npc

        Else
894         MsgBox "Falta el archivo Quests.DAT dentro de la carpeta dats."

        End If

896     If FileExist(App.Path & "\..\Recursos\init\sugerencias.ini", vbNormal) Then
898         MapFile = App.Path & "\..\Recursos\init\sugerencias.ini"
900         Call Mapa.Initialize(MapFile)

            Dim NumSug As Integer

902         NumSug = Val(Mapa.GetValue("Sugerencias", "NumSugerencias"))
904         Label3.Caption = "0/" & CStr(NumSug)
906         ReDim Sugerencia(1 To NumSug) As String

908         For Npc = 1 To NumSug
910             DoEvents
912             Sugerencia(Npc) = Mapa.GetValue("Sugerencias", "Sugerencia" & Npc)
914             Label3.ForeColor = vbRed
916             Label3.Caption = "Leyendo: " & Npc & "/" & nunquest
918         Next Npc

920         Npc = 1
922         Call Manager.ChangeValue("INIT", "NumSugerencias", NumSug)

924         For Npc = 1 To NumSug
926             DoEvents
928             Call Manager.ChangeValue("Sugerencias", "Sugerencia" & Npc, Sugerencia(Npc))
930             Label3.Caption = "Grabando: " & Npc & "/" & NumSug
932             Label3.ForeColor = &HC0C0&
934         Next Npc

        Else
936         MsgBox "Falta el archivo Sugerencias.ini dentro de la carpeta init."

        End If

        Dim ListaRazas(1 To NUMRAZAS) As String

938     ListaRazas(1) = "Humano"
940     ListaRazas(2) = "Elfo"
942     ListaRazas(3) = "Elfo Oscuro"
944     ListaRazas(4) = "Gnomo"
946     ListaRazas(5) = "Enano"
948     ListaRazas(6) = "Orco"
950     Call Leer.Initialize(App.Path & "\..\Recursos\Dat\Balance.dat")

        Dim SearchVar As String

952     For Raza = 1 To NUMRAZAS

954         With ModRaza(Raza)
956             SearchVar = Replace(ListaRazas(Raza), " ", vbNullString)
958             .Fuerza = Val(Leer.GetValue("MODRAZA", SearchVar + "Fuerza"))
960             .Agilidad = Val(Leer.GetValue("MODRAZA", SearchVar + "Agilidad"))
962             .Inteligencia = Val(Leer.GetValue("MODRAZA", SearchVar + "Inteligencia"))
964             .Constitucion = Val(Leer.GetValue("MODRAZA", SearchVar + "Constitucion"))
966             .Carisma = Val(Leer.GetValue("MODRAZA", SearchVar + "Carisma"))
968             Call Manager.ChangeValue("MODRAZA", SearchVar + "Fuerza", .Fuerza)
970             Call Manager.ChangeValue("MODRAZA", SearchVar + "Agilidad", .Agilidad)
972             Call Manager.ChangeValue("MODRAZA", SearchVar + "Inteligencia", .Inteligencia)
974             Call Manager.ChangeValue("MODRAZA", SearchVar + "Constitucion", .Constitucion)
976             Call Manager.ChangeValue("MODRAZA", SearchVar + "Carisma", .Carisma)

            End With

978     Next Raza

980     Set Leer = Nothing
982     Call Manager.DumpFile(OutputFile)
984     Set Manager = Nothing
986     Label3.ForeColor = vbGreen
988     Label3.Caption = "Creado localindex.dat"

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
