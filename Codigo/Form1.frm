VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Creador de indices"
   ClientHeight    =   2355
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   2970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   2970
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
      Left            =   1560
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
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
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Preparado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Programado por Ladder"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
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
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2895
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
    Dim hechizo As Integer
    Dim Raza    As Integer
    Dim numobjs As Long

    If FileExist(OutputFile, vbNormal) Then
        Clean_File OutputFile
    End If

    If FileExist(App.Path & "\..\Recursos\Dat\obj.dat", vbNormal) Then

        ObjFile = App.Path & "\..\Recursos\Dat\obj.dat"
        numobjs = Val(GetVar(ObjFile, "INIT", "NumOBJs"))

        Label3.Caption = "0/" & numobjs

        ReDim ObjData(1 To numobjs) As ObjDatas

        Dim Leer As New clsIniReader
        Call Leer.Initialize(ObjFile)
    
        For Obj = 1 To numobjs
            DoEvents
            
            ObjData(Obj).grhindex = Val(Leer.GetValue("OBJ" & Obj, "grhindex"))
            ObjData(Obj).Name = Leer.GetValue("OBJ" & Obj, "Name")
            ObjData(Obj).en_Name = Leer.GetValue("OBJ" & Obj, "en_Name")
            ObjData(Obj).texto = Leer.GetValue("OBJ" & Obj, "Texto")
            ObjData(Obj).en_texto = Leer.GetValue("OBJ" & Obj, "en_Texto")
            ObjData(Obj).Info = Leer.GetValue("OBJ" & Obj, "Info")
            ObjData(Obj).en_Info = Leer.GetValue("OBJ" & Obj, "en_Info")
            ObjData(Obj).MINDEF = Val(Leer.GetValue("OBJ" & Obj, "MinDef"))
            ObjData(Obj).MaxDEF = Val(Leer.GetValue("OBJ" & Obj, "MaxDef"))
            ObjData(Obj).MinHit = Val(Leer.GetValue("OBJ" & Obj, "MinHit"))
            ObjData(Obj).MaxHit = Val(Leer.GetValue("OBJ" & Obj, "MaxHit"))
            ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
            ObjData(Obj).CreaGRH = Leer.GetValue("OBJ" & Obj, "CreaGRH")
            ObjData(Obj).CreaLuz = Leer.GetValue("OBJ" & Obj, "CreaLuz")
            ObjData(Obj).CreaParticulaPiso = Val(Leer.GetValue("OBJ" & Obj, "CreaParticulaPiso"))
            ObjData(Obj).Proyectil = Val(Leer.GetValue("OBJ" & Obj, "Proyectil"))
            ObjData(Obj).Raices = Val(Leer.GetValue("OBJ" & Obj, "Raices"))
            ObjData(Obj).Madera = Val(Leer.GetValue("OBJ" & Obj, "Madera"))
            ObjData(Obj).MaderaElfica = Val(Leer.GetValue("OBJ" & Obj, "MaderaElfica"))
            ObjData(Obj).PielLobo = Val(Leer.GetValue("OBJ" & Obj, "PielLobo"))
            ObjData(Obj).PielOsoPardo = Val(Leer.GetValue("OBJ" & Obj, "PielOsoPardo"))
            ObjData(Obj).PielOsoPolar = Val(Leer.GetValue("OBJ" & Obj, "PielOsoPolar"))
            ObjData(Obj).LingH = Val(Leer.GetValue("OBJ" & Obj, "LingH"))
            ObjData(Obj).LingP = Val(Leer.GetValue("OBJ" & Obj, "LingP"))
            ObjData(Obj).LingO = Val(Leer.GetValue("OBJ" & Obj, "LingO"))
            ObjData(Obj).Destruye = Val(Leer.GetValue("OBJ" & Obj, "Destruye"))
            ObjData(Obj).SkHerreria = Val(Leer.GetValue("OBJ" & Obj, "SkHerreria"))
            ObjData(Obj).SkPociones = Val(Leer.GetValue("OBJ" & Obj, "SkPociones"))
            ObjData(Obj).Sksastreria = Val(Leer.GetValue("OBJ" & Obj, "Sksastreria"))
            ObjData(Obj).Valor = Val(Leer.GetValue("OBJ" & Obj, "Valor"))
            ObjData(Obj).Agarrable = Val(Leer.GetValue("OBJ" & Obj, "Agarrable"))
            ObjData(Obj).Llave = Val(Leer.GetValue("OBJ" & Obj, "Llave"))
            ObjData(Obj).Municiones = Val(Leer.GetValue("OBJ" & Obj, "Municiones"))
            ObjData(Obj).Cooldown = Val(Leer.GetValue("OBJ" & Obj, "CD"))
            ObjData(Obj).CdType = Val(Leer.GetValue("OBJ" & Obj, "CDType"))
            
            Label3.ForeColor = vbRed
            Label3.Caption = "Leyendo: " & Obj & "/" & numobjs
        Next Obj
    
        Obj = 1

        Dim Manager  As clsIniReader
        Set Manager = New clsIniReader
        Call Manager.Initialize(OutputFile)
    
        Call Manager.ChangeValue("INIT", "NumOBJs", numobjs)

        For Obj = 1 To numobjs
            DoEvents
            
            Call Manager.ChangeValue("OBJ" & Obj, "GrhIndex", ObjData(Obj).grhindex)
            
            If Len(ObjData(Obj).Name) <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "Name", ObjData(Obj).Name)
            End If
            
            If Len(ObjData(Obj).texto) <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "Texto", ObjData(Obj).texto)
            End If
        
            If Len(ObjData(Obj).Info) <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "Info", ObjData(Obj).Info)
            End If
            
            'English
            If Len(ObjData(Obj).en_Name) <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "en_Name", ObjData(Obj).en_Name)
            End If
            
            If Len(ObjData(Obj).en_texto) <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "en_Texto", ObjData(Obj).en_texto)
            End If
        
            If Len(ObjData(Obj).en_Info) <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "en_Info", ObjData(Obj).en_Info)
            End If
        
            If ObjData(Obj).MINDEF > 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "MINDEF", ObjData(Obj).MINDEF)
            End If
        
            If ObjData(Obj).MaxDEF > 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "MaxDEF", ObjData(Obj).MaxDEF)
            End If
        
            If ObjData(Obj).MinHit > 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "MinHIt", ObjData(Obj).MinHit)
            End If
        
            If ObjData(Obj).MaxHit > 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "maxhit", ObjData(Obj).MaxHit)
            End If
        
            If ObjData(Obj).ObjType > 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "ObjType", ObjData(Obj).ObjType)
            End If
        
            If Len(ObjData(Obj).CreaLuz) <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "CreaLuz", ObjData(Obj).CreaLuz)
            End If

            If Len(ObjData(Obj).CreaGRH) <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "CreaGRH", ObjData(Obj).CreaGRH)
            End If
        
            If ObjData(Obj).Raices <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "Raices", ObjData(Obj).Raices)
            End If
        
            
            If ObjData(Obj).Madera <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "Madera", ObjData(Obj).Madera)
            End If
            
            If ObjData(Obj).MaderaElfica <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "MaderaElfica", ObjData(Obj).MaderaElfica)
            End If

            If ObjData(Obj).PielLobo <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "PielLobo", ObjData(Obj).PielLobo)
            End If

            If ObjData(Obj).PielOsoPardo <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "PielOsoPardo", ObjData(Obj).PielOsoPardo)
            End If

            If ObjData(Obj).PielOsoPolar <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "PielOsoPolar", ObjData(Obj).PielOsoPolar)
            End If

            If ObjData(Obj).LingH <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "LingH", ObjData(Obj).LingH)
            End If

            If ObjData(Obj).LingP <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "LingP", ObjData(Obj).LingP)
            End If

            If ObjData(Obj).LingO <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "LingO", ObjData(Obj).LingO)
            End If
        
            If ObjData(Obj).Destruye <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "Destruye", ObjData(Obj).Destruye)
            End If
        
            If ObjData(Obj).SkHerreria <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "SkHerreria", ObjData(Obj).SkHerreria)
            End If
        
            If ObjData(Obj).SkPociones <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "SkPociones", ObjData(Obj).SkPociones)
            End If
        
            If ObjData(Obj).Sksastreria <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "Sksastreria", ObjData(Obj).Sksastreria)
            End If
        
            If ObjData(Obj).Valor <> 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "Valor", ObjData(Obj).Valor)
            End If
            
            If ObjData(Obj).Agarrable Then
                Call Manager.ChangeValue("OBJ" & Obj, "Agarrable", 1)
            End If
        
            If ObjData(Obj).CreaParticulaPiso > 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "CreaParticulaPiso", ObjData(Obj).CreaParticulaPiso)
            End If
        
            If ObjData(Obj).Proyectil > 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "Proyectil", ObjData(Obj).Proyectil)
            End If
            
            If ObjData(Obj).Municiones > 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "Municiones", ObjData(Obj).Municiones)
            End If
        
            If ObjData(Obj).Llave > 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "Llave", ObjData(Obj).Llave)
            End If
            
            If ObjData(Obj).Cooldown > 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "CD", ObjData(Obj).Cooldown)
            End If
            If ObjData(Obj).CdType > 0 Then
                Call Manager.ChangeValue("OBJ" & Obj, "CDType", ObjData(Obj).CdType)
            End If

            Label3.Caption = "Grabando: " & Obj & "/" & numobjs
            Label3.ForeColor = &HC0C0&
        Next Obj

        Label3.ForeColor = vbGreen
        Label3.Caption = "Creado objindex.dat"
    Else
        MsgBox "Falta el archivo obj.dat dentro de la carpeta INIT."

    End If
    
    If FileExist(App.Path & "\..\Recursos\Dat\npcs.dat", vbNormal) Then

        NpcFile = App.Path & "\..\Recursos\Dat\npcs.dat"

        Call Leer.Initialize(NpcFile)

        Dim numnpcs As Long
        numnpcs = Val(GetVar(NpcFile, "INIT", "NumNPCs"))

        Label3.Caption = "0/" & numnpcs

        ReDim NpcData(1 To numnpcs) As NpcDatas
    
        Dim aux As String

        For Npc = 1 To numnpcs
            DoEvents
            
            NpcData(Npc).Name = Leer.GetValue("npc" & Npc, "Name")
            NpcData(Npc).en_Name = Leer.GetValue("npc" & Npc, "en_Name")
            NpcData(Npc).desc = Leer.GetValue("npc" & Npc, "desc")
            NpcData(Npc).en_desc = Leer.GetValue("npc" & Npc, "en_desc")
            NpcData(Npc).Body = Val(Leer.GetValue("npc" & Npc, "Body"))
            NpcData(Npc).Exp = Val(Leer.GetValue("npc" & Npc, "GiveEXP"))
            NpcData(Npc).Head = Val(Leer.GetValue("npc" & Npc, "Head"))
            NpcData(Npc).Hp = Val(Leer.GetValue("npc" & Npc, "MaxHP"))
            NpcData(Npc).MaxHit = Val(Leer.GetValue("npc" & Npc, "MaxHit"))
            NpcData(Npc).MinHit = Val(Leer.GetValue("npc" & Npc, "MinHit"))
            NpcData(Npc).Oro = Val(Leer.GetValue("npc" & Npc, "GiveGLD"))
            NpcData(Npc).ExpClan = Val(Leer.GetValue("npc" & Npc, "GiveEXPClan"))
            NpcData(Npc).PuedeInvocar = Val(Leer.GetValue("npc" & Npc, "PuedeInvocar"))
        
            aux = Val(GetVar(NpcFile, "Npc" & Npc, "NumQuiza"))

            If aux = 0 Then
                NpcData(Npc).NumQuiza = 0
                
            Else
            
                NpcData(Npc).NumQuiza = Val(aux)
                ReDim NpcData(Npc).QuizaDropea(1 To NpcData(Npc).NumQuiza) As Integer
                
                Dim LoopC As Long
                For LoopC = 1 To NpcData(Npc).NumQuiza
                    NpcData(Npc).QuizaDropea(LoopC) = Val(Leer.GetValue("npc" & Npc, "QuizaDropea" & LoopC))
                Next LoopC

            End If

            Label3.ForeColor = vbRed
            Label3.Caption = "Leyendo: " & Npc & "/" & numnpcs
        Next Npc
    
        Npc = 1
    
        Call Manager.ChangeValue("INIT", "NumNPCs", numnpcs)
    
        For Npc = 1 To numnpcs
            DoEvents

            If Len(NpcData(Npc).Name) <> 0 Then
                Call Manager.ChangeValue("Npc" & Npc, "Name", NpcData(Npc).Name)
            End If
                
            If Len(NpcData(Npc).en_Name) <> 0 Then
                Call Manager.ChangeValue("Npc" & Npc, "en_Name", NpcData(Npc).en_Name)
            End If
                
            If Len(NpcData(Npc).en_desc) <> 0 Then
                Call Manager.ChangeValue("Npc" & Npc, "en_desc", NpcData(Npc).en_desc)
            End If
        
            If Len(NpcData(Npc).desc) <> 0 Then
                Call Manager.ChangeValue("Npc" & Npc, "Desc", NpcData(Npc).desc)
            End If
        
            If NpcData(Npc).Body <> 0 Then
                Call Manager.ChangeValue("Npc" & Npc, "Body", NpcData(Npc).Body)
            End If

            If NpcData(Npc).Head <> 0 Then
                Call Manager.ChangeValue("Npc" & Npc, "Head", NpcData(Npc).Head)
            End If
        
            If NpcData(Npc).Exp <> 0 Then
                Call Manager.ChangeValue("Npc" & Npc, "Exp", NpcData(Npc).Exp)
            End If

            If NpcData(Npc).Hp <> 0 Then
                Call Manager.ChangeValue("Npc" & Npc, "Hp", NpcData(Npc).Hp)
            End If

            If NpcData(Npc).MaxHit <> 0 Then
                Call Manager.ChangeValue("Npc" & Npc, "MaxHit", NpcData(Npc).MaxHit)

            End If

            If NpcData(Npc).MinHit <> 0 Then
                Call Manager.ChangeValue("Npc" & Npc, "MinHit", NpcData(Npc).MinHit)
            End If
        
            If NpcData(Npc).Oro <> 0 Then
                Call Manager.ChangeValue("Npc" & Npc, "Oro", NpcData(Npc).Oro)
            End If
        
            If NpcData(Npc).ExpClan <> 0 Then
                Call Manager.ChangeValue("Npc" & Npc, "GiveEXPClan", NpcData(Npc).ExpClan)
            End If
        
            If NpcData(Npc).NumQuiza <> 0 Then
                Call Manager.ChangeValue("Npc" & Npc, "NumQuiza", NpcData(Npc).NumQuiza)
            
                For LoopC = 1 To NpcData(Npc).NumQuiza
                    Call Manager.ChangeValue("Npc" & Npc, "QuizaDropea" & LoopC, NpcData(Npc).QuizaDropea(LoopC))
                Next LoopC
            
            End If
            
            If NpcData(Npc).PuedeInvocar <> 0 Then
                Call Manager.ChangeValue("Npc" & Npc, "PuedeInvocar", NpcData(Npc).PuedeInvocar)
            End If
            
            Label3.Caption = "Grabando: " & Npc & "/" & numnpcs
            Label3.ForeColor = &HC0C0&
        
        Next Npc

    Else
        MsgBox "Falta el archivo npcs.dat dentro de la carpeta dats."

    End If
    
    If FileExist(App.Path & "\..\Recursos\Dat\hechizos.dat", vbNormal) Then

        Dim hechizosFile As String, numhechizos As Long
        hechizosFile = App.Path & "\..\Recursos\Dat\hechizos.dat"
        numhechizos = Val(GetVar(hechizosFile, "INIT", "NumeroHechizos"))

        Dim hechic As New clsIniReader
        Call hechic.Initialize(hechizosFile)

        Label3.Caption = "0/" & numhechizos

        ReDim HechizoData(1 To numhechizos) As HechizoDatas
    
        For hechizo = 1 To numhechizos
            DoEvents
            
            HechizoData(hechizo).Nombre = hechic.GetValue("Hechizo" & hechizo, "Nombre")
            HechizoData(hechizo).desc = hechic.GetValue("Hechizo" & hechizo, "desc")
            HechizoData(hechizo).PalabrasMagicas = hechic.GetValue("Hechizo" & hechizo, "PalabrasMagicas")
            HechizoData(hechizo).HechizeroMsg = hechic.GetValue("Hechizo" & hechizo, "HechizeroMsg")
            HechizoData(hechizo).TargetMsg = hechic.GetValue("Hechizo" & hechizo, "TargetMsg")
            HechizoData(hechizo).PropioMsg = hechic.GetValue("Hechizo" & hechizo, "PropioMsg")
            HechizoData(hechizo).ManaRequerido = Val(hechic.GetValue("Hechizo" & hechizo, "ManaRequerido"))
            HechizoData(hechizo).StaRequerido = Val(hechic.GetValue("Hechizo" & hechizo, "StaRequerido"))
            HechizoData(hechizo).MinSkill = Val(hechic.GetValue("Hechizo" & hechizo, "MinSkill"))
            HechizoData(hechizo).StaRequerido = Val(hechic.GetValue("Hechizo" & hechizo, "StaRequerido"))
            HechizoData(hechizo).IconoIndex = Val(hechic.GetValue("Hechizo" & hechizo, "IconoIndex"))
            Label3.ForeColor = vbRed
            Label3.Caption = "Leyendo: " & hechizo & "/" & numhechizos
        Next hechizo

        Call Manager.ChangeValue("INIT", "NumeroHechizo", numhechizos)
    
        For hechizo = 1 To numhechizos
            DoEvents
            
            Call Manager.ChangeValue("Hechizo" & hechizo, "Nombre", HechizoData(hechizo).Nombre)
            Call Manager.ChangeValue("Hechizo" & hechizo, "Desc", HechizoData(hechizo).desc)
            Call Manager.ChangeValue("Hechizo" & hechizo, "PalabrasMagicas", HechizoData(hechizo).PalabrasMagicas)
            Call Manager.ChangeValue("Hechizo" & hechizo, "HechizeroMsg", HechizoData(hechizo).HechizeroMsg)
            Call Manager.ChangeValue("Hechizo" & hechizo, "TargetMsg", HechizoData(hechizo).TargetMsg)
            Call Manager.ChangeValue("Hechizo" & hechizo, "PropioMsg", HechizoData(hechizo).PropioMsg)
            Call Manager.ChangeValue("Hechizo" & hechizo, "ManaRequerido", HechizoData(hechizo).ManaRequerido)
            Call Manager.ChangeValue("Hechizo" & hechizo, "StaRequerido", HechizoData(hechizo).StaRequerido)
            Call Manager.ChangeValue("Hechizo" & hechizo, "MinSkill", HechizoData(hechizo).MinSkill)
            Call Manager.ChangeValue("Hechizo" & hechizo, "IconoIndex", HechizoData(hechizo).IconoIndex)

            Label3.Caption = "Grabando: " & hechizo & "/" & numhechizos
            Label3.ForeColor = &HC0C0&
        Next hechizo

    End If
    
    If FileExist(App.Path & "\..\Recursos\init\LocalMsg.dat", vbNormal) Then
        
        Dim MsgFile As String
            MsgFile = App.Path & "\..\Recursos\init\LocalMsg.dat"

        Dim Msgsss As New clsIniReader
        Call Msgsss.Initialize(MsgFile)

        numnpcs = Val(Msgsss.GetValue("INIT", "NumLocaleMsg"))

        Label3.Caption = "0/" & CStr(numnpcs)

        ReDim arrLocale_SMG(1 To numnpcs) As String
    
        For Npc = 1 To numnpcs
            DoEvents
            
            arrLocale_SMG(Npc) = Msgsss.GetValue("msg", "Msg" & Npc)
        
            Label3.ForeColor = vbRed
            Label3.Caption = "Leyendo: " & Npc & "/" & numnpcs
        Next Npc

        Npc = 1
    
        Call Manager.ChangeValue("INIT", "NumLocaleMsg", numnpcs)
    
        For Npc = 1 To numnpcs
            DoEvents

            Call Manager.ChangeValue("Msg", "Msg" & Npc, arrLocale_SMG(Npc))

            Label3.Caption = "Grabando: " & Npc & "/" & numnpcs
            Label3.ForeColor = &HC0C0&
        Next Npc

    Else
        MsgBox "Falta el archivo LocalMsg.dat dentro de la carpeta dats."

    End If

    If FileExist(App.Path & "\..\Recursos\init\NameMapa.dat", vbNormal) Then
        
        Dim MapFile As String
            MapFile = App.Path & "\..\Recursos\init\NameMapa.dat"

        Dim Mapa As New clsIniReader
        Call Mapa.Initialize(MapFile)

        Label3.Caption = "0/" & 306

        ReDim MapName(1 To 306) As String
        ReDim MapDesc(1 To 306) As String
    
        For Npc = 1 To 306
        
            DoEvents
            MapName(Npc) = Mapa.GetValue("NameMapa", "mapa" & Npc)
            MapDesc(Npc) = Mapa.GetValue("NameMapa", "mapa" & Npc & "desc")
        
            Label3.ForeColor = vbRed
            Label3.Caption = "Leyendo: " & Npc & "/" & 304
        Next Npc

        Npc = 1
    
        Call Manager.ChangeValue("INIT", "NumMapas", 306)
    
        For Npc = 1 To 306
            DoEvents

            Call Manager.ChangeValue("NAMEMAPA", "Mapa" & Npc, MapName(Npc))
            Call Manager.ChangeValue("NAMEMAPA", "Mapa" & Npc & "Desc", MapDesc(Npc))

            Label3.Caption = "Grabando: " & Npc & "/" & 306
            Label3.ForeColor = &HC0C0&
        Next Npc

    Else
        MsgBox "Falta el archivo NameMapa.dat dentro de la carpeta dats."

    End If

    'quest

    If FileExist(App.Path & "\..\Recursos\Dat\Quests.DAT", vbNormal) Then

        MapFile = App.Path & "\..\Recursos\Dat\Quests.DAT"

        Call Mapa.Initialize(MapFile)
        
        Dim nunquest As Integer
            nunquest = Mapa.GetValue("INIT", "NumQuests")

        Label3.Caption = "0/" & nunquest

        ReDim QuestName(1 To nunquest) As String
        ReDim QuestDesc(1 To nunquest) As String
    
        ReDim QuestFin(1 To nunquest) As String
        ReDim QuestNext(1 To nunquest) As String
        ReDim QuestPos(1 To nunquest) As Integer
        ReDim QuestRepetible(1 To nunquest) As Byte
    
        ReDim RequiredLevel(1 To nunquest) As Integer
    
        For Npc = 1 To nunquest
            DoEvents
            
            QuestName(Npc) = Mapa.GetValue("QUEST" & Npc, "Nombre")
            QuestDesc(Npc) = Mapa.GetValue("QUEST" & Npc, "Desc")
        
            QuestFin(Npc) = Mapa.GetValue("QUEST" & Npc, "DescFinal")
            QuestNext(Npc) = Mapa.GetValue("QUEST" & Npc, "NextQuest")
            QuestRepetible(Npc) = Val(Mapa.GetValue("QUEST" & Npc, "Repetible"))
        
            QuestPos(Npc) = Val(Mapa.GetValue("QUEST" & Npc, "PosMap"))
        
            RequiredLevel(Npc) = Val(Mapa.GetValue("QUEST" & Npc, "RequiredLevel"))
        
            Label3.ForeColor = vbRed
            Label3.Caption = "Leyendo: " & Npc & "/" & nunquest
        Next Npc

        Npc = 1
    
        Call Manager.ChangeValue("INIT", "NumQuests", nunquest)
    
        For Npc = 1 To nunquest
            DoEvents

            Call Manager.ChangeValue("QUEST" & Npc, "Nombre", QuestName(Npc))
            Call Manager.ChangeValue("QUEST" & Npc, "Desc", QuestDesc(Npc))
        
            Call Manager.ChangeValue("QUEST" & Npc, "DescFinal", QuestFin(Npc))
            Call Manager.ChangeValue("QUEST" & Npc, "NextQuest", QuestNext(Npc))
            Call Manager.ChangeValue("QUEST" & Npc, "Repetible", QuestRepetible(Npc))
        
            Call Manager.ChangeValue("QUEST" & Npc, "RequiredLevel", RequiredLevel(Npc))
        
            Call Manager.ChangeValue("QUEST" & Npc, "PosMap", QuestPos(Npc))

            Label3.Caption = "Grabando: " & Npc & "/" & nunquest
            Label3.ForeColor = &HC0C0&
        Next Npc

    Else
        MsgBox "Falta el archivo Quests.DAT dentro de la carpeta dats."

    End If

    If FileExist(App.Path & "\..\Recursos\init\sugerencias.ini", vbNormal) Then

        MapFile = App.Path & "\..\Recursos\init\sugerencias.ini"

        Call Mapa.Initialize(MapFile)
        
        Dim NumSug As Integer
            NumSug = Val(Mapa.GetValue("Sugerencias", "NumSugerencias"))

        Label3.Caption = "0/" & CStr(NumSug)

        ReDim Sugerencia(1 To NumSug) As String
    
        For Npc = 1 To NumSug
            DoEvents
            
            Sugerencia(Npc) = Mapa.GetValue("Sugerencias", "Sugerencia" & Npc)
        
            Label3.ForeColor = vbRed
            Label3.Caption = "Leyendo: " & Npc & "/" & nunquest
        Next Npc

        Npc = 1
    
        Call Manager.ChangeValue("INIT", "NumSugerencias", NumSug)
    
        For Npc = 1 To NumSug
            DoEvents

            Call Manager.ChangeValue("Sugerencias", "Sugerencia" & Npc, Sugerencia(Npc))

            Label3.Caption = "Grabando: " & Npc & "/" & NumSug
            Label3.ForeColor = &HC0C0&
        Next Npc

    Else
        MsgBox "Falta el archivo Sugerencias.ini dentro de la carpeta init."

    End If

    Dim ListaRazas(1 To NUMRAZAS) As String
    ListaRazas(1) = "Humano"
    ListaRazas(2) = "Elfo"
    ListaRazas(3) = "Elfo Oscuro"
    ListaRazas(4) = "Gnomo"
    ListaRazas(5) = "Enano"
    ListaRazas(6) = "Orco"
    
    Call Leer.Initialize(App.Path & "\..\Recursos\Dat\Balance.dat")
    
    Dim SearchVar As String

    For Raza = 1 To NUMRAZAS

        With ModRaza(Raza)
            SearchVar = Replace(ListaRazas(Raza), " ", vbNullString)

            .Fuerza = Val(Leer.GetValue("MODRAZA", SearchVar + "Fuerza"))
            .Agilidad = Val(Leer.GetValue("MODRAZA", SearchVar + "Agilidad"))
            .Inteligencia = Val(Leer.GetValue("MODRAZA", SearchVar + "Inteligencia"))
            .Constitucion = Val(Leer.GetValue("MODRAZA", SearchVar + "Constitucion"))
            .Carisma = Val(Leer.GetValue("MODRAZA", SearchVar + "Carisma"))
            
            Call Manager.ChangeValue("MODRAZA", SearchVar + "Fuerza", .Fuerza)
            Call Manager.ChangeValue("MODRAZA", SearchVar + "Agilidad", .Agilidad)
            Call Manager.ChangeValue("MODRAZA", SearchVar + "Inteligencia", .Inteligencia)
            Call Manager.ChangeValue("MODRAZA", SearchVar + "Constitucion", .Constitucion)
            Call Manager.ChangeValue("MODRAZA", SearchVar + "Carisma", .Carisma)

        End With

    Next Raza
    
    Set Leer = Nothing
       
    Call Manager.DumpFile(OutputFile)
    
    Set Manager = Nothing
    
    Label3.ForeColor = vbGreen
    Label3.Caption = "Creado localindex.dat"

End Sub

Private Sub Command2_Click()
    Form2.Show

End Sub

Public Sub LeerLineaComandos()
    Dim rdata As String
    rdata = Command
    
    Dim FileTypeName As String
    Dim FileTypeIndex As Integer
      
    FileTypeName = ReadField(1, rdata, Asc("*")) ' File Type Name

    If Len(FileTypeName) > 0 Then
    
        FileTypeName = UCase(FileTypeName)
        
        Select Case FileTypeName
            Case Is = "CREAR_ARCHIVO"
                Call Command1_Click
            End Select
        End
    
    End If

End Sub

Private Sub Form_Load()
    OutputFile = App.Path & "\..\Recursos\init\localindex.dat"
    
    ' Leer argumentos
    Call LeerLineaComandos
End Sub

