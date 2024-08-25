VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMensajes 
      BackColor       =   &H000080FF&
      Caption         =   "Mensajes"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdCrearArchivo 
      BackColor       =   &H0000C0FF&
      Caption         =   "Crear archivo"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
