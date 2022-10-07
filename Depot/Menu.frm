VERSION 5.00
Begin VB.Form MenuVenster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   13.125
   ScaleMode       =   4  'Character
   ScaleWidth      =   14.125
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton InstellingenKnop 
      Caption         =   "&Instellingen"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton KlantenKnop 
      Caption         =   "&Klanten"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton SluitenKnop 
      Caption         =   "&Sluiten"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton VoorraadKnop 
      Caption         =   "&Voorraad"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton FakturenKnop 
      Caption         =   "&Fakturen"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "MenuVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het menu venster.
Option Explicit

'Deze procedure toont het faktuur venster.
Private Sub FakturenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   FaktuurVenster.Show
   FaktuurVenster.ZOrder
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stelt dit venster in.
Private Sub Form_Load()
On Error GoTo Fout
Dim Keuze As Long
   
   FakturenKnop.ToolTipText = "Toont het fakturen venster."
   InstellingenKnop.ToolTipText = "Toont het instellingen venster."
   KlantenKnop.ToolTipText = "Toont het klanten venster."
   SluitenKnop.ToolTipText = "Sluit dit programma."
   VoorraadKnop.ToolTipText = "Toont het voorraad venster."
      
   Me.Left = 128
   Me.Top = 128
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure toont het instellingen venster.
Private Sub InstellingenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   InstellingenVenster.Show
   InstellingenVenster.ZOrder
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure toont het klanten venster.
Private Sub KlantenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   KlantenVenster.Show
   KlantenVenster.ZOrder
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure sluit dit programma af.
Private Sub SluitenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   Unload DepotBeheerderVenster
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure toont het voorraad venster.
Private Sub VoorraadKnop_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   VoorraadVenster.Show
   VoorraadVenster.ZOrder
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

