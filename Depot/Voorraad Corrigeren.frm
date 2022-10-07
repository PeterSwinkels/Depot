VERSION 5.00
Begin VB.Form VoorraadCorrigerenVenster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artikel"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   27.125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox AantalVeld 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      MaxLength       =   250
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox ErAfTotaalVeld 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      MaxLength       =   250
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox ErBijTotaalVeld 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      MaxLength       =   250
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton AnnulerenKnop 
      Cancel          =   -1  'True
      Caption         =   "&Annuleren"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton CorrigerenKnop 
      Caption         =   "&Corrigeren"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label AantalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Aantal:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label ErAfLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Er Af Totaal:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label ErBijLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Er Bij Totaal:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "VoorraadCorrigerenVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het voorraad corrigeren venster.
Option Explicit

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub AantalVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   AantalVeld.SelStart = 0
   AantalVeld.SelLength = Len(AantalVeld.Text)
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure sluit dit venster.
Private Sub AnnulerenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   Unload Me
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waardes vast.
Private Sub CorrigerenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   Buffer(VrAantal) = MaakGetalOp(AantalVeld.Text)
   Buffer(VrAftotaal) = MaakGetalOp(ErAfTotaalVeld.Text)
   Buffer(VrBijTotaal) = MaakGetalOp(ErBijTotaalVeld.Text)
   
   Unload Me
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub ErAfTotaalVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long
   
   ErAfTotaalVeld.SelStart = 0
   ErAfTotaalVeld.SelLength = Len(ErAfTotaalVeld.Text)
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub ErBijTotaalVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long
   
   ErBijTotaalVeld.SelStart = 0
   ErBijTotaalVeld.SelLength = Len(ErBijTotaalVeld.Text)
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
   
   AantalVeld.ToolTipText = "Het aantal artikelen."
   AnnulerenKnop.ToolTipText = "Sluit dit venster."
   CorrigerenKnop.ToolTipText = "Corrigeert de aantallen."
   ErAfTotaalVeld.ToolTipText = "Het er af totaal."
   ErBijTotaalVeld.ToolTipText = "Het er bij totaal."

   AantalVeld.Text = Buffer(VrAantal)
   ErAfTotaalVeld.Text = Buffer(VrAftotaal)
   ErBijTotaalVeld.Text = Buffer(VrBijTotaal)
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

