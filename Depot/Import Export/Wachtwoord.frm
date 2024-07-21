VERSION 5.00
Begin VB.Form WachtwoordVenster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Beheerderwachtwoord"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4.125
   ScaleMode       =   4  'Character
   ScaleWidth      =   27.125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton InloggenKnop 
      Caption         =   "&Inloggen"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox BeheerderwachtwoordVeld 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "WachtwoordVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze procedure bevat het wachtwoord venster.
Option Explicit
Private VorigeMuisAanwijzer As Integer

'Deze procedure stelt dit venster in.
Private Sub Form_Load()
On Error GoTo Fout
Dim Keuze As Long

   BeheerderWachtwoordVeld.ToolTipText = "Het beheerder wachtwoord."
   InloggenKnop.ToolTipText = "Logt in als gebruiker of beheerder."
   
   VorigeMuisAanwijzer = Screen.MousePointer
   Screen.MousePointer = vbDefault

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure logt in met het opgegeven wachtwoord.
Private Sub InloggenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   Screen.MousePointer = VorigeMuisAanwijzer
   IngevoerdWachtwoord = ZetBitsOm(BeheerderWachtwoordVeld.Text)
   
   Unload Me
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

