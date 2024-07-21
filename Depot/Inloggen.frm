VERSION 5.00
Begin VB.Form InloggenVenster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inloggen"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5.625
   ScaleMode       =   4  'Character
   ScaleWidth      =   42.625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton GebruikerKnop 
      Cancel          =   -1  'True
      Caption         =   "&Gebruiker"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton BeheerderKnop 
      Caption         =   "&Beheerder"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox BeheerderwachtwoordVeld 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label LogInAlsLabel 
      Caption         =   "Log in als:"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.Label BeheerderwachtwoordLabel 
      Caption         =   "Beheerderwachtwoord:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "InloggenVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze procedure bevat het wachtwoord venster.
Option Explicit
Private VorigeMuisAanwijzer As Integer

'Deze procedure logt in met het opgegeven wachtwoord.
Private Sub BeheerderKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   If BeheerderWachtwoordVeld.Text = vbNullString Then
      MsgBox "Geen wachtwoord ingevoerd.", vbExclamation
   Else
      Screen.MousePointer = VorigeMuisAanwijzer
      IngevoerdWachtwoord = ZetBitsOm(BeheerderWachtwoordVeld.Text)
      Unload Me
   End If

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

   BeheerderKnop.ToolTipText = "Logt in als beheerder."
   BeheerderWachtwoordVeld.ToolTipText = "Het beheerder wachtwoord."
   GebruikerKnop.ToolTipText = "Logt in als gebruiker."

   VorigeMuisAanwijzer = Screen.MousePointer
   Screen.MousePointer = vbDefault

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

Private Sub GebruikerKnop_Click()
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

