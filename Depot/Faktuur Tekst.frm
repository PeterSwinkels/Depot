VERSION 5.00
Begin VB.Form FaktuurtekstVenster 
   Caption         =   "Faktuurtekst"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   ClipControls    =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   15.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   38.125
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox KnoppenBalk 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   375
      Left            =   1920
      ScaleHeight     =   1.563
      ScaleMode       =   4  'Character
      ScaleWidth      =   21.125
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3120
      Width           =   2535
      Begin VB.CommandButton WijzigenKnop 
         Caption         =   "&Wijzigen"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton AnnulerenKnop 
         Cancel          =   -1  'True
         Caption         =   "&Annuleren"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox FaktuurtekstVeld 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FaktuurtekstVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het faktuur tekst venster.
Option Explicit

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

'Deze procedure stelt dit venster in.
Private Sub Form_Load()
On Error GoTo Fout
Dim Keuze As Long

   AnnulerenKnop.ToolTipText = "Sluit dit venster."
   FaktuurtekstVeld.ToolTipText = "Deze tekst wordt op de fakturen afgedrukt."
   WijzigenKnop.ToolTipText = "Wijzigt de faktuur tekst."

   Me.Width = DepotbeheerderVenster.Width / 2
   Me.Height = DepotbeheerderVenster.Height / 2
   Me.Left = (DepotbeheerderVenster.Width / 2) - (Me.Width / 2)
   Me.Top = (DepotbeheerderVenster.Height / 3) - (Me.Height / 2)

   FaktuurtekstVeld.Text = Faktuur.FaktuurTekst()

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure past de objecten in dit venster aan de nieuwe afmetingen aan.
Private Sub Form_Resize()
On Error Resume Next

   FaktuurtekstVeld.Width = Me.ScaleWidth - 2
   FaktuurtekstVeld.Height = Me.ScaleHeight - 3
   KnoppenBalk.Left = Me.ScaleWidth - KnoppenBalk.Width - 1
   KnoppenBalk.Top = Me.ScaleHeight - 2
End Sub

'Deze procedure legt de ingevoerde faktuur tekst vast en sluit dit venster.
Private Sub WijzigenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   Faktuur.FaktuurTekst = FaktuurtekstVeld.Text
   Unload Me

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

