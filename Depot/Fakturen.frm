VERSION 5.00
Begin VB.Form FakturenVenster 
   Caption         =   "Fakturen"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   ClipControls    =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   18.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   38.125
   Begin VB.PictureBox KnoppenBalk 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   375
      Left            =   2160
      ScaleHeight     =   1.563
      ScaleMode       =   4  'Character
      ScaleWidth      =   19.125
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3840
      Width           =   2295
      Begin VB.CommandButton AnnulerenKnop 
         Cancel          =   -1  'True
         Caption         =   "&Annuleren"
         Height          =   375
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton OpenenKnop 
         Caption         =   "&Openen"
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   1095
      End
   End
   Begin DepotBeheerderProgramma.LijstObject FakturenLijst 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6376
   End
End
Attribute VB_Name = "FakturenVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het fakturen venster.
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

'Deze procudure werkt de lijst van opgeslagen fakturen bij.
Private Sub Form_Activate()
On Error GoTo Fout
Dim Keuze As Long

   Fakturen.MaakLijst
   FakturenLijst.WerkLijstBij
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
   OpenenKnop.ToolTipText = "Opent het geselecteerde fakturen."
   
   Me.Left = MenuVenster.Left + MenuVenster.Width + 128
   Me.Top = MenuVenster.Top
   Me.Width = DepotBeheerderVenster.Width / 2
   Me.Height = DepotBeheerderVenster.Height / 2
   
   Fakturen.MaakLijst
   Set FakturenLijst.DataBron = Fakturen
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
   FakturenLijst.Width = Me.ScaleWidth - 2
   FakturenLijst.Height = Me.ScaleHeight - 3
   KnoppenBalk.Left = Me.ScaleWidth - KnoppenBalk.Width - 1
   KnoppenBalk.Top = Me.ScaleHeight - 2
   
   FakturenLijst.WerkLijstBij
End Sub

'Deze procedure geeft de opdracht om de geselecteerde faktuur te openen en sluit dit venster.
Private Sub OpenenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   If Not Fakturen.AantalItems() = 0 Then
      Faktuur.OpenFaktuur Fakturen.Data(FkFaktuurNummer, FakturenLijst.Selectie())
      Unload Me
   End If
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

