VERSION 5.00
Begin VB.Form VoorraadVenster 
   Caption         =   "Voorraad"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   16.563
   ScaleMode       =   4  'Character
   ScaleWidth      =   52.125
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox KnoppenBalk 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   6015
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3480
      Width           =   6015
      Begin VB.CommandButton AfdrukkenKnop 
         Caption         =   "&Afdrukken"
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton CorrigerenKnop 
         Caption         =   "&Corrigeren"
         Height          =   375
         Left            =   4800
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton WijzigenKnop 
         Caption         =   "&Wijzigen"
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton ToevoegenKnop 
         Caption         =   "&Toevoegen"
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton VerwijderenKnop 
         Caption         =   "&Verwijderen"
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
   End
   Begin DepotbeheerderProgramma.LijstObject ArtikelenLijst 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _extentx        =   9340
      _extenty        =   5741
   End
End
Attribute VB_Name = "VoorraadVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het voorraad venster.
Option Explicit

'Deze procedure drukt na bevestiging van de gebruiker de voorraad lijst af wanneer de beheerder is ingelogd.
Private Sub AfdrukkenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   If Beheerder() Then
      If MsgBox("Voorraad afdrukken?", vbQuestion Or vbYesNo) = vbYes Then
         DrukDataAf Voorraad
      End If
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure toont het voorraad correctie venster wanneer de beheerder is ingelogd.
Private Sub CorrigerenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   If Beheerder() Then
      ActieIsToeVoegen = False
      If Not Voorraad.AantalItems() = 0 Then
         Buffer(VrAantal) = Voorraad.Data(VrAantal, ArtikelenLijst.Selectie())
         Buffer(VrAftotaal) = Voorraad.Data(VrAftotaal, ArtikelenLijst.Selectie())
         Buffer(VrBijTotaal) = Voorraad.Data(VrBijTotaal, ArtikelenLijst.Selectie())

         VoorraadCorrigerenVenster.Show vbModal
         Voorraad.Corrigeer ArtikelenLijst.Selectie()
         ArtikelenLijst.WerkLijstBij
      End If
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure werkt de artikelen lijst bij wanneer dit venster wordt geactiveerd.
Private Sub Form_Activate()
On Error GoTo Fout
Dim Keuze As Long

   ArtikelenLijst.WerkLijstBij

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure kopieeert het geselecteerde artikel.
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Fout
Dim Keuze As Long

   If KeyCode = vbKeyC And Shift = vbCtrlMask Then
      GekopieerdArtikel = Voorraad.Data(VrArtikelnr, ArtikelenLijst.Selectie())
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

   AfdrukkenKnop.ToolTipText = "Drukt de voorraad af."
   CorrigerenKnop.ToolTipText = "Corrigeert een artikel."
   ToevoegenKnop.ToolTipText = "Voegt een artikel toe."
   VerwijderenKnop.ToolTipText = "Verwijdert een artikel."
   WijzigenKnop.ToolTipText = "Wijzigt een artikel."

   Me.Width = DepotbeheerderVenster.Width / 1.3
   Me.Height = DepotbeheerderVenster.Height / 1.15
   Me.Left = MenuVenster.Left + MenuVenster.Width + 128
   Me.Top = MenuVenster.Top

   Set ArtikelenLijst.Databron = Voorraad

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

   ArtikelenLijst.Width = Me.ScaleWidth - 2
   ArtikelenLijst.Height = Me.ScaleHeight - 3
   KnoppenBalk.Left = Me.ScaleWidth - KnoppenBalk.Width - 1
   KnoppenBalk.Top = Me.ScaleHeight - 2

   ArtikelenLijst.WerkLijstBij
End Sub

'Deze procedure voegt een artikel toe aan de voorraad wanneer de beheerder is ingelogd.
Private Sub ToevoegenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   If Beheerder() Then
      ActieIsToeVoegen = True
      Voorraad.StelStandaardWaardesIn
      VoorraadInvoerVenster.Show vbModal

      If Not Buffer(VrArtikelnr) = vbNullString Then
         Voorraad.VoegItemToe
         ArtikelenLijst.Selectie = Voorraad.AantalItems() - 1
         Voorraad.WijzigItem ArtikelenLijst.Selectie()
         Voorraad.SorteerItems
         Voorraad.VerwijderOudArtikel
         ArtikelenLijst.WerkLijstBij
      End If
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure verwijdert een artikel wanneer de beheerder is ingelogd.
Private Sub VerwijderenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   If Beheerder() Then
      If MsgBox("Artikel verwijderen?", vbQuestion Or vbYesNo) = vbYes Then
         Voorraad.VerwijderItem ArtikelenLijst.Selectie()
         ArtikelenLijst.WerkLijstBij
      End If
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure wijzigt een artikel in de voorraad wanneer de beheerder is ingelogd.
Private Sub WijzigenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   If Beheerder() Then
      If Not Voorraad.AantalItems() = 0 Then
         Buffer(VrArtikelnr) = Voorraad.Data(VrArtikelnr, ArtikelenLijst.Selectie())
         Buffer(VrNaam) = Voorraad.Data(VrNaam, ArtikelenLijst.Selectie())
         Buffer(VrStukprijs) = Voorraad.Data(VrStukprijs, ArtikelenLijst.Selectie())

         ActieIsToeVoegen = False
         VoorraadInvoerVenster.Show vbModal

         If Not Buffer(VrArtikelnr) = vbNullString Then
            Voorraad.WijzigItem ArtikelenLijst.Selectie()
            Voorraad.SorteerItems
            Voorraad.VerwijderOudArtikel
            ArtikelenLijst.WerkLijstBij
         End If
      End If
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

