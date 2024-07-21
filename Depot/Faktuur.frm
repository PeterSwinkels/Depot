VERSION 5.00
Begin VB.Form FaktuurVenster 
   Caption         =   "Faktuur"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10710
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   35.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   89.25
   ShowInTaskbar   =   0   'False
   Begin DepotbeheerderProgramma.LijstObject ArtikelenLijst 
      Height          =   3375
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5953
   End
   Begin VB.PictureBox KlantVelden 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   2775
      Left            =   240
      ScaleHeight     =   2775
      ScaleWidth      =   5175
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   120
      Width           =   5175
      Begin VB.TextBox KlantEmailAdresVeld 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   250
         TabIndex        =   6
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox KlantFaxnummerVeld 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   250
         TabIndex        =   5
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox KlantTelefoonnummerVeld 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   250
         TabIndex        =   4
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox KlantPostcodePlaatsVeld 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   250
         TabIndex        =   3
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox KlantAdresVeld 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   250
         TabIndex        =   2
         Top             =   960
         Width           =   2535
      End
      Begin VB.CheckBox ParticulierVeld 
         Caption         =   "Pa&rticulier"
         Height          =   255
         Left            =   4200
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox KlantnummerVeld 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         MaxLength       =   250
         TabIndex        =   0
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox KlantNaamVeld 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   250
         TabIndex        =   1
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label FaxnummerLabel 
         Caption         =   "Fanxnummer:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label EMailAdresLabel 
         Caption         =   "E-mail Adres:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label TelefoonnummerLabel 
         Caption         =   "Telefoonnummer:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label PostcodePlaatsLabel 
         Caption         =   "Postcode/Plaats:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label AdresLabel 
         Caption         =   "Adres:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label NaamLabel 
         Caption         =   "Naam:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label NummerLabel 
         Caption         =   "Nummer:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label KlantStatus 
         Caption         =   "Klant:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.PictureBox FaktuurVelden 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   855
      Left            =   7560
      ScaleHeight     =   855
      ScaleWidth      =   3015
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   120
      Width           =   3015
      Begin VB.TextBox DatumVeld 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   255
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox FaktuurnummerVeld 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   255
         TabIndex        =   8
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label FacktuurnummerLabel 
         Caption         =   "Faktuurnummer:"
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label DatumLabel 
         Caption         =   "Datum:"
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.PictureBox TotaalVelden 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1335
      Left            =   7560
      ScaleHeight     =   1335
      ScaleWidth      =   2895
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   6360
      Width           =   2895
      Begin VB.TextBox BTWVeld 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   0
         MaxLength       =   3
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox PortoVeld 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         MaxLength       =   250
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label BTWBedragVeld 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1200
         TabIndex        =   40
         Top             =   330
         Width           =   1575
      End
      Begin VB.Label BTWLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "% BTW:"
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   360
         Width           =   615
      End
      Begin VB.Label TotaalVeld 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   36
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label SubTotaalVeld 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label TotaalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Totaal:"
         Height          =   255
         Left            =   600
         TabIndex        =   35
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label PortoLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Porto:"
         Height          =   255
         Left            =   600
         TabIndex        =   34
         Top             =   720
         Width           =   495
      End
      Begin VB.Label SubtaalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Subtotaal:"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.PictureBox KnoppenBalk 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1935
      ScaleWidth      =   10455
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6360
      Width           =   10455
      Begin VB.CommandButton WijzigenKnop 
         Caption         =   "&Wijzigen"
         Height          =   375
         Left            =   9240
         TabIndex        =   19
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton NieuwKnop 
         Caption         =   "&Nieuw"
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton AfdrukkenKnop 
         Caption         =   "&Afdrukken"
         Height          =   375
         Left            =   5640
         TabIndex        =   16
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton OpslaanKnop 
         Caption         =   "O&pslaan"
         Height          =   375
         Left            =   4440
         TabIndex        =   15
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton OpenenKnop 
         Caption         =   "&Openen"
         Height          =   375
         Left            =   3240
         TabIndex        =   14
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton VerwijderenKnop 
         Caption         =   "&Verwijderen"
         Height          =   375
         Left            =   8040
         TabIndex        =   18
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton ToevoegenKnop 
         Caption         =   "&Toevoegen"
         Height          =   375
         Left            =   6840
         TabIndex        =   17
         Top             =   1560
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FaktuurVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het faktuur venster.
Option Explicit

'Deze procedure werkt de geopende faktuur bij.
Private Sub WerkFaktuurBij()
On Error GoTo Fout
Dim Keuze As Long

   Faktuur.WerkArtikelenBij
   ArtikelenLijst.WerkLijstBij

   DatumVeld.Text = Faktuur.Datum()
   FaktuurnummerVeld.Text = Faktuur.Faktuurnr()
   KlantnummerVeld.Text = Faktuur.Klantnummer()
   If Faktuur.IsParticulier() Then ParticulierVeld.Value = vbChecked Else ParticulierVeld.Value = vbUnchecked

   Faktuur.WerkKlantgegevensBij
   KlantStatus.Caption = "Klant"
   If Faktuur.KlantGevonden() Then KlantStatus.Caption = KlantStatus.Caption & ": " Else KlantStatus.Caption = KlantStatus.Caption & " bestaat niet! "
   KlantnummerVeld.Text = Faktuur.Klantnummer()
   KlantnaamVeld.Text = Faktuur.Klantnaam()
   KlantAdresVeld.Text = Faktuur.KlantAdres()
   KlantPostcodePlaatsVeld.Text = Faktuur.KlantPostcodePlaats()
   KlantTelefoonnummerVeld.Text = Faktuur.KlantTelefoonnummer()
   KlantFaxnummerVeld.Text = Faktuur.KlantFaxNummer()
   KlantEmailadresVeld.Text = Faktuur.KlantEmailadres()

   KlantnaamVeld.Locked = Not Faktuur.IsParticulier()
   KlantAdresVeld.Locked = Not Faktuur.IsParticulier()
   KlantPostcodePlaatsVeld.Locked = Not Faktuur.IsParticulier()
   KlantTelefoonnummerVeld.Locked = Not Faktuur.IsParticulier()
   KlantFaxnummerVeld.Locked = Not Faktuur.IsParticulier()
   KlantEmailadresVeld.Locked = Not Faktuur.IsParticulier()

   WerkTotalenBij

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure werkt de totaal velden bij.
Private Sub WerkTotalenBij()
On Error GoTo Fout
Dim Keuze As Long

   BTWVeld.Text = Faktuur.BTW()
   BTWBedragVeld.Caption = Faktuur.BTWBedrag()
   PortoVeld.Text = Faktuur.Porto()
   SubTotaalVeld.Caption = Faktuur.Subtotaal()
   TotaalVeld.Caption = Faktuur.Totaal()

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure drukt het faktuur af na bevestiging van de gebruiker.
Private Sub AfdrukkenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   If MsgBox("Faktuur afdrukken?", vbQuestion Or vbYesNo) = vbYes Then
      Faktuur.DrukFaktuurAf
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub BTWVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   BTWVeld.SelStart = 0
   BTWVeld.SelLength = Len(BTWVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waarde vast en past het faktuur aan.
Private Sub BTWVeld_LostFocus()
On Error GoTo Fout
Dim Keuze As Long

   BTWVeld.Text = MaakGetalOp(BTWVeld.Text)
   Faktuur.BTW() = BTWVeld.Text

   WerkTotalenBij

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub DatumVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   DatumVeld.SelStart = 0
   DatumVeld.SelLength = Len(DatumVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waarde vast.
Private Sub DatumVeld_LostFocus()
On Error GoTo Fout
Dim Keuze As Long

   Faktuur.Datum = DatumVeld.Text

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub FaktuurNummerVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   FaktuurnummerVeld.SelStart = 0
   FaktuurnummerVeld.SelLength = Len(FaktuurnummerVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure filtert ongeldige tekens uit de invoer.
Private Sub FaktuurNummerVeld_KeyPress(KeyAscii As Integer)
On Error GoTo Fout
Dim Keuze As Long

   If InStr(ONGELDIGE_TEKENS, Chr$(KeyAscii)) > 0 Then KeyAscii = Empty

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waarde vast.
Private Sub FaktuurNummerVeld_LostFocus()
On Error GoTo Fout
Dim Keuze As Long

   FaktuurnummerVeld.Text = UCase$(FaktuurnummerVeld.Text)
   Faktuur.Faktuurnr = FaktuurnummerVeld.Text

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht het faktuur bij te werken waneer het venster geactiveerd wordt.
Private Sub Form_Activate()
On Error GoTo Fout
Dim Keuze As Long

   WerkFaktuurBij

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure plakt uit de voorraad gekopieerde artikelen in de lijst.
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Fout
Dim Gevonden As Boolean
Dim Keuze As Long

   If KeyCode = vbKeyV And Shift = vbCtrlMask Then
      If Not GekopieerdArtikel = vbNullString Then
         Buffer(FkArtikelNr) = GekopieerdArtikel
         Buffer(FkAantal) = "0"
         Buffer(FkUitVoorraad) = "J"
         Faktuur.VoegItemToe
         ArtikelenLijst.Selectie = Faktuur.AantalItems() - 1
         Faktuur.WijzigItem ArtikelenLijst.Selectie, Gevonden
         If Not Gevonden Then MsgBox "Artikel kan niet worden gevonden.", vbExclamation
         Faktuur.SorteerArtikelen
         ArtikelenLijst.WerkLijstBij
         WerkTotalenBij
         WerkFaktuurBij
      End If
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

   AfdrukkenKnop.ToolTipText = "Drukt dit faktuur af."
   BTWBedragVeld.ToolTipText = "Het BTW bedrag."
   BTWVeld.ToolTipText = "Het BTW percentage."
   DatumVeld.ToolTipText = "De faktuur datum."
   FaktuurnummerVeld.ToolTipText = "Het faktuurnummer."
   KlantAdresVeld.ToolTipText = "Het klant adres."
   KlantEmailadresVeld.ToolTipText = "Het klant e-mailadres."
   KlantFaxnummerVeld.ToolTipText = "Het klant faxnummer."
   KlantnaamVeld.ToolTipText = "De klant naam."
   KlantnummerVeld.ToolTipText = "Het klantnummer."
   KlantPostcodePlaatsVeld.ToolTipText = "De klant postcode en plaats."
   KlantTelefoonnummerVeld.ToolTipText = "Het klant telefoonnummer."
   NieuwKnop.ToolTipText = "Opent een nieuw faktuur."
   OpenenKnop.ToolTipText = "Toont de opgeslagen fakturen."
   OpslaanKnop.ToolTipText = "Slaat dit faktuur op."
   ParticulierVeld.ToolTipText = "Geeft aan of het faktuur voor een particulier bestemd is."
   PortoVeld.ToolTipText = "Het porto bedrag."
   ToevoegenKnop.ToolTipText = "Voegt een artikel toe."
   VerwijderenKnop.ToolTipText = "Verwijdert een artikel."
   WijzigenKnop.ToolTipText = "Wijzigt een artikel."

   Me.Width = DepotbeheerderVenster.Width / 1.3
   Me.Height = DepotbeheerderVenster.Height / 1.15
   Me.Left = MenuVenster.Left + MenuVenster.Width + 128
   Me.Top = MenuVenster.Top

   Set ArtikelenLijst.DataBron = Faktuur
   Faktuur.MaakNieuwFaktuur

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
   ArtikelenLijst.Height = Me.ScaleHeight - 20.5
   FaktuurVelden.Left = Me.ScaleWidth - FaktuurVelden.Width - 2
   KnoppenBalk.Left = Me.ScaleWidth - KnoppenBalk.Width - 1
   KnoppenBalk.Top = Me.ScaleHeight - 8.5
   TotaalVelden.Left = Me.ScaleWidth - TotaalVelden.Width - 2
   TotaalVelden.Top = Me.ScaleHeight - 8

   ArtikelenLijst.WerkLijstBij
End Sub

'Deze procedure legt de ingevoerde waarde vast.
Private Sub KlantAdresVeld_LostFocus()
On Error GoTo Fout
Dim Keuze As Long

   Faktuur.KlantAdres = KlantAdresVeld.Text

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waarde vast.
Private Sub KlantEmailadresVeld_LostFocus()
On Error GoTo Fout
Dim Keuze As Long

   Faktuur.KlantEmailadres = KlantEmailadresVeld.Text

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waarde vast.
Private Sub KlantFaxNummerVeld_LostFocus()
On Error GoTo Fout
Dim Keuze As Long

   Faktuur.KlantFaxNummer = KlantFaxnummerVeld.Text

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waarde vast.
Private Sub KlantnaamVeld_LostFocus()
On Error GoTo Fout
Dim Keuze As Long

   Faktuur.Klantnaam = KlantnaamVeld.Text

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub KlantnummerVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   KlantnummerVeld.SelStart = 0
   KlantnummerVeld.SelLength = Len(KlantnummerVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure filtert ongeldige tekens uit de invoer.
Private Sub KlantnummerVeld_KeyPress(KeyAscii As Integer)
On Error GoTo Fout
Dim Keuze As Long

   If InStr(ONGELDIGE_TEKENS, Chr$(KeyAscii)) > 0 Then KeyAscii = Empty

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waarde vast en past het faktuur aan.
Private Sub KlantnummerVeld_LostFocus()
On Error GoTo Fout
Dim Keuze As Long

   Faktuur.Klantnummer = KlantnummerVeld.Text
   WerkFaktuurBij

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waarde vast.
Private Sub KlantPostcodePlaatsVeld_LostFocus()
On Error GoTo Fout
Dim Keuze As Long

   Faktuur.KlantPostcodePlaats = KlantPostcodePlaatsVeld.Text

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waarde vast.
Private Sub KlantTelefoonnummerVeld_LostFocus()
On Error GoTo Fout
Dim Keuze As Long

   Faktuur.KlantTelefoonnummer = KlantTelefoonnummerVeld.Text

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure opent na bevestiging van de gebruiker een nieuwe faktuur.
Private Sub NieuwKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   If MsgBox("Nieuw faktuur openen?", vbQuestion Or vbYesNo) = vbYes Then
      Faktuur.MaakNieuwFaktuur
      WerkFaktuurBij
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure toont het fakturen venster.
Private Sub OpenenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   FakturenVenster.Show
   FakturenVenster.ZOrder

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om dit faktuur op te slaan.
Private Sub OpslaanKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   If Faktuur.Faktuurnr() = vbNullString Then
      MsgBox "Faktuur heeft geen nummer.", vbExclamation
   Else
      Faktuur.SlaFaktuurOp
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt vast of het faktuur voor een particulier is en past het faktuur aan.
Private Sub ParticulierVeld_Click()
On Error GoTo Fout
Dim Keuze As Long

   Faktuur.IsParticulier = (ParticulierVeld.Value = vbChecked)
   WerkFaktuurBij

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub PortoVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   PortoVeld.SelStart = 0
   PortoVeld.SelLength = Len(PortoVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waarde vast en past het faktuur aan.
Private Sub PortoVeld_LostFocus()
On Error GoTo Fout
Dim Keuze As Long

   PortoVeld.Text = RondAf(PortoVeld.Text)
   Faktuur.Porto = PortoVeld.Text

   WerkTotalenBij

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om een artikel aan de geopende faktuur toe te voegen.
Private Sub ToevoegenKnop_Click()
On Error GoTo Fout
Dim Gevonden As Boolean
Dim Keuze As Long

   ActieIsToeVoegen = True
   Faktuur.StelStandaardWaardesIn
   FaktuurInvoerVenster.Show vbModal

   If Not Buffer(FkArtikelNr) = vbNullString Then
      Faktuur.VoegItemToe
      ArtikelenLijst.Selectie = Faktuur.AantalItems() - 1
      Faktuur.WijzigItem ArtikelenLijst.Selectie, Gevonden
      If Not Gevonden Then MsgBox "Artikel bestaat niet.", vbExclamation
   End If

   Faktuur.SorteerArtikelen
   ArtikelenLijst.WerkLijstBij
   WerkTotalenBij

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure verwijdert na bevesting van de gebruiker het artikel uit de geopende faktuur.
Private Sub VerwijderenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   If MsgBox("Artikel verwijderen?", vbQuestion Or vbYesNo) = vbYes Then
      Faktuur.VerwijderItem ArtikelenLijst.Selectie()
      ArtikelenLijst.WerkLijstBij
      WerkTotalenBij
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om een artikel in de geopende faktuur toe te wijzigen.
Private Sub WijzigenKnop_Click()
On Error GoTo Fout
Dim Gevonden As Boolean
Dim Keuze As Long

   If Not Faktuur.AantalItems() = 0 Then
      Buffer(FkArtikelNr) = Faktuur.Data(FkArtikelNr, ArtikelenLijst.Selectie())
      Buffer(FkAantal) = Faktuur.Data(FkAantal, ArtikelenLijst.Selectie())
      Buffer(FkAtikelNaam) = Faktuur.Data(FkAtikelNaam, ArtikelenLijst.Selectie())
      Buffer(FkStukprijs) = Faktuur.Data(FkStukprijs, ArtikelenLijst.Selectie())
      Buffer(FkUitVoorraad) = Faktuur.Data(FkUitVoorraad, ArtikelenLijst.Selectie())

      ActieIsToeVoegen = False
      FaktuurInvoerVenster.Show vbModal

      If Not Buffer(FkArtikelNr) = vbNullString Then
         Faktuur.WijzigItem ArtikelenLijst.Selectie(), Gevonden
         If Not Gevonden Then MsgBox "Artikel bestaat niet.", vbExclamation
      End If

      Faktuur.SorteerArtikelen
      ArtikelenLijst.WerkLijstBij
      WerkTotalenBij
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

