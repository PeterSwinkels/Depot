VERSION 5.00
Begin VB.Form KlantenVenster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Klanten"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   ClipControls    =   0   'False
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   15.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   51.25
   Begin VB.CommandButton ToevoegenKnop 
      Caption         =   "&Toevoegen"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton VerwijderKnop 
      Caption         =   "&Verwijderen"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton WijzigenKnop 
      Caption         =   "&Wijzigen"
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton AfdrukkenKnop 
      Caption         =   "&Afdrukken"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.PictureBox KlantVelden 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   5895
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   5895
      Begin VB.TextBox KlantTelefoonnummerVeld 
         Height          =   285
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   4
         Top             =   1440
         Width           =   4455
      End
      Begin VB.TextBox KlantPostcodePlaatsVeld 
         Height          =   285
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   3
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox KlantAdresVeld 
         Height          =   285
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   2
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox KlantnaamVeld 
         Height          =   285
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   1
         Top             =   360
         Width           =   4455
      End
      Begin VB.TextBox KlantnummerVeld 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   0
         Top             =   0
         Width           =   4455
      End
      Begin VB.TextBox KlantEmailadresVeld 
         Height          =   285
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   6
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox KlantFaxnummerVeld 
         Height          =   285
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   5
         Top             =   1800
         Width           =   4455
      End
      Begin VB.Label TelefoonnummerLabel 
         Caption         =   "Telefoonnummer:"
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label PostcodePlaatsLabel 
         Caption         =   "Postcode/Plaats:"
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label AdresLabel 
         Caption         =   "Adres:"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label NaamLabel 
         Caption         =   "Naam:"
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label NummerLabel 
         Caption         =   "Nummer:"
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label EMailadresLabel 
         Caption         =   "E-mailadres:"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label FaxnummerLabel 
         Caption         =   "Faxnummer:"
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   1800
         Width           =   1335
      End
   End
   Begin VB.CommandButton VolgendeKnop 
      Caption         =   "Vo&lgende"
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton VorigeKnop 
      Caption         =   "V&orige"
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "KlantenVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het klanten venster.
Option Explicit
Private SelectieV As Long     'Bevat de geselecteerde klant.

'Deze procedure legt de ingevoerde waardes vast.
Private Sub HaalInvoer()
On Error GoTo Fout
Dim Keuze As Long

   Buffer(KlNummer) = KlantnummerVeld.Text
   Buffer(KlNaam) = KlantnaamVeld.Text
   Buffer(KlAdres) = KlantAdresVeld.Text
   Buffer(KlPostcodePlaats) = KlantPostcodePlaatsVeld.Text
   Buffer(KlTelefoonNummer) = KlantTelefoonnummerVeld.Text
   Buffer(KlFaxNummer) = KlantFaxnummerVeld.Text
   Buffer(KlEmailadres) = KlantEmailAdresVeld.Text

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stuurt de selectie terug.
Public Property Get Selectie() As Long
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Selectie = SelectieV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure selecteert de opgegeven klant.
Public Property Let Selectie(NieuweSelectie As Long)
On Error GoTo Fout
Dim Keuze As Long

   SelectieV = NieuweSelectie
   WerkVeldenBij

EindeRoutine:
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure werkt de klant gegevens invoer velden bij.
Private Sub WerkVeldenBij()
On Error GoTo Fout
Dim Keuze As Long

   If SelectieV > Klanten.AantalItems() Then SelectieV = Klanten.AantalItems()
   Me.Caption = "Klanten - Klant: " & SelectieV & "/" & Klanten.AantalItems()

   If SelectieV = Klanten.AantalItems() Then
      KlantnummerVeld.Text = vbNullString
      KlantnaamVeld.Text = vbNullString
      KlantAdresVeld.Text = vbNullString
      KlantPostcodePlaatsVeld.Text = vbNullString
      KlantTelefoonnummerVeld.Text = vbNullString
      KlantFaxnummerVeld.Text = vbNullString
      KlantEmailAdresVeld.Text = vbNullString
   Else
      KlantnummerVeld.Text = Klanten.Data(KlNummer, SelectieV)
      KlantnaamVeld.Text = Klanten.Data(KlNaam, SelectieV)
      KlantAdresVeld.Text = Klanten.Data(KlAdres, SelectieV)
      KlantPostcodePlaatsVeld.Text = Klanten.Data(KlPostcodePlaats, SelectieV)
      KlantTelefoonnummerVeld.Text = Klanten.Data(KlTelefoonNummer, SelectieV)
      KlantFaxnummerVeld.Text = Klanten.Data(KlFaxNummer, SelectieV)
      KlantEmailAdresVeld.Text = Klanten.Data(KlEmailadres, SelectieV)
   End If

   HaalInvoer

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure drukt de klant gegevens af na bevestiging van de gebruiker.
Private Sub AfdrukkenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   If MsgBox("Klant gegevens afdrukken?", vbQuestion Or vbYesNo) = vbYes Then
      DrukDataAf Klanten
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

   AfdrukkenKnop.ToolTipText = "Drukt de gegevens van alle klanten af."
   KlantAdresVeld.ToolTipText = "Het klant adres."
   KlantEmailAdresVeld.ToolTipText = "Het klant e-mailadres."
   KlantFaxnummerVeld.ToolTipText = "Het klant faxnummer."
   KlantnaamVeld.ToolTipText = "De klant naam."
   KlantnummerVeld.ToolTipText = "Het klantnummer."
   KlantPostcodePlaatsVeld.ToolTipText = "De klant postcode en plaats."
   KlantTelefoonnummerVeld.ToolTipText = "Het klant telefoonnummer."
   ToevoegenKnop.ToolTipText = "Voegt een klant toe."
   VerwijderKnop.ToolTipText = "Verwijdert een klant."
   VolgendeKnop.ToolTipText = "Toont de volgende klant."
   VorigeKnop.ToolTipText = "Toont de vorige klant."
   WijzigenKnop.ToolTipText = "Wijzigt een klant."

   Me.Left = MenuVenster.Left + MenuVenster.Width + 128
   Me.Top = MenuVenster.Top

   SelectieV = 0
   WerkVeldenBij

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure werkt de klanten lijst bij wanneer dit venster wordt gesloten.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Fout
Dim Keuze As Long

   HaalInvoer

   If Not Buffer(KlNummer) = vbNullString Then
      If SelectieV = Klanten.AantalItems() Then Klanten.VoegItemToe
      Klanten.WijzigItem SelectieV
      Klanten.SorteerItems
      Klanten.VerwijderOudeKlant
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub KlantAdresVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   KlantAdresVeld.SelStart = 0
   KlantAdresVeld.SelLength = Len(KlantAdresVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waarde vast.
Private Sub KlantAdresVeld_LostFocus()
On Error GoTo Fout
Dim Keuze As Long

   HaalInvoer

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub KlantEmailAdresVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   KlantEmailAdresVeld.SelStart = 0
   KlantEmailAdresVeld.SelLength = Len(KlantEmailAdresVeld.Text)

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

   HaalInvoer

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub KlantFaxnummerVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   KlantFaxnummerVeld.SelStart = 0
   KlantFaxnummerVeld.SelLength = Len(KlantFaxnummerVeld.Text)

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

   HaalInvoer

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub KlantnaamVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   KlantnaamVeld.SelStart = 0
   KlantnaamVeld.SelLength = Len(KlantnaamVeld.Text)

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

   HaalInvoer

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

'Deze procedure verwijdert ongeldige tekens uit de invoer.
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

'Deze procedure legt de ingevoerde waarde vast.
Private Sub KlantnummerVeld_LostFocus()
On Error GoTo Fout
Dim Keuze As Long

   HaalInvoer

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub KlantPostcodePlaatsVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   KlantPostcodePlaatsVeld.SelStart = 0
   KlantPostcodePlaatsVeld.SelLength = Len(KlantPostcodePlaatsVeld.Text)

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

   HaalInvoer

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub KlantTelefoonnummerVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   KlantTelefoonnummerVeld.SelStart = 0
   KlantTelefoonnummerVeld.SelLength = Len(KlantTelefoonnummerVeld.Text)

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

   HaalInvoer

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure voegt een nieuwe klant aan de lijst toe.
Private Sub ToevoegenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   KlantVelden.Enabled = False
   If Not Buffer(KlNummer) = vbNullString Then
      If SelectieV = Klanten.AantalItems() Then Klanten.VoegItemToe
      Klanten.WijzigItem SelectieV
      Klanten.SorteerItems
      Klanten.VerwijderOudeKlant
      SelectieV = Klanten.AantalItems()
      WerkVeldenBij
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure verwijdert een klant uit de lijst na bevestiging van de gebruiker.
Private Sub VerwijderKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   If MsgBox("Klant verwijderen?", vbQuestion Or vbYesNo) = vbYes Then
      KlantVelden.Enabled = False
      If SelectieV < Klanten.AantalItems() Then
         Klanten.VerwijderItem SelectieV
         WerkVeldenBij
      End If
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert de volgende klant uit de lijst.
Private Sub VolgendeKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   KlantVelden.Enabled = False
   If SelectieV < Klanten.AantalItems() Then
      If Not Buffer(KlNummer) = vbNullString Then
         Klanten.WijzigItem SelectieV
         Klanten.SorteerItems
         Klanten.VerwijderOudeKlant
      End If
      SelectieV = SelectieV + 1
      WerkVeldenBij
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert de vorige klant uit de lijst.
Private Sub VorigeKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   KlantVelden.Enabled = False
   If SelectieV > 0 Then
      If Not Buffer(KlNummer) = vbNullString Then
         If SelectieV = Klanten.AantalItems() Then Klanten.VoegItemToe
         Klanten.WijzigItem SelectieV
         Klanten.SorteerItems
         Klanten.VerwijderOudeKlant
      End If
      SelectieV = SelectieV - 1
      WerkVeldenBij
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de invoer velden vrij voor invoer.
Private Sub WijzigenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   KlantVelden.Enabled = True
   KlantnummerVeld.SetFocus

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

