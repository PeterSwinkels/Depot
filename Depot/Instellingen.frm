VERSION 5.00
Begin VB.Form InstellingenVenster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instellingen"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   ClipControls    =   0   'False
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   23.125
   ScaleMode       =   4  'Character
   ScaleWidth      =   42.25
   Begin VB.TextBox LaatsteFaktuurdatumVeld 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      MaxLength       =   8
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton PrinterKnop 
      Caption         =   "&Printer"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton FaktuurtekstKnop 
      Caption         =   "&Faktuurtekst"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox NieuweFaktuurSubnrVeld 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      MaxLength       =   5
      TabIndex        =   8
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox FaktuurMaxSubNrsVeld 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      MaxLength       =   5
      TabIndex        =   7
      Top             =   2760
      Width           =   855
   End
   Begin VB.ComboBox FaktuurSubNrPeriodeVeld 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Instellingen.frx":0000
      Left            =   2520
      List            =   "Instellingen.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox FaktuurnrTekstVeld 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      MaxLength       =   250
      TabIndex        =   5
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox FaktuurNrOpmaakVeld 
      Height          =   285
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox DatumOpmaakVeld 
      Height          =   285
      Left            =   2520
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton AnnulerenKnop 
      Cancel          =   -1  'True
      Caption         =   "&Annuleren"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton WijzigenKnop 
      Caption         =   "&Wijzigen"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox NieuwWachtwoordVeld 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   250
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox HuidigWachtwoordVeld 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   250
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox DepotNaamVeld 
      Height          =   285
      Left            =   2520
      MaxLength       =   250
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label LaatsteFaktuurdatumLabel 
      Caption         =   "Laatste faktuurdatum:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label OverigeLabel 
      Caption         =   "Overige:"
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
      TabIndex        =   25
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label NieuweFaktuurnummerLabel 
      Caption         =   "Nieuw faktuurnummer:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "is:"
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label HetMaximumAantalFakturenPerLabel 
      Caption         =   "Het maximum aantal fakturen per"
      Height          =   240
      Left            =   120
      TabIndex        =   22
      Top             =   2760
      Width           =   2400
   End
   Begin VB.Label Label4 
      Caption         =   "Fakturen:"
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
      TabIndex        =   21
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Beheerder:"
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
      TabIndex        =   20
      Top             =   120
      Width           =   975
   End
   Begin VB.Label FaktuurnummertekstLabel 
      Caption         =   "Faktuurnummertekst:"
      Height          =   240
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   2400
   End
   Begin VB.Label FaktuurnummeropmaakLabel 
      Caption         =   "Faktuurnummeropmaak:"
      Height          =   240
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   2400
   End
   Begin VB.Label DatumopmaakLabel 
      Caption         =   "Datumopmaak:"
      Height          =   240
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   2400
   End
   Begin VB.Label NieuwBeheerderwachtwoordLabel 
      Caption         =   "Nieuw beheerderwachtwoord:"
      Height          =   240
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   2400
   End
   Begin VB.Label HuidigBeheerderwachtwoordLabel 
      Caption         =   "Huidig beheerderwachtwoord:"
      Height          =   240
      Left            =   120
      TabIndex        =   15
      Top             =   360
      Width           =   2400
   End
   Begin VB.Label DepotnaamLabel 
      Caption         =   "Depotnaam:"
      Height          =   240
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   2400
   End
End
Attribute VB_Name = "InstellingenVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het venster waarin instellingen aangepast kunnen worden.
Option Explicit

'Deze procedure legt de ingevoerde waardes vast.
Private Sub HaalInvoer()
On Error GoTo Fout
Dim Keuze As Long

   Faktuur.DatumOpmaak = DatumOpmaakVeld.Text
   Faktuur.DepotNaam = DepotNaamVeld.Text
   Faktuur.MaxSubNrs = Trim$(Str$(MaakGetalOp(FaktuurMaxSubNrsVeld.Text) - 1))
   Faktuur.NrOpmaak = FaktuurNrOpmaakVeld.Text
   FaktuurnrTekstVeld.Text = UCase$(FaktuurnrTekstVeld.Text)
   Faktuur.SubNrTekst = FaktuurnrTekstVeld.Text
   Faktuur.LaatsteDatum = LaatsteFaktuurdatumVeld.Text
   Faktuur.LaatsteSubNr = Trim$(Str$(MaakGetalOp(NieuweFaktuurSubnrVeld.Text) - 1))
   IngevoerdWachtwoord = ZetBitsOm(HuidigWachtwoordVeld.Text)

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

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub DatumOpmaakVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   DatumOpmaakVeld.SelStart = 0
   DatumOpmaakVeld.SelLength = Len(DatumOpmaakVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure verwijdert ongeldige tekens uit de invoer.
Private Sub DatumOpmaakVeld_KeyPress(KeyAscii As Integer)
On Error GoTo Fout
Dim Keuze As Long

   If InStr("DJjM" & Chr$(vbKeyBack), Chr$(KeyAscii)) = 0 Then KeyAscii = Empty
   If InStr(UCase$(DatumOpmaakVeld.Text), UCase$(Chr$(KeyAscii))) > 0 Then KeyAscii = Empty

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waarde vast.
Private Sub DatumOpmaakVeld_LostFocus()
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
Private Sub DepotNaamVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   DepotNaamVeld.SelStart = 0
   DepotNaamVeld.SelLength = Len(DepotNaamVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waarde vast.
Private Sub DepotNaamVeld_LostFocus()
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
Private Sub FaktuurMaxSubNrsVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   FaktuurMaxSubNrsVeld.SelStart = 0
   FaktuurMaxSubNrsVeld.SelLength = Len(FaktuurMaxSubNrsVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waarde vast.
Private Sub FaktuurMaxSubNrsVeld_LostFocus()
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

'Deze procedure past de invoer velden aan de faktuur nummer opmaak aan.
Private Sub FaktuurNrOpmaakVeld_Change()
On Error GoTo Fout
Dim Keuze As Long

   FaktuurMaxSubNrsVeld.Enabled = (InStr(FaktuurNrOpmaakVeld.Text, "#") > 0)
   FaktuurnrTekstVeld.Enabled = (InStr(FaktuurNrOpmaakVeld.Text, "T") > 0)
   FaktuurSubNrPeriodeVeld.Enabled = (InStr(FaktuurNrOpmaakVeld.Text, "#") > 0)
   LaatsteFaktuurdatumVeld.Enabled = (InStr(FaktuurNrOpmaakVeld.Text, "#") > 0)
   NieuweFaktuurSubnrVeld.Enabled = (InStr(FaktuurNrOpmaakVeld.Text, "#") > 0)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub FaktuurNrOpmaakVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   FaktuurNrOpmaakVeld.SelStart = 0
   FaktuurNrOpmaakVeld.SelLength = Len(FaktuurNrOpmaakVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure verwijdert ongeldige tekens uit de invoer.
Private Sub FaktuurNrOpmaakVeld_KeyPress(KeyAscii As Integer)
On Error GoTo Fout
Dim Keuze As Long

   If InStr("DJjKMT#" & Chr$(vbKeyBack), Chr$(KeyAscii)) = 0 Then KeyAscii = Empty
   If InStr(UCase$(FaktuurNrOpmaakVeld.Text), UCase$(Chr$(KeyAscii))) > 0 Then KeyAscii = Empty

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waarde vast.
Private Sub FaktuurNrOpmaakVeld_LostFocus()
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
Private Sub FaktuurNrTekstVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   FaktuurnrTekstVeld.SelStart = 0
   FaktuurnrTekstVeld.SelLength = Len(FaktuurnrTekstVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure filtert ongeldige tekens uit de invoer.
Private Sub FaktuurNrTekstVeld_KeyPress(KeyAscii As Integer)
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
Private Sub FaktuurNrTekstVeld_LostFocus()
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

'Deze procedure legt de ingevoerde waarde vast.
Private Sub FaktuurSubNrPeriodeVeld_Click()
On Error GoTo Fout
Dim Keuze As Long

   Faktuur.SubNrPeriode = Mid$("DMJ", FaktuurSubNrPeriodeVeld.ListIndex + 1, 1)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure toont het faktuur tekst venster.
Private Sub FaktuurTekstKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   FaktuurtekstVenster.Show
   FaktuurtekstVenster.ZOrder

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure past dit venster aan wanneer het wordt geactiveerd.
Private Sub Form_Activate()
On Error GoTo Fout
Dim Keuze As Long

   AnnulerenKnop.ToolTipText = "Sluit dit venster."
   DatumOpmaakVeld.ToolTipText = "De faktuur datum opmaak."
   DepotNaamVeld.ToolTipText = "De naam van het depot dat op het faktuur wordt afgedrukt."
   FaktuurMaxSubNrsVeld.ToolTipText = "Het maximum aantal faktuursubnummers per periode."
   FaktuurSubNrPeriodeVeld.ToolTipText = "De faktuursubnummer periode."
   FaktuurNrOpmaakVeld.ToolTipText = "De faktuurnummer opmaak."
   FaktuurnrTekstVeld.ToolTipText = "Deze tekst wordt aan het faktuur nummer toegevoegd."
   FaktuurTekstKnop.ToolTipText = "Opent het faktuur tekst venster."
   HuidigWachtwoordVeld.ToolTipText = "Het huidige beheerder wachtwoord."
   LaatsteFaktuurdatumVeld.ToolTipText = "De datum van de laatst opgeslagen faktuur."
   NieuweFaktuurSubnrVeld.ToolTipText = "Het eerst volgende nieuwe faktuur nummer."
   NieuwWachtwoordVeld.ToolTipText = "Het nieuwe beheerder wachtwoord."
   PrinterKnop.ToolTipText = "Opent het printer instellingen venster."
   WijzigenKnop.ToolTipText = "Wijzigt de instellingen."

   If Faktuur.LaatsteSubNr() = vbNullString Then
      NieuweFaktuurSubnrVeld.Text = "0"
   Else
      NieuweFaktuurSubnrVeld.Text = CStr(Val(Faktuur.LaatsteSubNr()) + 1)
   End If

   LaatsteFaktuurdatumVeld.Text = Faktuur.LaatsteDatum()

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

   Me.Left = (DepotbeheerderVenster.Width / 2) - (Me.Width / 2)
   Me.Top = (DepotbeheerderVenster.Height / 2.5) - (Me.Height / 2)

   DatumOpmaakVeld.Text = Faktuur.DatumOpmaak()
   DepotNaamVeld.Text = Faktuur.DepotNaam()
   FaktuurMaxSubNrsVeld.Text = Faktuur.MaxSubNrs() + 1
   FaktuurNrOpmaakVeld.Text = Faktuur.NrOpmaak()
   FaktuurnrTekstVeld.Text = Faktuur.SubNrTekst()
   FaktuurSubNrPeriodeVeld.ListIndex = InStr("DMJ", Faktuur.SubNrPeriode()) - 1

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub HuidigWachtwoordVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   HuidigWachtwoordVeld.SelStart = 0
   HuidigWachtwoordVeld.SelLength = Len(HuidigWachtwoordVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure verwijdert ongeldige tekens uit de invoer.
Private Sub LaatsteFaktuurDatumVeld_KeyPress(KeyAscii As Integer)
On Error GoTo Fout
Dim Keuze As Long

   If InStr("0123456789" & Chr$(vbKeyBack), Chr$(KeyAscii)) = 0 Then KeyAscii = Empty

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub NieuweFaktuurSubNrVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   NieuweFaktuurSubnrVeld.SelStart = 0
   NieuweFaktuurSubnrVeld.SelLength = Len(NieuweFaktuurSubnrVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waarde vast.
Private Sub NieuweFaktuurSubNrVeld_LostFocus()
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
Private Sub NieuwWachtwoordVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   NieuwWachtwoordVeld.SelStart = 0
   NieuwWachtwoordVeld.SelLength = Len(NieuwWachtwoordVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure toont het printer venster.
Private Sub PrinterKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   PrinterVenster.Show
   PrinterVenster.ZOrder

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om de instellingen te bewaren en sluit dit venster.
Private Sub WijzigenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   HaalInvoer
   StelWachtwoordIn NieuwWachtwoordVeld.Text
   SlaInstellingenOp
   Unload Me

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

