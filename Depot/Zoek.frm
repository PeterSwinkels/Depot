VERSION 5.00
Begin VB.Form ZoekVenster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zoeken"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10.563
   ScaleMode       =   4  'Character
   ScaleWidth      =   25.125
   Begin VB.CheckBox HoofdlettergevoeligVeld 
      Caption         =   "&Hoofdlettergevoelig"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.ComboBox BronVeld 
      Height          =   315
      ItemData        =   "Zoek.frx":0000
      Left            =   1560
      List            =   "Zoek.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton ZoekVooruitKnop 
      Caption         =   "Zoek &Vooruit"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton ZoekAchteruitKnop 
      Caption         =   "Zoek &Achteruit"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox ZoektekstVeld 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.CheckBox ZoekHeleVeldVeld 
      Caption         =   "&Zoek Hele Veld"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ComboBox DataBronVeld 
      Height          =   315
      ItemData        =   "Zoek.frx":0004
      Left            =   120
      List            =   "Zoek.frx":0014
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label ZoekInLabel 
      Caption         =   "Zoek in:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label ZoektekstLabel 
      Caption         =   "Zoektekst:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "ZoekVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze procedure toont het zoek venster.
Option Explicit
Private Databron As Object
Private Gevonden As Boolean

'Deze procedure zoekt naar een item die voldoet aan de geselecteerde zoekcriteria.
Private Sub Zoek(Richting As ZoekrichtingE)
On Error GoTo Fout
Dim Keuze As Long
Dim SelectieV As Long

   Screen.MousePointer = vbHourglass

   Select Case DataBronVeld.ListIndex
      Case DbFakturen
         SelectieV = FakturenVenster.FakturenLijst.Selectie()
      Case DbFaktuur
         SelectieV = FaktuurVenster.ArtikelenLijst.Selectie()
      Case DbKlanten
         SelectieV = KlantenVenster.Selectie()
      Case DbVoorraad
         SelectieV = VoorraadVenster.ArtikelenLijst.Selectie()
   End Select

   SelectieV = ZoekItemIndex(Databron, SelectieV, BronVeld.ListIndex, ZoektekstVeld.Text, (ZoekHeleVeldVeld.Value = vbChecked), (HoofdlettergevoeligVeld.Value = vbChecked), Richting, Gevonden)

   If Gevonden Then
      Select Case DataBronVeld.ListIndex
         Case DbFakturen
            FakturenVenster.Show
            FakturenVenster.ZOrder
            FakturenVenster.FakturenLijst.BovensteItem = SelectieV
            FakturenVenster.FakturenLijst.Selectie = SelectieV
         Case DbFaktuur
            FaktuurVenster.Show
            FaktuurVenster.ZOrder
            FaktuurVenster.ArtikelenLijst.BovensteItem = SelectieV
            FaktuurVenster.ArtikelenLijst.Selectie = SelectieV
         Case DbKlanten
            KlantenVenster.Show
            KlantenVenster.ZOrder
            KlantenVenster.Selectie = SelectieV
         Case DbVoorraad
            VoorraadVenster.Show
            VoorraadVenster.ZOrder
            VoorraadVenster.ArtikelenLijst.BovensteItem = SelectieV
            VoorraadVenster.ArtikelenLijst.Selectie = SelectieV
      End Select

      Me.ZOrder
   ElseIf Databron.AantalItems() > 0 Then
      MsgBox "Tekst niet gevonden.", vbInformation
   End If

EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stelt de zoekresultaten opnieuw in wanneer een databron wordt geselecteerd.
Private Sub BronVeld_Click()
On Error GoTo Fout
Dim Keuze As Long

   Gevonden = False

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert een databron.
Private Sub DataBronVeld_Click()
On Error GoTo Fout
Dim Keuze As Long
Dim Veld As Long

   Gevonden = False

   Select Case DataBronVeld.ListIndex
      Case DbFakturen
         Set Databron = Fakturen
      Case DbFaktuur
         Set Databron = Faktuur
      Case DbKlanten
         Set Databron = Klanten
      Case DbVoorraad
         Set Databron = Voorraad
   End Select

   BronVeld.Clear
   For Veld = 0 To Databron.AantalVelden - 1
      BronVeld.AddItem Databron.VeldNaam(Veld)
   Next Veld
   BronVeld.ListIndex = 0

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

   BronVeld.ToolTipText = "Het te doorzoeken veld."
   DataBronVeld.ToolTipText = "De te doorzoeken databron."
   HoofdlettergevoeligVeld.ToolTipText = "Geeft aan of de zoekopdracht hoofdletter gevoelig is."
   ZoekAchteruitKnop.ToolTipText = "Zoekt achteruit."
   ZoekHeleVeldVeld.ToolTipText = "Geeft aan of het hele veld overeen moet komen met de zoektekst."
   ZoektekstVeld.ToolTipText = "De te zoeken tekst."
   ZoekVooruitKnop.ToolTipText = "Zoekt vooruit."

   Me.Left = MenuVenster.Left + MenuVenster.Width + 128
   Me.Top = MenuVenster.Top

   DataBronVeld.ListIndex = DbFakturen
   Gevonden = False

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stelt de zoekresultaten opnieuw in wanneer de zoekcriteria worden gewijzigd.
Private Sub HoofdlettergevoeligVeld_Click()
On Error GoTo Fout
Dim Keuze As Long

   Gevonden = False

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht achteruit te zoeken.
Private Sub ZoekAchteruitKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   Zoek ZrAchteruit

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stelt de zoekresultaten opnieuw in wanneer de zoekcriteria worden gewijzigd.
Private Sub ZoekHeleVeldVeld_Click()
On Error GoTo Fout
Dim Keuze As Long

   Gevonden = False

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stelt de zoekresultaten opnieuw in wanneer de zoektekst wordt gewijzigd.
Private Sub ZoektekstVeld_Change()
On Error GoTo Fout
Dim Keuze As Long

   Gevonden = False

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub ZoektekstVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   ZoektekstVeld.SelStart = 0
   ZoektekstVeld.SelLength = Len(ZoektekstVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht vooruit te zoeken.
Private Sub ZoekVooruitKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   Zoek ZrVooruit

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

