VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KlantenObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Deze module bevat het klantenobject.
Option Explicit

'Deze opsomming definieert de beschikbare klantenvelden.
Public Enum KlantenE
   KlNummer             'Definieert een klantnummer.
   KlNaam               'Definieert een klantnaam.
   KlAdres              'Definieert een klant adres.
   KlPostcodePlaats     'Definieert een klant postcode en plaatsnaam.
   KlTelefoonnummer     'Definieert een klant telefoonnummer.
   KlFaxnummer          'Definieert een klant faxnummer.
   KlEmailadres         'Definieert een klant e-mailadres.
End Enum

Private AantalItemsV As Long      'Bevat het aantal klanten in de lijst.
Private AfdrukkenV As Byte        'Geeft aan welke velden worden afgedrukt.
Private DataV() As String         'Bevat de velden met gegevens.
Private LeesFout As Boolean       'Geeft aan of er een fout is opgetreden tijdens het lezen van het klanten bestand.

'Deze procedure stuurt het aantal klanten terug.
Public Property Get AantalItems() As Long
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   AantalItems = AantalItemsV
   Exit Property

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het aantal velden voor dit object terug.
Public Property Get AantalVelden() As Long
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   AantalVelden = 7
   Exit Property

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt welke velden worden afgedrukt terug.
Public Property Get Afdrukken() As Byte
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Afdrukken = AfdrukkenV
   Exit Property

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze eigenschap procedure legt welke velden worden afgedrukt vast.
Public Property Let Afdrukken(NieuwAfdrukken As Byte)
On Error GoTo Fout
Dim Keuze As Long

   AfdrukkenV = NieuwAfdrukken

EindeRoutine:
   Exit Property

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de gegevens van het opgegeven veld van de opgegeven klant terug.
Public Property Get Data(Veld As Long, Klant As Long) As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Data = DataV(Veld, Klant)
   Exit Property

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de gegevens van het opgegeven veld van de opgegeven klant vast.
Public Property Let Data(Veld As Long, Klant As Long, NieuweData As String)
On Error GoTo Fout
Dim Keuze As Long

   DataV(Veld, Klant) = NieuweData

EindeRoutine:
   Exit Property

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure maakt de klantgegevens op.
Public Sub MaakItemsOp()
On Error GoTo Fout
Dim Keuze As Long
Dim Klant As Long

   Screen.MousePointer = vbHourglass
   For Klant = 0 To AantalItemsV - 1
      DataV(KlNummer, Klant) = Trim$(UCase$(DataV(KlNummer, Klant)))
      DataV(KlNaam, Klant) = Trim$(DataV(KlNaam, Klant))
      DataV(KlAdres, Klant) = Trim$(DataV(KlAdres, Klant))
      DataV(KlPostcodePlaats, Klant) = Trim$(DataV(KlPostcodePlaats, Klant))
      DataV(KlTelefoonnummer, Klant) = Trim$(DataV(KlTelefoonnummer, Klant))
      DataV(KlFaxnummer, Klant) = Trim$(DataV(KlFaxnummer, Klant))
      DataV(KlEmailadres, Klant) = Trim$(DataV(KlEmailadres, Klant))
   Next Klant

EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure opent de klantenlijst.
Public Sub OpenGegevens()
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim BestandH As Integer
Dim Keuze As Long
Dim Lengte As Long
Dim Veld As Long

   Screen.MousePointer = vbHourglass
   ActieveMap = ".\Data\"
   ActiefBestand = "Klanten.dat"
   BestandH = FreeFile()
   LeesFout = False
   Open ".\Data\Klanten.dat" For Binary Lock Read Write As BestandH
      Do Until Loc(BestandH) = LOF(BestandH)
         VoegItemToe
         For Veld = LBound(DataV(), 1) To UBound(DataV(), 1)
            Lengte = Asc(Input$(1, BestandH))
            DataV(Veld, AantalItemsV - 1) = Input$(Lengte, BestandH)
         Next Veld
      Loop
   Close BestandH

EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   LeesFout = True
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure maakt de klantenlijst leeg.
Public Sub ResetGegevens()
On Error GoTo Fout
Dim Keuze As Long

   AantalItemsV = 0
   ReDim DataV(0 To 6, 0 To AantalItemsV) As String

EindeRoutine:
   Exit Sub

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure slaat de klantenlijst op.
Public Sub SlaGegevensOp()
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim Afbreken As Boolean
Dim BestandH As Integer
Dim Keuze As Long
Dim Klant As Long
Dim Veld As Long

   ActieveMap = ".\Data\"
   ActiefBestand = "Klanten.dat"
   Afbreken = False
   If LeesFout Then
      Afbreken = BevestigAfbrekenNaLeesFout(ActiefBestand)
   End If

   If Not Afbreken Then
      Screen.MousePointer = vbHourglass
      BestandH = FreeFile()
      Open ".\Data\Klanten.dat" For Output Lock Read Write As BestandH
         For Klant = 0 To AantalItemsV - 1
            For Veld = LBound(DataV(), 1) To UBound(DataV(), 1)
               Print #BestandH, Chr$(Len(DataV(Veld, Klant))); DataV(Veld, Klant);
            Next Veld
         Next Klant
      Close BestandH
   End If

EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure sorteert de klanten op nummer.
Public Sub SorteerItems()
On Error GoTo Fout
Dim AndereKlant As Long
Dim Keuze As Long
Dim Klant As Long
Dim Veld As Long

   Screen.MousePointer = vbHourglass
   For Klant = 0 To AantalItemsV - 1
      For AndereKlant = 0 To AantalItemsV - 1
         If Not AndereKlant = Klant Then
            If DataV(KlNummer, AndereKlant) > DataV(KlNummer, Klant) Then
               For Veld = LBound(DataV(), 1) To UBound(DataV(), 1)
                  Verwissel DataV(Veld, AndereKlant), DataV(Veld, Klant)
               Next Veld
            End If
         End If
      Next AndereKlant
   Next Klant

EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure maakt de invoerbuffer leeg.
Public Sub StelStandaardWaardesIn()
On Error GoTo Fout
Dim Keuze As Long

   Erase Buffer

EindeRoutine:
   Exit Sub

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stuurt de totale breedte van alle velden terug.
Public Property Get TotaleVeldBreedte() As Long
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   TotaleVeldBreedte = 9
   Exit Property

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de breedte van het opgegeven veld in verhouding tot de andere velden terug.
Public Property Get VeldBreedte(Veld As Long) As Long
On Error GoTo Fout
Dim Breedte As Long
Dim Keuze As Long

   Select Case Veld
      Case KlNummer
         Breedte = 1
      Case KlNaam
         Breedte = 2
      Case KlAdres
         Breedte = 1
      Case KlPostcodePlaats
         Breedte = 1
      Case KlTelefoonnummer
         Breedte = 1
      Case KlFaxnummer
         Breedte = 1
      Case KlEmailadres
         Breedte = 2
      Case Else
         Breedte = 0
   End Select

EindeRoutine:
   VeldBreedte = Breedte
   Exit Property

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de naam van het opgegeven veld terug.
Public Property Get VeldNaam(Veld As Long) As String
On Error GoTo Fout
Dim Keuze As Long
Dim Naam As String

   Select Case Veld
      Case KlNummer
         Naam = "Nummer"
      Case KlNaam
         Naam = "Naam"
      Case KlAdres
         Naam = "Adres"
      Case KlPostcodePlaats
         Naam = "Postcode/Plaats"
      Case KlTelefoonnummer
         Naam = "Telefoonnummer"
      Case KlFaxnummer
         Naam = "Faxnummer"
      Case KlEmailadres
         Naam = "E-mailadres"
      Case Else
         Naam = vbNullString
   End Select

EindeRoutine:
   VeldNaam = Naam
   Exit Property

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt terug of het opgegeven veld rechtsuitgelijnd wordt.
Public Property Get VeldRechtsUitlijnen(Veld As Long) As Boolean
On Error GoTo Fout
Dim Keuze As Long
Dim RechtsUitlijnen As Boolean

   Select Case Veld
      Case KlNummer
         RechtsUitlijnen = True
      Case KlNaam
         RechtsUitlijnen = False
      Case KlAdres
         RechtsUitlijnen = False
      Case KlPostcodePlaats
         RechtsUitlijnen = False
      Case KlTelefoonnummer
         RechtsUitlijnen = False
      Case KlFaxnummer
         RechtsUitlijnen = False
      Case KlEmailadres
         RechtsUitlijnen = False
      Case Else
         RechtsUitlijnen = False
   End Select

EindeRoutine:
   VeldRechtsUitlijnen = RechtsUitlijnen
   Exit Property

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure verwijdert de opgegeven klant uit de lijst.
Public Sub VerwijderItem(Klant As Long)
On Error GoTo Fout
Dim Keuze As Long

   If Not AantalItemsV = 0 Then
      VerplaatsItem DataV(), Klant, AantalItemsV - 1
      AantalItemsV = AantalItemsV - 1
   End If

EindeRoutine:
   Exit Sub

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure verwijdert de klanten die voldoen aan de opgegeven voorwaarde.
Public Sub VerwijderItems(Voorwaarde As VerwijderVoorwaardenE)
On Error GoTo Fout
Dim Keuze As Long
Dim Klant As Long

   Screen.MousePointer = vbHourglass
   Klant = 0
   Select Case Voorwaarde
      Case VVDubbelNummer
         Do Until Klant > AantalItemsV - 1
            If Klant = DataV(KlNummer, Klant + 1) Then
               VerwijderItem Klant
               Klant = 0
            Else
               Klant = Klant + 1
            End If
         Loop
      Case VVGeenNummer
         Do Until Klant > AantalItemsV - 1
            If Klant = vbNullString Then
               VerwijderItem Klant
               Klant = 0
            Else
               Klant = Klant + 1
            End If
         Loop
   End Select

EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure verwijdert of overschrijft een klant.
Public Sub VerwijderOudeKlant()
On Error GoTo Fout
Dim Keuze As Long
Dim Klant As Long

   Screen.MousePointer = vbHourglass
   For Klant = 0 To AantalItemsV - 2
      If DataV(KlNummer, Klant) = DataV(KlNummer, Klant + 1) Then
         If MsgBox("Deze klant is al ingevoerd. Vervangen?", vbExclamation Or vbYesNo) = vbYes Then
            VerwijderItem Klant
         Else
            VerwijderItem Klant + 1
         End If
         Exit For
      End If
   Next Klant

EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure voegt een klant toe aan de lijst.
Public Sub VoegItemToe()
On Error GoTo Fout
Dim Keuze As Long

   AantalItemsV = AantalItemsV + 1
   ReDim Preserve DataV(0 To 6, 0 To AantalItemsV) As String

EindeRoutine:
   Exit Sub

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure wijzigt de gegevens van de opgegeven klant.
Public Sub WijzigItem(Klant As Long)
On Error GoTo Fout
Dim Keuze As Long

   DataV(KlNummer, Klant) = Buffer(KlNummer)
   DataV(KlNaam, Klant) = Buffer(KlNaam)
   DataV(KlAdres, Klant) = Buffer(KlAdres)
   DataV(KlPostcodePlaats, Klant) = Buffer(KlPostcodePlaats)
   DataV(KlTelefoonnummer, Klant) = Buffer(KlTelefoonnummer)
   DataV(KlFaxnummer, Klant) = Buffer(KlFaxnummer)
   DataV(KlEmailadres, Klant) = Buffer(KlEmailadres)

EindeRoutine:
   Exit Sub

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stelt dit object in.
Private Sub Class_Initialize()
On Error GoTo Fout
Dim Keuze As Long

   AantalItemsV = 0
   LeesFout = False
   ReDim DataV(0 To 6, 0 To AantalItemsV) As String

EindeRoutine:
   Exit Sub

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

