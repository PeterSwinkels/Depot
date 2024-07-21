VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FakturenObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Deze module bevat het fakturenobject.
Option Explicit

'Deze opsomming bevat de fakturendatavelden.
Public Enum FakturenDataVeldenE
   FkFaktuurnummer   'Definieert een faktuurnummer.
   FkFaktuurdatum    'Definieert een faktuurdatum.
End Enum

Private AantalItemsV As Long       'Bevat het aantal opgeslagen fakturen.
Private ActiefFaktuurV As String   'Bevat het nummer van een actief faktuur.
Private DataV() As String          'Bevat de gegevens van de fakturen.

'Deze procedure stuurt het aantal opgeslagen fakturen terug.
Public Property Get AantalItems() As Long
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   AantalItems = AantalItemsV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het aantal velden voor dit object terug.
Public Property Get AantalVelden() As Long
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   AantalVelden = 2
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het geselecteerde opgeslagen faktuur terug.
Public Property Get ActiefFaktuur() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   ActiefFaktuur = ActiefFaktuurV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de nieuwe faktuurselectie vast.
Public Property Let ActiefFaktuur(NieuwActiefFaktuur As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   ActiefFaktuurV = NieuwActiefFaktuur
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de gegevens van het opgegeven veld van het opgegeven opgeslagen faktuur terug.
Public Property Get Data(Veld As Long, Faktuurnr As Long) As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Data = DataV(Veld, Faktuurnr)
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven faktuurgegevens vast.
Public Property Let Data(Veld As Long, Faktuurnr As Long, NieuweData As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   DataV(Veld, Faktuurnr) = NieuweData
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure maakt een lijst van de opgeslagen fakturen.
Public Sub MaakLijst()
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim Bestand As String
Dim Keuze As Long

   Screen.MousePointer = vbHourglass
   ActieveMap = ".\Data\"
   AantalItemsV = 0
   ReDim DataV(0 To 1, 0 To AantalItemsV) As String
   Bestand = Dir$(".\Data\*.fkt", vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
   Do Until Bestand = vbNullString
      ActiefBestand = Bestand
      ReDim Preserve DataV(0 To 1, 0 To AantalItemsV) As String
      DataV(FkFaktuurnummer, AantalItemsV) = Left$(Bestand, InStr(UCase$(Bestand), ".FKT") - 1)
      DataV(FkFaktuurdatum, AantalItemsV) = FileDateTime(".\Data\" & Bestand)
      AantalItemsV = AantalItemsV + 1
      Bestand = Dir$()
   Loop
   Screen.MousePointer = vbDefault

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stuurt de totale breedte van alle velden terug.
Public Property Get TotaleVeldBreedte() As Long
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   TotaleVeldBreedte = 2
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de breedte van het opgegeven veld in verhouding tot de andere velden terug.
Public Property Get VeldBreedte(Veld As Long) As Long
On Error GoTo Fout
Dim Breedte As Long
Dim Keuze As Long

   Select Case Veld
      Case FkFaktuurnummer
         Breedte = 1
      Case FkFaktuurdatum
         Breedte = 1
      Case Else
         Breedte = 0
   End Select

EindeRoutine:
   VeldBreedte = Breedte
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de naam van het opgegeven veld terug.
Public Property Get VeldNaam(Veld As Long) As String
On Error GoTo Fout
Dim Keuze As Long
Dim Naam As String

   Select Case Veld
      Case FkFaktuurnummer
         Naam = "Faktuur Nummer"
      Case FkFaktuurdatum
         Naam = "Faktuur Datum"
      Case Else
         Naam = vbNullString
   End Select

EindeRoutine:
   VeldNaam = Naam
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt terug of het opgegeven veld rechtsuitgelijnd wordt.
Public Property Get VeldRechtsUitlijnen(Veld As Long) As Boolean
On Error GoTo Fout
Dim Keuze As Long
Dim RechtsUitlijnen As Boolean

   Select Case Veld
      Case FkFaktuurnummer
         RechtsUitlijnen = False
      Case FkFaktuurdatum
         RechtsUitlijnen = False
      Case Else
         RechtsUitlijnen = False
   End Select

EindeRoutine:
   VeldRechtsUitlijnen = RechtsUitlijnen
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stelt dit object in.
Private Sub Class_Initialize()
On Error GoTo Fout
Dim Keuze As Long

   AantalItemsV = 0
   ReDim DataV(0 To 1, 0 To AantalItemsV) As String

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

