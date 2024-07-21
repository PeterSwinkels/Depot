VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FaktuurObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Deze module beheert het actieve faktuur.
Option Explicit

'De opsomming bevat de faktuurvelden.
Public Enum FaktuurE
   FkArtikelNr      'Definieert het faktuurartikelnummer.
   FkAantal         'Definieert het aantal artikelen.
   FkAtikelNaam     'Definieert de artikelnaam.
   FkStukprijs      'Definieert de stukprijs.
   FkTotaalprijs    'Definieert de totaalprijs.
   FkUitVoorraad    'Definieert of een artikel uit de voorraad komt.
   FkLeverbaar      'Definieert of een artikel leverbaar is.
End Enum

Private AantalItemsV As Long              'Bevat het aantal artikelen in de lijst.
Private BTWBedragV As String              'Bevat de BTW berekend op basis van de totaal prijs.
Private BTWV As String                    'Bevat het BTW percentage.
Private DataV() As String                 'Bevat de velden met gegevens.
Private DatumopmaakV As String            'Bevat de faktuur datumopmaak.
Private DatumV As String                  'Bevat de faktuur datum.
Private DepotnaamV As String              'Bevat de naam van het depot.
Private FaktuurnrV As String              'Bevat het faktuurnummer.
Private FaktuurtekstV As String           'Bevat tekst die onderaan faktuur afgedrukt wordt.
Private IsParticulierV As Boolean         'Geeft aan of het faktuur voor een particulier is.
Private KlantAdresV As String             'Bevat het klant adres.
Private KlantEmailadresV As String        'Bevat het klant e-mailadres.
Private KlantFaxnummerV As String         'Bevat het klant faxnummer.
Private KlantGevondenV As Boolean         'Geeft aan of klant gevonden is in de klantenlijst.
Private KlantnaamV As String              'Bevat de klantnaam.
Private KlantnummerV As String            'Bevat het klantnummer.
Private KlantPostcodePlaatsV As String    'Bevat de klant postcode en woonplaats.
Private KlanttelefoonnummerV As String    'Beat het klanttelefoonnummer.
Private LaatsteDatumV As String           'Bevat de meest recente datum waarop een faktuur is opgeslagen.
Private LaatsteSubnrV As String           'Bevat het subnummer van de meest recent opgeslagen faktuur.
Private LeesFout As Boolean               'Geeft aan of er een fout is opgetreden tijdens het lezen van een faktuurbestand.
Private MaxSubnrsV As String              'Bevat het maximaal toegestane aantal faktuursubnummers voor een periode.
Private NropmaakV As String               'Bevat de symbolen voor de faktuur nummer opmaak.
Private PortoV As String                  'Bevat de te betalen porto.
Private SubnrperiodeV As String           'Bevat de duur van de periode waarna het faktuursubnummer wordt gereset.
Private SubnrtekstV As String             'Bevat de tekst die aan het faktuurnummer wordt toegevoegd.
Private SubnrV As String                  'Bevat het faktuursubnummer.
Private SubtotaalV As String              'Bevat het faktuursubtotaal.
Private TotaalV As String                 'Bevat het faktuurtotaal.

'Deze procedure stuurt het aantal artikelen terug.
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
   AantalVelden = 7
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure berekent het faktuurtotaal.
Private Sub BerekenTotaal()
On Error GoTo Fout
Dim Keuze As Long

   TotaalV = RondAf(Val(SubtotaalV))
   BTWBedragV = RondAf(BerekenProcent(Val(BTWV), Val(TotaalV)))
   TotaalV = RondAf(Val(TotaalV) + Val(BTWBedragV) + Val(PortoV))

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stuurt het BTW-percentage terug.
Public Property Get BTW() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   BTW = BTWV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure ontvang het nieuwe BTW-percentage.
Public Property Let BTW(NieuweBTW As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   BTWV = NieuweBTW
   BerekenTotaal
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het BTW-bedrag terug.
Public Property Get BTWBedrag() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   BTWBedrag = BTWBedragV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de gegevens van het opgegeven veld van het opgegeven artikel terug.
Public Property Get Data(Veld As Long, Artikel As Long) As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Data = DataV(Veld, Artikel)
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de gegevens van het opgegeven veld van het opgegeven artikel vast.
Public Property Let Data(Veld As Long, Artikel As Long, NieuweData As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   DataV(Veld, Artikel) = NieuweData
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de faktuurdatum terug.
Public Property Get Datum() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Datum = DatumV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven faktuurdatum vast.
Public Property Let Datum(NieuweDatum As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   DatumV = NieuweDatum
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure controleert of de opgegeven datum binnen de opgegeven periode valt.
Private Function DatumIsZelfdePeriode(Dag1 As Long, Maand1 As Long, Jaar1 As Long, Dag2 As Long, Maand2 As Long, Jaar2 As Long, Periode As String) As Boolean
On Error GoTo Fout
Dim IsZelfdePeriode As Boolean
Dim Keuze As Long

   IsZelfdePeriode = False

   Select Case Periode
      Case "D"
         If Jaar1 = Jaar2 Then
            If Maand1 = Maand2 Then
               If Dag1 = Dag2 Then IsZelfdePeriode = True
            End If
         End If
      Case "M"
         If Jaar1 = Jaar2 Then
            If Maand1 = Maand2 Then IsZelfdePeriode = True
         End If
      Case "J"
         If Jaar1 = Jaar2 Then IsZelfdePeriode = True
   End Select

EindeRoutine:
   DatumIsZelfdePeriode = IsZelfdePeriode
   Exit Function

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function

'Deze procedure stuurt de datumopmaak terug.
Public Property Get Datumopmaak() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Datumopmaak = DatumopmaakV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven datumopmaak vast.
Public Property Let Datumopmaak(NieuweDatumOpmaak As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   DatumopmaakV = NieuweDatumOpmaak
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de depotnaam terug.
Public Property Get Depotnaam() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Depotnaam = DepotnaamV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven depotnaam vast.
Public Property Let Depotnaam(NieuweDepotnaam As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   DepotnaamV = NieuweDepotnaam
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure drukt een scheidingsbalk af op de opgegeven locatie.
Private Sub DrukBalkAf(x As Single, y As Single)
On Error GoTo Fout
Dim Keuze As Long

   Printer.Line (x, y + 0.5)-(x + 80, y + 0.5)
   Printer.Line (x, y + 0.7)-(x + 80, y + 0.9), , BF
   Printer.Line (x, y + 1.05)-(x + 80, y + 1.05)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure drukt het faktuur af.
Public Sub DrukFaktuurAf()
On Error GoTo Fout
Dim Artikel As Long
Dim Keuze As Long
Dim Regel As Single

   StelPrinterIn

   Screen.MousePointer = vbHourglass
   Artikel = 0
   Do
      DrukFaktuurkopAf

      DrukAf 71, 42, "sub", Cursief:=False, Vet:=True
      DrukAf 74 - Printer.TextWidth(BTWV & "% BTW"), 43, BTWV & "% BTW", Cursief:=False, Vet:=True
      DrukAf 69.5, 44, "porto", Cursief:=False, Vet:=True
      DrukAf 69, 45, "totaal", Cursief:=False, Vet:=True
      DrukAf 84 - Printer.TextWidth(SubtotaalV), 42, SubtotaalV, Cursief:=False, Vet:=False
      DrukAf 84 - Printer.TextWidth(BTWBedragV), 43, BTWBedragV, Cursief:=False, Vet:=False
      DrukAf 84 - Printer.TextWidth(PortoV), 44, PortoV, Cursief:=False, Vet:=False
      DrukAf 84 - Printer.TextWidth(TotaalV), 45, TotaalV, Cursief:=False, Vet:=True
      DrukAf 76, 45, EURO_TEKEN, Cursief:=False, Vet:=False
      Printer.Line (75, 42)-(85, 43), , B
      Printer.Line (75, 43)-(85, 44), , B
      Printer.Line (75, 44)-(85, 45), , B
      Printer.Line (75, 45)-(85, 46), , B

      Regel = 21
      Do
         If Artikel < AantalItemsV Then
            DrukAf 7, Regel, DataV(FkArtikelNr, Artikel), False, False
            DrukAf 23 - Printer.TextWidth(DataV(FkAantal, Artikel)), Regel, DataV(FkAantal, Artikel), False, False
            DrukAf 25, Regel, DataV(FkAtikelNaam, Artikel), False, False
            DrukAf 74 - Printer.TextWidth(DataV(FkStukprijs, Artikel)), Regel, DataV(FkStukprijs, Artikel), False, False
            Select Case DataV(FkLeverbaar, Artikel)
               Case "J"
                  DrukAf 84 - Printer.TextWidth(DataV(FkTotaalprijs, Artikel)), Regel, DataV(FkTotaalprijs, Artikel), False, False
               Case "N"
                  DrukAf 85 - Printer.TextWidth("niet leverbaar"), Regel, "niet leverbaar", False, False
            End Select
            Artikel = Artikel + 1
         End If
         Printer.Line (6, Regel)-(18, Regel + 1), , B
         Printer.Line (18, Regel)-(24, Regel + 1), , B
         Printer.Line (24, Regel)-(65, Regel + 1), , B
         Printer.Line (65, Regel)-(75, Regel + 1), , B
         Printer.Line (75, Regel)-(85, Regel + 1), , B
         Regel = Regel + 1
      Loop Until Regel = 41

      DrukFaktuurvoetAf

      If Not Artikel = AantalItemsV Then
         Printer.NewPage
      End If
   Loop Until Artikel = AantalItemsV
   Printer.EndDoc

EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure drukt de faktuurkop af.
Private Sub DrukFaktuurkopAf()
On Error GoTo Fout
Dim Keuze As Long

   Printer.Font.Size = 14
   DrukAf 6, 8, "FAKTUUR", Cursief:=True, Vet:=True
   Printer.Font.Size = 10
   DrukAf 67, 9.5, "Faktuurno. " & FaktuurnrV, Cursief:=False, Vet:=True
   DrukAf 6, 10, DepotnaamV, Cursief:=True, Vet:=True
   DrukBalkAf 6, 11

   If IsParticulierV Then
      DrukAf 6, 12.5, "Naam:", Cursief:=False, Vet:=False
      DrukAf 6, 13.5, "Adres:", Cursief:=False, Vet:=False
      DrukAf 6, 14.5, "Postcode/plaats:", Cursief:=False, Vet:=False
      DrukAf 6, 15.5, "E-mail:", Cursief:=False, Vet:=False
      DrukAf 22, 12.5, KlantnaamV, Cursief:=False, Vet:=False
      DrukAf 22, 13.5, KlantAdresV, Cursief:=False, Vet:=False
      DrukAf 22, 14.5, KlantPostcodePlaatsV, Cursief:=False, Vet:=False
      DrukAf 22, 15.5, KlantEmailadresV, Cursief:=False, Vet:=False
   Else
      DrukAf 6, 12.5, "Klant:", Cursief:=False, Vet:=True
      DrukAf 6, 13.5, "Klantnummer:", Cursief:=False, Vet:=False
      DrukAf 6, 14.5, "Naam:", Cursief:=False, Vet:=False
      DrukAf 6, 15.5, "Adres:", Cursief:=False, Vet:=False
      DrukAf 6, 16.5, "Postcode/plaats:", Cursief:=False, Vet:=False
      DrukAf 6, 17.5, "E-mail:", Cursief:=False, Vet:=False
      DrukAf 22, 12.5, KlantnaamV, Cursief:=False, Vet:=False
      DrukAf 22, 13.5, KlantnummerV, Cursief:=False, Vet:=False
      DrukAf 22, 14.5, KlantnaamV, Cursief:=False, Vet:=False
      DrukAf 22, 15.5, KlantAdresV, Cursief:=False, Vet:=False
      DrukAf 22, 16.5, KlantPostcodePlaatsV, Cursief:=False, Vet:=False
      DrukAf 22, 17.5, KlantEmailadresV, Cursief:=False, Vet:=False
   End If

   DrukAf 67.5, 13, "Datum: " & DatumV, Cursief:=False, Vet:=True
   DrukAf 71, 16.5, "Tel: " & KlantTelefoonnummer, Cursief:=False, Vet:=False
   DrukAf 70.5, 17.5, "Fax: " & KlantFaxNummer, Cursief:=False, Vet:=False

   DrukAf 8, 19.5, "artikelnr.:", Cursief:=False, Vet:=True
   DrukAf 18.5, 19.5, "aantal:", Cursief:=False, Vet:=True
   DrukAf 42, 19.5, "artikel:", Cursief:=False, Vet:=True
   DrukAf 66, 19.5, "stukprijs:", Cursief:=False, Vet:=True
   DrukAf 77, 19.5, "totaal:", Cursief:=False, Vet:=True

EindeRoutine:
   Exit Sub

Fout:
   LeesFout = True
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure drukt de faktuurvoet af.
Private Sub DrukFaktuurvoetAf()
On Error GoTo Fout
Dim Keuze As Long

   DrukBalkAf 6, 46
   DrukAf 6, 48, FaktuurtekstV, False, False

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stuurt het faktuurnummer terug.
Public Property Get Faktuurnr() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Faktuurnr = FaktuurnrV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt het opgegeven faktuurnummer vast.
Public Property Let Faktuurnr(NieuwFaktuurnr As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   FaktuurnrV = NieuwFaktuurnr
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de faktuurtekst terug.
Public Property Get Faktuurtekst() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Faktuurtekst = FaktuurtekstV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven faktuurtekst vast.
Public Property Let Faktuurtekst(NieuweFaktuurTekst As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   FaktuurtekstV = NieuweFaktuurTekst
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt terug of het factuur voor een particulier is bestemd terug.
Public Property Get IsParticulier() As Boolean
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   IsParticulier = IsParticulierV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt vast of de factuur voor een particulier is bestemd.
Public Property Let IsParticulier(NieuweIsParticulier As Boolean)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   IsParticulierV = NieuweIsParticulier
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het klant adres terug.
Public Property Get KlantAdres() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   KlantAdres = KlantAdresV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure ontvangt het opgegeven klant adres vast.
Public Property Let KlantAdres(NieuwKlantAdres As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   KlantAdresV = NieuwKlantAdres
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het klant e-mailadres terug.
Public Property Get KlantEmailadres() As String
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:

   KlantEmailadres = KlantEmailadresV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt het opgegeven klant e-mailadres vast.
Public Property Let KlantEmailadres(NieuwKlantEmailadres As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   KlantEmailadresV = NieuwKlantEmailadres
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het klant faxnummer terug.
Public Property Get KlantFaxNummer() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   KlantFaxNummer = KlantFaxnummerV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt het opgegeven klant faxnummer vast.
Public Property Let KlantFaxNummer(NieuwKlantFaxNummer As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   KlantFaxnummerV = NieuwKlantFaxNummer
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt of the klant is gevonden terug.
Public Property Get KlantGevonden() As Boolean
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   KlantGevonden = KlantGevondenV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de klantnaam terug.
Public Property Get Klantnaam() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Klantnaam = KlantnaamV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven klantnaam vast.
Public Property Let Klantnaam(NieuweKlantNaam As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   KlantnaamV = NieuweKlantNaam
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het klantnummer terug.
Public Property Get Klantnummer() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Klantnummer = KlantnummerV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt het opgegeven klantnummer vast.
Public Property Let Klantnummer(NieuwKlantnummer As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   KlantnummerV = NieuwKlantnummer
   If InStr(Nropmaak, "K") > 0 Then MaakFaktuurnr
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de klant postcode en plaats terug.
Public Property Get KlantPostcodePlaats() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   KlantPostcodePlaats = KlantPostcodePlaatsV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven postcode en plaats van de klant vast.
Public Property Let KlantPostcodePlaats(NieuweKlantPostcodePlaats As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   KlantPostcodePlaatsV = NieuweKlantPostcodePlaats
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het klant telefoonnummer terug.
Public Property Get KlantTelefoonnummer() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   KlantTelefoonnummer = KlanttelefoonnummerV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt het opgegeven klant telefoonnummer vast.
Public Property Let KlantTelefoonnummer(NieuwKlantTelefoonNummer As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   KlanttelefoonnummerV = NieuwKlantTelefoonNummer
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de laatste faktuur datum terug.
Public Property Get LaatsteDatum() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   LaatsteDatum = LaatsteDatumV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven faktuurdatum vast.
Public Property Let LaatsteDatum(NieuweLaatsteDatum As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   LaatsteDatumV = NieuweLaatsteDatum
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het laatste faktuursubnummer terug.
Public Property Get LaatsteSubnr() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   LaatsteSubnr = LaatsteSubnrV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt het opgegeven faktuursubnummer vast.
Public Property Let LaatsteSubnr(NieuweLaatsteSubnr As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   LaatsteSubnrV = NieuweLaatsteSubnr
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure maakt een faktuurnummer aan op basis van de faktuurnummeropmaak.
Public Sub MaakFaktuurnr()
On Error GoTo Fout
Dim Keuze As Long
Dim SymboolIndex As Long

   FaktuurnrV = vbNullString
   For SymboolIndex = 1 To Len(NropmaakV)
      FaktuurnrV = FaktuurnrV & SymboolWaarde(NropmaakV, SymboolIndex)
   Next SymboolIndex

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure maakt een nieuw faktuur aan.
Public Sub MaakNieuwFaktuur()
On Error GoTo Fout
Dim Keuze As Long

   AantalItemsV = 0
   BTWBedragV = RondAf("0")
   BTWV = "0"
   ReDim DataV(0 To 7, 0 To AantalItemsV) As String
   DatumV = SymboolWaarde(DatumopmaakV, 1) & "-" & SymboolWaarde(DatumopmaakV, 2) & "-" & SymboolWaarde(DatumopmaakV, 3)
   Fakturen.ActiefFaktuur = vbNullString
   If InStr(NropmaakV, "#") > 0 Then SubnrV = MaakNieuwSubnr()
   KlantnummerV = vbNullString
   KlantnaamV = vbNullString
   KlantAdresV = vbNullString
   KlantPostcodePlaatsV = vbNullString
   KlanttelefoonnummerV = vbNullString
   KlantFaxnummerV = vbNullString
   KlantEmailadresV = vbNullString
   IsParticulierV = False
   LeesFout = False
   PortoV = RondAf("0")
   SubtotaalV = RondAf("0")
   TotaalV = RondAf("0")

   MaakFaktuurnr

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stuurt een nieuw faktuursubnummer terug.
Public Function MaakNieuwSubnr() As String
On Error GoTo Fout
Dim Keuze As Long
Dim NieuwSubnr As String

   If LaatsteDatumV = vbNullString Then
      LaatsteDatumV = Format$(Day(Date), "00") & Format$(Month(Date), "00") & Format$(Year(Date), "0000")
   End If

   NieuwSubnr = vbNullString
   If DatumIsZelfdePeriode(Val(Left$(LaatsteDatumV, 2)), Val(Mid$(LaatsteDatumV, 3, 2)), Val(Mid$(LaatsteDatumV, 5)), Day(Date), Month(Date), Year(Date), SubnrperiodeV) Then
      If Val(LaatsteSubnrV) >= Val(MaxSubnrsV) Then
         SubnrV = vbNullString
         MsgBox "Geen nieuwe faktuurnummers beschikbaar voor deze periode.", vbExclamation
      Else
         NieuwSubnr = Trim$(Str$(Val(LaatsteSubnrV) + 1))
      End If
   Else
      NieuwSubnr = "0"
   End If

EindeRoutine:
   MaakNieuwSubnr = NieuwSubnr
   Exit Function

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function

'Deze procedure stuurt het maximum toegestane aantal faktuursubnummers terug.
Public Property Get MaxSubnrs() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   MaxSubnrs = MaxSubnrsV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt het maximum toegestane aantal faktuursubnummers vast.
Public Property Let MaxSubnrs(NieuwMaxSubnrs As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   MaxSubnrsV = NieuwMaxSubnrs
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de faktuurnummeropmaak terug.
Public Property Get Nropmaak() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Nropmaak = NropmaakV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven faktuurnummeropmaak vast.
Public Property Let Nropmaak(NieuweNropmaak As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   NropmaakV = NieuweNropmaak
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure opent het faktuur met het opgegeven nummer.
Public Sub OpenFaktuur(Faktuurnr As String)
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim Artikel As Long
Dim BestandH As Integer
Dim Keuze As Long
Dim Lengte As Long

   MaakNieuwFaktuur

   Screen.MousePointer = vbHourglass
   ActieveMap = ".\Data\"
   ActiefBestand = Faktuurnr & ".fkt"
   BestandH = FreeFile()
   LeesFout = False
   Open ".\Data\" & Faktuurnr & ".fkt" For Binary Lock Read Write As BestandH
      Lengte = Asc(Input$(1, BestandH)): BTWV = Input$(Lengte, BestandH)
      Lengte = Asc(Input$(1, BestandH)): DatumV = Input$(Lengte, BestandH)
      Lengte = Asc(Input$(1, BestandH)): KlantnummerV = Input$(Lengte, BestandH)
      Lengte = Asc(Input$(1, BestandH)): KlantnaamV = Input$(Lengte, BestandH)
      Lengte = Asc(Input$(1, BestandH)): KlantAdresV = Input$(Lengte, BestandH)
      Lengte = Asc(Input$(1, BestandH)): KlantPostcodePlaatsV = Input$(Lengte, BestandH)
      Lengte = Asc(Input$(1, BestandH)): KlanttelefoonnummerV = Input$(Lengte, BestandH)
      Lengte = Asc(Input$(1, BestandH)): KlantFaxnummerV = Input$(Lengte, BestandH)
      Lengte = Asc(Input$(1, BestandH)): KlantEmailadresV = Input$(Lengte, BestandH)
      IsParticulierV = -Asc(Input$(1, BestandH))
      Lengte = Asc(Input$(1, BestandH)): PortoV = Input$(Lengte, BestandH)
      Artikel = 0
      Do Until Loc(BestandH) = LOF(BestandH)
         VoegItemToe
         Lengte = Asc(Input$(1, BestandH)): DataV(FkArtikelNr, Artikel) = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): DataV(FkAantal, Artikel) = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): DataV(FkAtikelNaam, Artikel) = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): DataV(FkStukprijs, Artikel) = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): DataV(FkTotaalprijs, Artikel) = Input$(Lengte, BestandH)
         DataV(FkUitVoorraad, Artikel) = Trim$(Input$(1, BestandH))
         DataV(FkLeverbaar, Artikel) = Trim$(Input$(1, BestandH))
         Artikel = Artikel + 1
      Loop
   Close BestandH

   FaktuurnrV = Faktuurnr
   Fakturen.ActiefFaktuur = FaktuurnrV

EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   LeesFout = True
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure opent de faktuurtekst.
Public Sub OpenFaktuurtekst()
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim BestandH As Integer
Dim Keuze As Long

   ActieveMap = ".\Data\"
   ActiefBestand = "Faktuur.txt"
   BestandH = FreeFile()
   Open ".\Data\Faktuur.txt" For Binary Lock Read Write As BestandH
      FaktuurtekstV = Input$(LOF(BestandH), BestandH)
   Close BestandH

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stuurt het portobedrag terug.
Public Property Get Porto() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Porto = PortoV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt het opgegeven portobedrag vast.
Public Property Let Porto(NieuwePorto As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   PortoV = NieuwePorto
   BerekenTotaal
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure slaat het faktuur op.
Public Sub SlaFaktuurOp()
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim Afbreken As Boolean
Dim Artikel As Long
Dim BestandH As Integer
Dim Keuze As Long

   Afbreken = False
   If Not FaktuurnrV = vbNullString Then
      ActieveMap = ".\Data\"
      ActiefBestand = FaktuurnrV & ".fkt"
      BestandH = FreeFile()
      Open ".\Data\" & FaktuurnrV & ".fkt" For Binary Lock Read Write As BestandH
         If LOF(BestandH) > 0 Then
            If Not FaktuurnrV = Fakturen.ActiefFaktuur() Then
               Afbreken = True
               MsgBox "Faktuur nummer al in gebruik.", vbExclamation
            End If
         End If
      Close BestandH

      If LeesFout Then
         Afbreken = BevestigAfbrekenNaLeesFout(ActiefBestand)
      End If

      If Not Afbreken Then
         Screen.MousePointer = vbHourglass
         BestandH = FreeFile()
         Open ".\Data\" & FaktuurnrV & ".fkt" For Output Lock Read Write As 1
            Print #BestandH, Chr$(Len(BTWV)); BTWV;
            Print #BestandH, Chr$(Len(DatumV)); DatumV;
            Print #BestandH, Chr$(Len(KlantnummerV)); KlantnummerV;
            Print #BestandH, Chr$(Len(KlantnaamV)); KlantnaamV;
            Print #BestandH, Chr$(Len(KlantAdresV)); KlantAdresV;
            Print #BestandH, Chr$(Len(KlantPostcodePlaatsV)); KlantPostcodePlaatsV;
            Print #BestandH, Chr$(Len(KlanttelefoonnummerV)); KlanttelefoonnummerV;
            Print #BestandH, Chr$(Len(KlantFaxnummerV)); KlantFaxnummerV;
            Print #BestandH, Chr$(Len(KlantEmailadresV)); KlantEmailadresV;
            Print #BestandH, Chr$(-IsParticulierV);
            Print #BestandH, Chr$(Len(PortoV)); PortoV;

            For Artikel = 0 To AantalItemsV - 1
               Print #BestandH, Chr$(Len(DataV(FkArtikelNr, Artikel))); DataV(FkArtikelNr, Artikel);
               Print #BestandH, Chr$(Len(DataV(FkAantal, Artikel))); DataV(FkAantal, Artikel);
               Print #BestandH, Chr$(Len(DataV(FkAtikelNaam, Artikel))); DataV(FkAtikelNaam, Artikel);
               Print #BestandH, Chr$(Len(DataV(FkStukprijs, Artikel))); DataV(FkStukprijs, Artikel);
               Print #BestandH, Chr$(Len(DataV(FkTotaalprijs, Artikel))); DataV(FkTotaalprijs, Artikel);
               Print #BestandH, DataV(FkUitVoorraad, Artikel);
               Print #BestandH, DataV(FkLeverbaar, Artikel);
            Next Artikel
         Close BestandH

         LaatsteDatumV = Format$(Day(Date), "00") & Format$(Month(Date), "00") & Format$(Year(Date), "0000")
         LaatsteSubnrV = SubnrV
         Fakturen.ActiefFaktuur = FaktuurnrV
         SlaInstellingenOp
      End If
   End If

EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure deze procedure slaat de faktuurtekst op.
Public Sub SlaFaktuurtekstOp()
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim BestandH As Integer
Dim Keuze As Long

   ActieveMap = ".\Data\"
   ActiefBestand = "Faktuur.txt"
   BestandH = FreeFile()
   Open ".\Data\Faktuur.txt" For Output Lock Read Write As BestandH
      Print #BestandH, FaktuurtekstV;
   Close BestandH

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure sorteert de artikelen.
Public Sub SorteerArtikelen()
On Error GoTo Fout
Dim AnderArtikel As Long
Dim Artikel As Long
Dim Keuze As Long
Dim Veld As Long

   Screen.MousePointer = vbHourglass

   For Artikel = 0 To AantalItemsV - 1
      For AnderArtikel = 0 To AantalItemsV - 1
         If Not AnderArtikel = Artikel Then
            If DataV(FkArtikelNr, AnderArtikel) > DataV(FkArtikelNr, Artikel) Then
               For Veld = LBound(DataV(), 1) To UBound(DataV(), 1)
                  Verwissel DataV(Veld, AnderArtikel), DataV(Veld, Artikel)
               Next Veld
            End If
         End If
      Next AnderArtikel
   Next Artikel

   For Artikel = 0 To AantalItemsV - 2
      If DataV(FkArtikelNr, Artikel) = DataV(FkArtikelNr, Artikel + 1) Then
         If MsgBox("Dit artikel is al ingevoerd. Vervangen?", vbExclamation Or vbYesNo) = vbYes Then
            VerwijderItem Artikel
         Else
            VerwijderItem Artikel + 1
         End If
         Exit For
      End If
   Next Artikel

   SubtotaalV = RondAf(0)
   For Artikel = 0 To AantalItemsV - 1
      SubtotaalV = RondAf(Str$(Val(SubtotaalV) + Val(DataV(FkTotaalprijs, Artikel))))
   Next Artikel

   BerekenTotaal

EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedures stelt de standaardinvoerwaardes in.
Public Sub StelStandaardWaardesIn()
On Error GoTo Fout
Dim Keuze As Long

   Buffer(FkArtikelNr) = vbNullString
   Buffer(FkAantal) = "0"
   Buffer(FkAtikelNaam) = vbNullString
   Buffer(FkStukprijs) = RondAf("0")
   Buffer(FkTotaalprijs) = RondAf("0")
   Buffer(FkUitVoorraad) = "J"
   Buffer(FkLeverbaar) = "J"

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stuurt het faktuursubnummer terug.
Public Property Get Subnr() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Subnr = SubnrV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt het opgegeven faktuursubnummer vast.
Public Property Let Subnr(NieuwSubnr As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   SubnrV = NieuwSubnr
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de faktuursubnummerperiode terug.
Public Property Get Subnrperiode() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Subnrperiode = SubnrperiodeV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de faktuur subnummerperiode vast.
Public Property Let Subnrperiode(NieuweSubnrperiode As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   SubnrperiodeV = NieuweSubnrperiode
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de faktuursubnummertekst terug.
Public Property Get Subnrtekst() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Subnrtekst = SubnrtekstV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven faktuursubnummertekst vast.
Public Property Let Subnrtekst(NieuweNrTekst As String)
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   SubnrtekstV = NieuweNrTekst
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het subtotaalbedrag terug.
Public Property Get Subtotaal() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Subtotaal = SubtotaalV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het totaalbedrag terug.
Public Property Get Totaal() As String
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Totaal = TotaalV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de totale breedte van alle velden terug.
Public Property Get TotaleVeldBreedte() As Long
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   TotaleVeldBreedte = 8
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
      Case FkArtikelNr
         Breedte = 1
      Case FkAantal
         Breedte = 1
      Case FkAtikelNaam
         Breedte = 2
      Case FkStukprijs
         Breedte = 1
      Case FkTotaalprijs
         Breedte = 1
      Case FkUitVoorraad
         Breedte = 1
      Case FkLeverbaar
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
      Case FkArtikelNr
         Naam = "Artikelnr."
      Case FkAantal
         Naam = "Aantal"
      Case FkAtikelNaam
         Naam = "Artikel"
      Case FkStukprijs
         Naam = "Stukprijs"
      Case FkTotaalprijs
         Naam = "Totaalprijs"
      Case FkUitVoorraad
         Naam = "Uit Voorraad"
      Case FkLeverbaar
         Naam = "Leverbaar"
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
      Case FkArtikelNr
         RechtsUitlijnen = False
      Case FkAtikelNaam
         RechtsUitlijnen = False
      Case FkAantal
         RechtsUitlijnen = True
      Case FkStukprijs
         RechtsUitlijnen = True
      Case FkTotaalprijs
         RechtsUitlijnen = True
      Case FkUitVoorraad
         RechtsUitlijnen = False
      Case FkLeverbaar
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

'Deze procedure verwijdert het opgegeven artikel uit de lijst.
Public Sub VerwijderItem(Artikel As Long)
On Error GoTo Fout
Dim Keuze As Long

   If Not AantalItemsV = 0 Then
      VerplaatsItem DataV(), Artikel, AantalItemsV - 1
      AantalItemsV = AantalItemsV - 1
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure voegt een artikel toe aan de lijst.
Public Sub VoegItemToe()
On Error GoTo Fout
Dim Keuze As Long

   AantalItemsV = AantalItemsV + 1
   ReDim Preserve DataV(0 To 7, 0 To AantalItemsV) As String

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure werkt de artikellijst bij.
Public Sub WerkArtikelenBij()
On Error GoTo Fout
Dim Artikel As Long
Dim Gevonden As Boolean
Dim Keuze As Long

   Screen.MousePointer = vbHourglass
   ActieIsToeVoegen = False
   For Artikel = 0 To AantalItemsV - 1
      Buffer(FkArtikelNr) = DataV(FkArtikelNr, Artikel)
      Buffer(FkAantal) = DataV(FkAantal, Artikel)
      Buffer(FkUitVoorraad) = DataV(FkUitVoorraad, Artikel)
      Buffer(FkLeverbaar) = DataV(FkLeverbaar, Artikel)
      If DataV(FkUitVoorraad, Artikel) = "N" Then
         Buffer(FkAtikelNaam) = DataV(FkAtikelNaam, Artikel)
         Buffer(FkStukprijs) = DataV(FkStukprijs, Artikel)
         Buffer(FkTotaalprijs) = DataV(FkTotaalprijs, Artikel)
      ElseIf IsParticulierV Then
         Buffer(FkStukprijs) = DataV(FkStukprijs, Artikel)
      End If
      WijzigItem Artikel, Gevonden
   Next Artikel

EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure werkt de klantgegevens bij.
Public Sub WerkKlantgegevensBij()
On Error GoTo Fout
Dim Keuze As Long
Dim KlantIndex As Long

   If IsParticulierV Then
      KlantGevondenV = True
   ElseIf Not IsParticulierV Then
      If KlantnummerV = vbNullString Then
         KlantGevondenV = True
         KlantnaamV = vbNullString
         KlantAdresV = vbNullString
         KlantPostcodePlaats = vbNullString
         KlanttelefoonnummerV = vbNullString
         KlantFaxnummerV = vbNullString
         KlantEmailadresV = vbNullString
      ElseIf Not KlantnummerV = vbNullString Then
         KlantIndex = ZoekKlantIndex()
         If KlantGevondenV Then
            KlantnaamV = Klanten.Data(KlNaam, KlantIndex)
            KlantAdresV = Klanten.Data(KlAdres, KlantIndex)
            KlantPostcodePlaatsV = Klanten.Data(KlPostcodePlaats, KlantIndex)
            KlanttelefoonnummerV = Klanten.Data(KlTelefoonNummer, KlantIndex)
            KlantFaxnummerV = Klanten.Data(KlFaxNummer, KlantIndex)
            KlantEmailadresV = Klanten.Data(KlEmailadres, KlantIndex)
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

'Deze procedure werkt de voorraad bij.
Private Sub WerkVoorraadBij(Artikel As Long, VoorraadArtikelIndex As Long)
On Error GoTo Fout
Dim Keuze As Long

   Voorraad.Data(VrAantal, VoorraadArtikelIndex) = Val(Voorraad.Data(VrAantal, VoorraadArtikelIndex)) + Val(DataV(FkAantal, Artikel))
   Voorraad.Data(VrAftotaal, VoorraadArtikelIndex) = Val(Voorraad.Data(VrAftotaal, VoorraadArtikelIndex)) - Val(DataV(FkAantal, Artikel))

   If Val(Voorraad.Data(VrAantal, VoorraadArtikelIndex)) - Val(Buffer(FkAantal)) < 0 Then
      MsgBox "Onvoldoende artikelen in voorraad.", vbExclamation
      DataV(FkAantal, Artikel) = "0"
   Else
      DataV(FkAantal, Artikel) = Buffer(FkAantal)
      Voorraad.Data(VrAantal, VoorraadArtikelIndex) = Val(Voorraad.Data(VrAantal, VoorraadArtikelIndex)) - Val(DataV(FkAantal, Artikel))
   End If
   Voorraad.Data(VrAftotaal, VoorraadArtikelIndex) = Val(Voorraad.Data(VrAftotaal, VoorraadArtikelIndex)) + Val(DataV(FkAantal, Artikel))

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure wijzigt het opgegeven artikel.
Public Sub WijzigItem(Artikel As Long, Gevonden As Boolean)
On Error GoTo Fout
Dim Keuze As Long
Dim VoorraadArtikelIndex As Long

   DataV(FkUitVoorraad, Artikel) = Buffer(FkUitVoorraad)

   If DataV(FkUitVoorraad, Artikel) = "J" Then
      VoorraadArtikelIndex = ZoekVoorraadArtikelIndex(Buffer(FkArtikelNr), Gevonden)

      If Gevonden Then
         DataV(FkUitVoorraad, Artikel) = "J"

         WerkVoorraadBij Artikel, VoorraadArtikelIndex
         DataV(FkAtikelNaam, Artikel) = Voorraad.Data(VrNaam, VoorraadArtikelIndex)
         If IsParticulierV Then
            DataV(FkStukprijs, Artikel) = RondAf(Buffer(FkStukprijs))
         Else
            DataV(FkStukprijs, Artikel) = Voorraad.Data(VrStukprijs, VoorraadArtikelIndex)
         End If
      Else
         If ActieIsToeVoegen Then
            VerwijderItem Artikel
         Else
            DataV(FkUitVoorraad, Artikel) = "X"
         End If
      End If
   ElseIf DataV(FkUitVoorraad, Artikel) = "N" Then
      Gevonden = True
      DataV(FkAantal, Artikel) = Buffer(FkAantal)
      DataV(FkAtikelNaam, Artikel) = Buffer(FkAtikelNaam)
      DataV(FkStukprijs, Artikel) = RondAf(Buffer(FkStukprijs))
   End If

   If Gevonden Then
      DataV(FkArtikelNr, Artikel) = Buffer(FkArtikelNr)
      DataV(FkTotaalprijs, Artikel) = RondAf(Str$(Val(DataV(FkStukprijs, Artikel)) * Val(DataV(FkAantal, Artikel))))
      DataV(FkLeverbaar, Artikel) = Buffer(FkLeverbaar)
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stuurt de klant met een specifiek nummer terug.
Private Function ZoekKlantIndex() As Long
On Error GoTo Fout
Dim Keuze As Long
Dim Klant As Long
Dim KlantIndex As Long

   KlantGevondenV = False

   If KlantnummerV = vbNullString Or IsParticulierV Then
      KlantGevondenV = True
   Else
      KlantIndex = 0
      For Klant = 0 To Klanten.AantalItems() - 1
         If Klanten.Data(KlNummer, Klant) = KlantnummerV Then
            KlantIndex = Klant
            KlantGevondenV = True
            Exit For
         End If
      Next Klant
   End If

EindeRoutine:
   ZoekKlantIndex = KlantIndex
   Exit Function

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function

'Deze procedure stuurt het artikel met het opgegeven nummer terug.
Private Function ZoekVoorraadArtikelIndex(ArtikelNr As String, Gevonden As Boolean) As Long
On Error GoTo Fout
Dim Artikel As Long
Dim ArtikelIndex As Long
Dim Keuze As Long

   Gevonden = False
   ArtikelIndex = 0
   For Artikel = 0 To Voorraad.AantalItems() - 1
      If Voorraad.Data(VrArtikelnr, Artikel) = ArtikelNr Then
         ArtikelIndex = Artikel
         Gevonden = True
         Exit For
      End If
   Next Artikel

EindeRoutine:
   ZoekVoorraadArtikelIndex = ArtikelIndex
   Exit Function

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function

'Deze procedure stelt dit object in.
Private Sub Class_Initialize()
On Error GoTo Fout
Dim Keuze As Long

   MaakNieuwFaktuur

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

