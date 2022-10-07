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

'De opsomming bevat de faktuur velden.
Public Enum FaktuurE
   FkArtikelNr      'Definieert het faktuur artikel nummer.
   FkAantal         'Definieert het aantal artikelen.
   FkAtikelNaam     'Definieert de artikel naam.
   FkStukprijs      'Definieert de stukprijs.
   FkTotaalPrijs    'Definieert de totaalprijs.
   FkUitVoorraad    'Definieert of een artikel uit de voorraad komt.
   FkLeverbaar      'Definieert of een artikel leverbaar is.
End Enum

Private AantalItemsV As Long              'Bevat het aantal artikelen in de lijst.
Private BTWBedragV As String              'Bevat de BTW berekend op basis van de totaal prijs.
Private BTWV As String                    'Bevat het BTW percentage.
Private DataV() As String                 'Bevat de velden met gegevens.
Private DatumOpmaakV As String            'Bevat de faktuur datum opmaak.
Private DatumV As String                  'Bevat de faktuur datum.
Private DepotNaamV As String              'Bevat de naam van het depot.
Private FaktuurNrV As String              'Bevat het faktuur nummer.
Private FaktuurTekstV As String           'Bevat tekst die onder aan faktuur afgedrukt wordt.
Private IsParticulierV As Boolean         'Geeft aan of het faktuur voor een particulier is.
Private KlantAdresV As String             'Bevat het klant adres.
Private KlantEmailAdresV As String        'Bevat het klant e-mail adres.
Private KlantFaxNummerV As String         'Bevat het klant faxnummer.
Private KlantGevondenV As Boolean         'Geeft aan of klant gevonden is in de klanten lijst.
Private KlantNaamV As String              'Bevat de klant naam.
Private KlantNummerV As String            'Bevat het klant nummer.
Private KlantPostcodePlaatsV As String    'Bevat de klant postcode en woonplaats.
Private KlantTelefoonNummerV As String    'Beat het klant telefoonnummer.
Private LaatsteDatumV As String           'Bevat de meest recente datum waarop een faktuur is opgeslagen.
Private LaatsteSubNrV As String           'Bevat het subnummer van de meest recent opgeslagen faktuur.
Private LeesFout As Boolean               'Geeft aan of er een fout is opgetreden tijdens het lezen van een faktuur bestand.
Private MaxSubNrsV As String              'Bevat het maximaal toegestane aantal faktuur subnummers voor een periode.
Private NrOpmaakV As String               'Bevat de symbolen voor de faktuur nummer opmaak.
Private PortoV As String                  'Bevat de te betalen porto.
Private SubNrPeriodeV As String           'Bevat de duur van de periode waarna het faktuur subnummer wordt gereset.
Private SubNrTekstV As String             'Bevat de tekst die aan het faktuur nummer wordt toegevoegd.
Private SubNrV As String                  'Bevat het faktuur subnummer.
Private SubtotaalV As String              'Bevat het faktuur subtotaal.
Private TotaalV As String                 'Bevat het faktuur totaal.

'Deze procedure berekent het faktuur totaal.
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

'Deze procedure controleert of de opgegeven datum binnen de opgegeven dag, maand en/of jaar valt.
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
      DrukFaktuurKopAf
      
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
                  DrukAf 84 - Printer.TextWidth(DataV(FkTotaalPrijs, Artikel)), Regel, DataV(FkTotaalPrijs, Artikel), False, False
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
     
      DrukFaktuurVoetAf
      
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

'Deze procedure drukt de faktuur kop af.
Private Sub DrukFaktuurKopAf()
On Error GoTo Fout
Dim Keuze As Long

   Printer.Font.Size = 14
   DrukAf 6, 8, "FAKTUUR", Cursief:=True, Vet:=True
   Printer.Font.Size = 10
   DrukAf 67, 9.5, "Faktuurno. " & FaktuurNrV, Cursief:=False, Vet:=True
   DrukAf 6, 10, DepotNaamV, Cursief:=True, Vet:=True
   DrukBalkAf 6, 11

   If IsParticulierV Then
      DrukAf 6, 12.5, "Naam:", Cursief:=False, Vet:=False
      DrukAf 6, 13.5, "Adres:", Cursief:=False, Vet:=False
      DrukAf 6, 14.5, "Postcode/plaats:", Cursief:=False, Vet:=False
      DrukAf 6, 15.5, "E-mail:", Cursief:=False, Vet:=False
      DrukAf 22, 12.5, KlantNaamV, Cursief:=False, Vet:=False
      DrukAf 22, 13.5, KlantAdresV, Cursief:=False, Vet:=False
      DrukAf 22, 14.5, KlantPostcodePlaatsV, Cursief:=False, Vet:=False
      DrukAf 22, 15.5, KlantEmailAdresV, Cursief:=False, Vet:=False
   Else
      DrukAf 6, 12.5, "Klant:", Cursief:=False, Vet:=True
      DrukAf 6, 13.5, "Klantnummer:", Cursief:=False, Vet:=False
      DrukAf 6, 14.5, "Naam:", Cursief:=False, Vet:=False
      DrukAf 6, 15.5, "Adres:", Cursief:=False, Vet:=False
      DrukAf 6, 16.5, "Postcode/plaats:", Cursief:=False, Vet:=False
      DrukAf 6, 17.5, "E-mail:", Cursief:=False, Vet:=False
      DrukAf 22, 12.5, KlantNaamV, Cursief:=False, Vet:=False
      DrukAf 22, 13.5, KlantNummerV, Cursief:=False, Vet:=False
      DrukAf 22, 14.5, KlantNaamV, Cursief:=False, Vet:=False
      DrukAf 22, 15.5, KlantAdresV, Cursief:=False, Vet:=False
      DrukAf 22, 16.5, KlantPostcodePlaatsV, Cursief:=False, Vet:=False
      DrukAf 22, 17.5, KlantEmailAdresV, Cursief:=False, Vet:=False
   End If
   
   DrukAf 67.5, 13, "Datum: " & DatumV, Cursief:=False, Vet:=True
   DrukAf 71, 16.5, "Tel: " & KlantTelefoonNummer, Cursief:=False, Vet:=False
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

'Deze procedure drukt de faktuur voet af.
Private Sub DrukFaktuurVoetAf()
On Error GoTo Fout
Dim Keuze As Long

   DrukBalkAf 6, 46
   DrukAf 6, 48, FaktuurTekstV, False, False
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure maakt een faktuur nummer aan op basis van de faktuur nummer opmaak.
Public Sub MaakFaktuurNr()
On Error GoTo Fout
Dim Keuze As Long
Dim SymboolIndex As Long

   FaktuurNrV = vbNullString
   For SymboolIndex = 1 To Len(NrOpmaakV)
      FaktuurNrV = FaktuurNrV & SymboolWaarde(NrOpmaakV, SymboolIndex)
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
   DatumV = SymboolWaarde(DatumOpmaakV, 1) & "-" & SymboolWaarde(DatumOpmaakV, 2) & "-" & SymboolWaarde(DatumOpmaakV, 3)
   Fakturen.ActiefFaktuur = vbNullString
   If InStr(NrOpmaakV, "#") > 0 Then SubNrV = MaakNieuwSubNr()
   KlantNummerV = vbNullString
   KlantNaamV = vbNullString
   KlantAdresV = vbNullString
   KlantPostcodePlaatsV = vbNullString
   KlantTelefoonNummerV = vbNullString
   KlantFaxNummerV = vbNullString
   KlantEmailAdresV = vbNullString
   IsParticulierV = False
   LeesFout = False
   PortoV = RondAf("0")
   SubtotaalV = RondAf("0")
   TotaalV = RondAf("0")
   
   MaakFaktuurNr
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stuurt een nieuw faktuur subnummer terug.
Public Function MaakNieuwSubNr() As String
On Error GoTo Fout
Dim Keuze As Long
Dim NieuwSubNr As String

   If LaatsteDatumV = vbNullString Then
      LaatsteDatumV = Format$(Day(Date), "00") & Format$(Month(Date), "00") & Format$(Year(Date), "0000")
   End If
 
   NieuwSubNr = vbNullString
   If DatumIsZelfdePeriode(Val(Left$(LaatsteDatumV, 2)), Val(Mid$(LaatsteDatumV, 3, 2)), Val(Mid$(LaatsteDatumV, 5)), Day(Date), Month(Date), Year(Date), SubNrPeriodeV) Then
      If Val(LaatsteSubNrV) >= Val(MaxSubNrsV) Then
         SubNrV = vbNullString
         MsgBox "Geen nieuwe faktuur nummers beschikbaar voor deze periode.", vbExclamation
      Else
         NieuwSubNr = Trim$(Str$(Val(LaatsteSubNrV) + 1))
      End If
   Else
      NieuwSubNr = "0"
   End If
   
EindeRoutine:
   MaakNieuwSubNr = NieuwSubNr
   Exit Function

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function

'Deze procedure opent het faktuur met het opgegeven nummer.
Public Sub OpenFaktuur(FaktuurNr As String)
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
   ActiefBestand = FaktuurNr & ".fkt"
   BestandH = FreeFile()
   LeesFout = False
   Open ".\Data\" & FaktuurNr & ".fkt" For Binary Lock Read Write As BestandH
      Lengte = Asc(Input$(1, BestandH)): BTWV = Input$(Lengte, BestandH)
      Lengte = Asc(Input$(1, BestandH)): DatumV = Input$(Lengte, BestandH)
      Lengte = Asc(Input$(1, BestandH)): KlantNummerV = Input$(Lengte, BestandH)
      Lengte = Asc(Input$(1, BestandH)): KlantNaamV = Input$(Lengte, BestandH)
      Lengte = Asc(Input$(1, BestandH)): KlantAdresV = Input$(Lengte, BestandH)
      Lengte = Asc(Input$(1, BestandH)): KlantPostcodePlaatsV = Input$(Lengte, BestandH)
      Lengte = Asc(Input$(1, BestandH)): KlantTelefoonNummerV = Input$(Lengte, BestandH)
      Lengte = Asc(Input$(1, BestandH)): KlantFaxNummerV = Input$(Lengte, BestandH)
      Lengte = Asc(Input$(1, BestandH)): KlantEmailAdresV = Input$(Lengte, BestandH)
      IsParticulierV = -Asc(Input$(1, BestandH))
      Lengte = Asc(Input$(1, BestandH)): PortoV = Input$(Lengte, BestandH)
      Artikel = 0
      Do Until Loc(BestandH) = LOF(BestandH)
         VoegItemToe
         Lengte = Asc(Input$(1, BestandH)): DataV(FkArtikelNr, Artikel) = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): DataV(FkAantal, Artikel) = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): DataV(FkAtikelNaam, Artikel) = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): DataV(FkStukprijs, Artikel) = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): DataV(FkTotaalPrijs, Artikel) = Input$(Lengte, BestandH)
         DataV(FkUitVoorraad, Artikel) = Trim$(Input$(1, BestandH))
         DataV(FkLeverbaar, Artikel) = Trim$(Input$(1, BestandH))
         Artikel = Artikel + 1
      Loop
   Close BestandH
      
   FaktuurNrV = FaktuurNr
   Fakturen.ActiefFaktuur = FaktuurNrV
EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   LeesFout = True
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure opent de faktuur tekst.
Public Sub OpenFaktuurTekst()
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim BestandH As Integer
Dim Keuze As Long

   ActieveMap = ".\Data\"
   ActiefBestand = "Faktuur.txt"
   BestandH = FreeFile()
   Open ".\Data\Faktuur.txt" For Binary Lock Read Write As BestandH
      FaktuurTekstV = Input$(LOF(BestandH), BestandH)
   Close BestandH
EindeRoutine:
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

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
   If Not FaktuurNrV = vbNullString Then
      ActieveMap = ".\Data\"
      ActiefBestand = FaktuurNrV & ".fkt"
      BestandH = FreeFile()
      Open ".\Data\" & FaktuurNrV & ".fkt" For Binary Lock Read Write As BestandH
         If LOF(BestandH) > 0 Then
            If Not FaktuurNrV = Fakturen.ActiefFaktuur() Then
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
         Open ".\Data\" & FaktuurNrV & ".fkt" For Output Lock Read Write As 1
            Print #BestandH, Chr$(Len(BTWV)); BTWV;
            Print #BestandH, Chr$(Len(DatumV)); DatumV;
            Print #BestandH, Chr$(Len(KlantNummerV)); KlantNummerV;
            Print #BestandH, Chr$(Len(KlantNaamV)); KlantNaamV;
            Print #BestandH, Chr$(Len(KlantAdresV)); KlantAdresV;
            Print #BestandH, Chr$(Len(KlantPostcodePlaatsV)); KlantPostcodePlaatsV;
            Print #BestandH, Chr$(Len(KlantTelefoonNummerV)); KlantTelefoonNummerV;
            Print #BestandH, Chr$(Len(KlantFaxNummerV)); KlantFaxNummerV;
            Print #BestandH, Chr$(Len(KlantEmailAdresV)); KlantEmailAdresV;
            Print #BestandH, Chr$(-IsParticulierV);
            Print #BestandH, Chr$(Len(PortoV)); PortoV;
            
            For Artikel = 0 To AantalItemsV - 1
               Print #BestandH, Chr$(Len(DataV(FkArtikelNr, Artikel))); DataV(FkArtikelNr, Artikel);
               Print #BestandH, Chr$(Len(DataV(FkAantal, Artikel))); DataV(FkAantal, Artikel);
               Print #BestandH, Chr$(Len(DataV(FkAtikelNaam, Artikel))); DataV(FkAtikelNaam, Artikel);
               Print #BestandH, Chr$(Len(DataV(FkStukprijs, Artikel))); DataV(FkStukprijs, Artikel);
               Print #BestandH, Chr$(Len(DataV(FkTotaalPrijs, Artikel))); DataV(FkTotaalPrijs, Artikel);
               Print #BestandH, DataV(FkUitVoorraad, Artikel);
               Print #BestandH, DataV(FkLeverbaar, Artikel);
            Next Artikel
         Close BestandH
                  
         LaatsteDatumV = Format$(Day(Date), "00") & Format$(Month(Date), "00") & Format$(Year(Date), "0000")
         LaatsteSubNrV = SubNrV
         Fakturen.ActiefFaktuur = FaktuurNrV
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

'Deze procedure deze procedure slaat de faktuur tekst op.
Public Sub SlaFaktuurTekstOp()
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim BestandH As Integer
Dim Keuze As Long

   ActieveMap = ".\Data\"
   ActiefBestand = "Faktuur.txt"
   BestandH = FreeFile()
   Open ".\Data\Faktuur.txt" For Output Lock Read Write As BestandH
      Print #BestandH, FaktuurTekstV;
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
      SubtotaalV = RondAf(Str$(Val(SubtotaalV) + Val(DataV(FkTotaalPrijs, Artikel))))
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

'Deze procedures stelt de standaard invoer waardes in.
Public Sub StelStandaardWaardesIn()
On Error GoTo Fout
Dim Keuze As Long

   Buffer(FkArtikelNr) = vbNullString
   Buffer(FkAantal) = "0"
   Buffer(FkAtikelNaam) = vbNullString
   Buffer(FkStukprijs) = RondAf("0")
   Buffer(FkTotaalPrijs) = RondAf("0")
   Buffer(FkUitVoorraad) = "J"
   Buffer(FkLeverbaar) = "J"
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

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

'Deze procedure werkt de artikel lijst bij.
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
         Buffer(FkTotaalPrijs) = DataV(FkTotaalPrijs, Artikel)
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

'Deze procedure werkt de klant gegevens bij.
Public Sub WerkKlantGegevensBij()
On Error GoTo Fout
Dim Keuze As Long
Dim KlantIndex As Long
 
   If IsParticulierV Then
      KlantGevondenV = True
   ElseIf Not IsParticulierV Then
      If KlantNummerV = vbNullString Then
         KlantGevondenV = True
         KlantNaamV = vbNullString
         KlantAdresV = vbNullString
         KlantPostcodePlaats = vbNullString
         KlantTelefoonNummerV = vbNullString
         KlantFaxNummerV = vbNullString
         KlantEmailAdresV = vbNullString
      ElseIf Not KlantNummerV = vbNullString Then
         KlantIndex = ZoekKlantIndex()
         If KlantGevondenV Then
            KlantNaamV = Klanten.Data(KlNaam, KlantIndex)
            KlantAdresV = Klanten.Data(KlAdres, KlantIndex)
            KlantPostcodePlaatsV = Klanten.Data(KlPostcodePlaats, KlantIndex)
            KlantTelefoonNummerV = Klanten.Data(KlTelefoonNummer, KlantIndex)
            KlantFaxNummerV = Klanten.Data(KlFaxNummer, KlantIndex)
            KlantEmailAdresV = Klanten.Data(KlEmailAdres, KlantIndex)
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
      DataV(FkTotaalPrijs, Artikel) = RondAf(Str$(Val(DataV(FkStukprijs, Artikel)) * Val(DataV(FkAantal, Artikel))))
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
   
   If KlantNummerV = vbNullString Or IsParticulierV Then
      KlantGevondenV = True
   Else
      KlantIndex = 0
      For Klant = 0 To Klanten.AantalItems() - 1
         If Klanten.Data(KlNummer, Klant) = KlantNummerV Then
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

'Deze procedure stuurt het BTW percentage terug.
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

'Deze procedure ontvang het nieuwe BTW percentage.
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

'Deze procedure stuurt het BTW bedrag terug.
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

'Deze procedure stuurt de faktuur datum terug.
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

'Deze procedure legt de opgegeven faktuur datum vast.
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

'Deze procedure stuurt de datum opmaak terug.
Public Property Get DatumOpmaak() As String
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   DatumOpmaak = DatumOpmaakV
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven datum opmaak vast.
Public Property Let DatumOpmaak(NieuweDatumOpmaak As String)
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   DatumOpmaakV = NieuweDatumOpmaak
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de depot naam terug.
Public Property Get DepotNaam() As String
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   DepotNaam = DepotNaamV
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven depot naam vast.
Public Property Let DepotNaam(NieuweDepotNaam As String)
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   DepotNaamV = NieuweDepotNaam
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het faktuur nummer terug.
Public Property Get FaktuurNr() As String
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   FaktuurNr = FaktuurNrV
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt het opgegeven faktuur nummer vast.
Public Property Let FaktuurNr(NieuwFaktuurNr As String)
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   FaktuurNrV = NieuwFaktuurNr
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de faktuur tekst terug.
Public Property Get FaktuurTekst() As String
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   FaktuurTekst = FaktuurTekstV
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven faktuur tekst vast.
Public Property Let FaktuurTekst(NieuweFaktuurTekst As String)
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   FaktuurTekstV = NieuweFaktuurTekst
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

'Deze procedure stuurt het klant e-mail adres terug.
Public Property Get KlantEmailAdres() As String
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   KlantEmailAdres = KlantEmailAdresV
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt het opgegeven klant e-mail adres vast.
Public Property Let KlantEmailAdres(NieuwKlantEmailAdres As String)
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   KlantEmailAdresV = NieuwKlantEmailAdres
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
   KlantFaxNummer = KlantFaxNummerV
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
   KlantFaxNummerV = NieuwKlantFaxNummer
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

'Deze procedure stuurt de klant naam terug.
Public Property Get KlantNaam() As String
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   KlantNaam = KlantNaamV
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven klant naam vast.
Public Property Let KlantNaam(NieuweKlantNaam As String)
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   KlantNaamV = NieuweKlantNaam
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het klant nummer terug.
Public Property Get KlantNummer() As String
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   KlantNummer = KlantNummerV
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt het opgegeven klant nummer vast.
Public Property Let KlantNummer(NieuwKlantNummer As String)
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   KlantNummerV = NieuwKlantNummer
   If InStr(NrOpmaak, "K") > 0 Then MaakFaktuurNr
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
Public Property Get KlantTelefoonNummer() As String
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   KlantTelefoonNummer = KlantTelefoonNummerV
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt het opgegeven klant telefoonnummer vast.
Public Property Let KlantTelefoonNummer(NieuwKlantTelefoonNummer As String)
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   KlantTelefoonNummerV = NieuwKlantTelefoonNummer
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

'Deze procedure legt de opgegeven faktuur datum vast.
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

'Deze procedure stuurt het laatste faktuur subnummer terug.
Public Property Get LaatsteSubNr() As String
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   LaatsteSubNr = LaatsteSubNrV
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt het opgegeven faktuur subnummer vast.
Public Property Let LaatsteSubNr(NieuweLaatsteSubNr As String)
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   LaatsteSubNrV = NieuweLaatsteSubNr
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het maximum toegestane aantal faktuur subnummers terug.
Public Property Get MaxSubNrs() As String
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   MaxSubNrs = MaxSubNrsV
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt het maximum toegestane aantal faktuur subnummers vast.
Public Property Let MaxSubNrs(NieuwMaxSubNrs As String)
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   MaxSubNrsV = NieuwMaxSubNrs
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de faktuur nummer opmaak terug.
Public Property Get NrOpmaak() As String
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   NrOpmaak = NrOpmaakV
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven faktuur nummer opmaak vast.
Public Property Let NrOpmaak(NieuweNrOpmaak As String)
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   NrOpmaakV = NieuweNrOpmaak
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het porto bedrag terug.
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

'Deze procedure legt het opgegeven porto bedrag vast.
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

'Deze procedure stuurt het faktuur subnummer terug.
Public Property Get SubNr() As String
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   SubNr = SubNrV
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt het opgegeven faktuur subnummer vast.
Public Property Let SubNr(NieuwSubNr As String)
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   SubNrV = NieuwSubNr
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de faktuur subnummer tekst terug.
Public Property Get SubNrTekst() As String
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   SubNrTekst = SubNrTekstV
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt de faktuur subnummer periode terug.
Public Property Get SubNrPeriode() As String
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   SubNrPeriode = SubNrPeriodeV
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven faktuur subnummer tekst vast.
Public Property Let SubNrTekst(NieuweNrTekst As String)
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   SubNrTekstV = NieuweNrTekst
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de faktuur subnummer periode vast.
Public Property Let SubNrPeriode(NieuweSubNrPeriode As String)
On Error GoTo Fout
Dim Keuze As Long
EindeRoutine:
   SubNrPeriodeV = NieuweSubNrPeriode
Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het subtotaal bedrag terug.
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

'Deze procedure stuurt het totaal bedrag terug.
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
      Case FkTotaalPrijs
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
      Case FkTotaalPrijs
         Naam = "Totaal Prijs"
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
      Case FkTotaalPrijs
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

