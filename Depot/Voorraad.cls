VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VoorraadObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Deze module bevat het voorraad object.
Option Explicit

'Deze opsomming definieert de beschikbare voorraad velden.
Public Enum VoorraadE
   VrArtikelnr           'Definieert een voorraad artikel nummer.
   VrNaam                'Definieert een voorraad artikel naam
   VrAantal              'Definieert een voorraad artikel aantal.
   VrAftotaal            'Definieert een voorraad artitkel er af totaal.
   VrBijTotaal           'Definieert een voorraad artitkel er bij totaal.
   VrStukprijs           'Definieert een voorraad artikel stukprijs.
End Enum

Private AantalItemsV As Long      'Bevat het aantal artikelen in de vooraad lijst.
Private AfdrukkenV As Byte        'Geeft aan welke velden worden afgedrukt.
Private DataV() As String         'Bevat de velden met gegevens.
Private LeesFout As Boolean       'Geeft aan of er een fout is opgetreden tijdens het lezen van het voorraad bestand.

'Deze procedure corrigeert de opgegeven artikel gegevens.
Public Sub Corrigeer(Artikel As Long)
On Error GoTo Fout
Dim Keuze As Long
   
   DataV(VrAantal, Artikel) = MaakGetalOp(Buffer(VrAantal))
   DataV(VrAftotaal, Artikel) = MaakGetalOp(Buffer(VrAftotaal))
   DataV(VrBijTotaal, Artikel) = MaakGetalOp(Buffer(VrBijTotaal))
EindeRoutine:
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure maakt artikel gegevens op.
Public Sub MaakItemsOp()
On Error GoTo Fout
Dim Artikel As Long
Dim Keuze As Long

   Screen.MousePointer = vbHourglass
   For Artikel = 0 To AantalItemsV - 1
      DataV(VrArtikelnr, Artikel) = Trim$(UCase$(DataV(VrArtikelnr, Artikel)))
      DataV(VrNaam, Artikel) = Trim$(DataV(VrNaam, Artikel))
      DataV(VrAantal, Artikel) = MaakGetalOp(DataV(VrAantal, Artikel))
      DataV(VrAftotaal, Artikel) = MaakGetalOp(DataV(VrAftotaal, Artikel))
      DataV(VrBijTotaal, Artikel) = MaakGetalOp(DataV(VrAftotaal, Artikel))
      DataV(VrStukprijs, Artikel) = RondAf(DataV(VrStukprijs, Artikel))
   Next Artikel

EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure opent de voorraad lijst.
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
   ActiefBestand = "Voorraad.dat"
   BestandH = FreeFile()
   LeesFout = False
   Open ".\Data\Voorraad.dat" For Binary Lock Read Write As BestandH
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

'Deze procedure maakt de voorraad lijst leeg.
Public Sub ResetGegevens()
On Error GoTo Fout
Dim Keuze As Long

   AantalItemsV = 0
   ReDim DataV(0 To 5, 0 To AantalItemsV) As String

EindeRoutine:
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure slaat voorraad lijst op.
Public Sub SlaGegevensOp()
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim Afbreken As Boolean
Dim Artikel As Long
Dim BestandH As Integer
Dim Keuze As Long
Dim Veld As Long

   ActieveMap = ".\Data\"
   ActiefBestand = "Voorraad.dat"
   Afbreken = False
   If LeesFout Then
      Afbreken = BevestigAfbrekenNaLeesFout(ActiefBestand)
   End If
   
   If Not Afbreken Then
     Screen.MousePointer = vbHourglass
     SetAttr ".\Data\Voorraad.dat", vbNormal
     BestandH = FreeFile()
     Open ".\Data\Voorraad.dat" For Output Lock Read Write As BestandH
         For Artikel = 0 To AantalItemsV - 1
            For Veld = LBound(DataV(), 1) To UBound(DataV(), 1)
               Print #BestandH, Chr$(Len(DataV(Veld, Artikel)));
               Print #BestandH, DataV(Veld, Artikel);
            Next Veld
         Next Artikel
     Close BestandH
     SetAttr ".\Data\Voorraad.dat", vbHidden Or vbReadOnly Or vbSystem
   End If
   
EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure sorteert de voorraad artikelen op artikel nummer.
Public Sub SorteerItems()
On Error GoTo Fout
Dim AnderArtikel As Long
Dim Artikel As Long
Dim Keuze As Long
Dim Veld As Long
  
   Screen.MousePointer = vbHourglass
   For Artikel = 0 To AantalItemsV - 1
      For AnderArtikel = 0 To AantalItemsV - 1
         If Not AnderArtikel = Artikel Then
            If DataV(VrArtikelnr, AnderArtikel) > DataV(VrArtikelnr, Artikel) Then
               For Veld = LBound(DataV(), 1) To UBound(DataV(), 1)
                  Verwissel DataV(Veld, AnderArtikel), DataV(Veld, Artikel)
               Next Veld
            End If
         End If
      Next AnderArtikel
   Next Artikel
   
EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure plaatst de standaard waardes in de invoer buffer.
Public Sub StelStandaardWaardesIn()
On Error GoTo Fout
Dim Keuze As Long
   
   Buffer(VrArtikelnr) = vbNullString
   Buffer(VrNaam) = vbNullString
   Buffer(VrAantal) = "0"
   Buffer(VrAftotaal) = "0"
   Buffer(VrBijTotaal) = "0"
   Buffer(VrStukprijs) = RondAf("0")

EindeRoutine:
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure verwijdert het opgegeven artikel uit de voorraad lijst.
Public Sub VerwijderItem(Artikel As Long)
On Error GoTo Fout
Dim Keuze As Long
   
   If Not AantalItemsV = 0 Then
      DataV(VrAantal, Artikel) = "0"
      DataV(VrAftotaal, Artikel) = "0"
      DataV(VrBijTotaal, Artikel) = "0"
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

'Deze procedure verwijdert de artikelen die voldoen aan de opgegeven voorwaarde.
Public Sub VerwijderItems(Voorwaarde As VerwijderVoorwaardenE)
On Error GoTo Fout
Dim Artikel As Long
Dim Keuze As Long

   Screen.MousePointer = vbHourglass
   Artikel = 0
   Select Case Voorwaarde
      Case VVDubbelNummer
         Do Until Artikel > AantalItemsV - 1
            If DataV(VrArtikelnr, Artikel) = DataV(VrArtikelnr, Artikel + 1) Then
               VerwijderItem Artikel
               Artikel = 0
            Else
               Artikel = Artikel + 1
            End If
         Loop
      Case VVGeenNummer
         Do Until Artikel > AantalItemsV - 1
            If DataV(VrArtikelnr, Artikel) = vbNullString Then
               VerwijderItem Artikel
               Artikel = 0
            Else
               Artikel = Artikel + 1
            End If
         Loop
   End Select
    
EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure verwijdert of overschrijft een artikel.
Public Sub VerwijderOudArtikel()
On Error GoTo Fout
Dim Artikel As Long
Dim Keuze As Long

   Screen.MousePointer = vbHourglass
   For Artikel = 0 To AantalItemsV - 2
      If DataV(VrArtikelnr, Artikel) = DataV(VrArtikelnr, Artikel + 1) Then
         If MsgBox("Dit artikel is al ingevoerd. Vervangen?", vbExclamation Or vbYesNo) = vbYes Then
            VerwijderItem Artikel
         Else
            VerwijderItem Artikel + 1
         End If
         Exit For
      End If
   Next Artikel

EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure voegt een artikel toe aan de vooraad lijst.
Public Sub VoegItemToe()
On Error GoTo Fout
Dim Keuze As Long
   
   AantalItemsV = AantalItemsV + 1
   ReDim Preserve DataV(0 To 5, 0 To AantalItemsV) As String

EindeRoutine:
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure wijzigt het opgegeven artikel.
Public Sub WijzigItem(Artikel As Long)
On Error GoTo Fout
Dim Keuze As Long
   
   DataV(VrArtikelnr, Artikel) = Buffer(VrArtikelnr)
   DataV(VrNaam, Artikel) = Buffer(VrNaam)
   DataV(VrStukprijs, Artikel) = RondAf(Buffer(VrStukprijs))
   If Val(DataV(VrAantal, Artikel)) - Val(Buffer(VrAftotaal)) < 0 Then
      MsgBox "Onvoldoende artikelen in voorraad.", vbExclamation
      If DataV(VrAftotaal, Artikel) = vbNullString Then DataV(VrAftotaal, Artikel) = "0"
   Else
      DataV(VrAantal, Artikel) = Val(DataV(VrAantal, Artikel)) - Val(Buffer(VrAftotaal))
      DataV(VrAftotaal, Artikel) = Val(DataV(VrAftotaal, Artikel)) + Val(Buffer(VrAftotaal))
   End If
   DataV(VrAantal, Artikel) = Val(DataV(VrAantal, Artikel)) + Val(Buffer(VrBijTotaal))
   DataV(VrBijTotaal, Artikel) = Val(DataV(VrBijTotaal, Artikel)) + Val(Buffer(VrBijTotaal))

EindeRoutine:
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stuurt het aantal artikelen in de voorraad terug.
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
   AantalVelden = 6
   Exit Property
   
Fout:
   Keuze = HandelFoutAf()
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
   Keuze = HandelFoutAf()
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

   DataV(Veld, Artikel) = NieuweData
   
EindeRoutine:
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
   TotaleVeldBreedte = 7
   Exit Property
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze eigenschap procedure stuurt terug hoeveel van de totale veld breedte een veld van het opgegeven veld type in beslag neemt.
Public Property Get VeldBreedte(Veld As Long) As Long
On Error GoTo Fout
Dim Breedte As Long
Dim Keuze As Long

   Select Case Veld
      Case VrArtikelnr
         Breedte = 1
      Case VrNaam
         Breedte = 2
      Case VrAantal
         Breedte = 1
      Case VrAftotaal
         Breedte = 1
      Case VrBijTotaal
         Breedte = 1
      Case VrStukprijs
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
      Case VrArtikelnr
         Naam = "Artikelnr."
      Case VrNaam
         Naam = "Artikel"
      Case VrAantal
         Naam = "Aantal"
      Case VrAftotaal
         Naam = "Er Af Totaal"
      Case VrBijTotaal
         Naam = "Er Bij Totaal"
      Case VrStukprijs
         Naam = "Stukprijs"
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

'Deze eigenschap procedure geeft aan of gegevens in het opgegeven veld type rechts uitgelijnd moeten worden.
Public Property Get VeldRechtsUitlijnen(Veld As Long) As Boolean
On Error GoTo Fout
Dim Keuze As Long
Dim RechtsUitlijnen As Boolean

   Select Case Veld
      Case VrArtikelnr
         RechtsUitlijnen = False
      Case VrNaam
         RechtsUitlijnen = False
      Case VrAantal
         RechtsUitlijnen = True
      Case VrAftotaal
         RechtsUitlijnen = True
      Case VrBijTotaal
         RechtsUitlijnen = True
      Case VrStukprijs
         RechtsUitlijnen = True
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
   LeesFout = False
   ReDim DataV(0 To 5, 0 To AantalItemsV) As String
   
EindeRoutine:
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

