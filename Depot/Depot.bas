Attribute VB_Name = "DepotBeheerderModule"
'Deze module bevat de kern procedures voor dit programma.
Option Explicit

'Deze opsomming definieert de beschikbare databronnen.
Public Enum DataBronE
   DbFakturen   'Definieert de fakturen databron.
   DbFaktuur    'Definieert de faktuur databron.
   DbKlanten    'Definieert de klanten databron.
   DbVoorraad   'Definieert de voorraad databron.
End Enum

Public BackupMap As String                'Bevat het pad van de map met de laatst gemaakte backup.
Public GebruikerIsBeheerder As Boolean    'Geeft aan of ingelogde gebruiker een beheerder is.
Public GegevensOpslaan As Boolean         'Geeft aan of gegevens bij het afsluiten worden opgeslagen.
Public GekopieerdArtikel As String        'Bevat het nummer van artikel dat is gekopieerd uit de voorraad lijst.

Public Fakturen As New FakturenObject     'Bevat een verwijzing naar het fakturen object.
Public Faktuur As New FaktuurObject       'Bevat een verwijzing naar het faktuur object.
Public Klanten As New KlantenObject       'Bevat een verwijzing naar het klanten object.
Public Voorraad As New VoorraadObject     'Bevat een verwijzing naar het voorraad object.

'Deze procedure controleert of gebruiker een beheerder is en geeft een melding wanneer dit niet het geval is.
Public Function Beheerder() As Boolean
On Error GoTo Fout
Dim Keuze As Long

   If Not GebruikerIsBeheerder Then
      MsgBox "Alleen de beheerder kan deze functie gebruiken.", vbExclamation
   End If
  
EindeRoutine:
   Beheerder = GebruikerIsBeheerder
   Exit Function

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function

'Deze procedure wordt uitgevoerd wanneer dit programma wordt gestart.
Public Sub Main()
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim Keuze As Long

   Screen.MousePointer = vbHourglass
   
   ActiefBestand = vbNullString
   GebruikerIsBeheerder = False
   GegevensOpslaan = True
   GekopieerdArtikel = vbNullString
   
   ActieveMap = App.Path
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   ActieveMap = ".\Data"
   If Dir$(".\Data\", vbArchive Or vbDirectory Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = vbNullString Then
      MkDir ".\Data"
   End If
   
   OpenInstellingen
   Faktuur.OpenFaktuurTekst
   Klanten.OpenGegevens
   Voorraad.OpenGegevens
   
   VraagWachtwoord
   
   DepotBeheerderVenster.Show
EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure opent de instellingen.
Public Sub OpenInstellingen()
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim BestandH As Integer
Dim Keuze As Long
Dim Lengte As Long

   ActieveMap = ".\Data\"
   ActiefBestand = "Depot.ins"
   BestandH = FreeFile()
   Open ".\Data\Depot.ins" For Binary Lock Read Write As BestandH
      If LOF(BestandH) = 0 Then
         Faktuur.DatumOpmaak = "DMJ"
         Faktuur.LaatsteSubNr = "0"
         Faktuur.MaxSubNrs = "0"
         Faktuur.SubNrPeriode = "D"
         Klanten.Afdrukken = 255
         Voorraad.Afdrukken = 255
      Else
         Klanten.Afdrukken = Asc(Input$(1, BestandH))
         Voorraad.Afdrukken = Asc(Input$(1, BestandH))
         Lengte = Asc(Input$(1, BestandH)): Wachtwoord = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): BackupMap = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): Faktuur.DepotNaam = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): Faktuur.DatumOpmaak = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): Faktuur.MaxSubNrs = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): Faktuur.NrOpmaak = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): Faktuur.SubNrTekst = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): Faktuur.LaatsteDatum = Input$(Lengte, BestandH)
         Lengte = Asc(Input$(1, BestandH)): Faktuur.LaatsteSubNr = Input$(Lengte, BestandH)
         Faktuur.SubNrPeriode = Input$(1, BestandH)
      End If
   Close BestandH
EindeRoutine:
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om gegevens en instellingen te bewaren.
Public Sub SlaGegevensOp()
On Error GoTo Fout
Dim Keuze As Long

   If GegevensOpslaan Then
      Faktuur.SlaFaktuurOp
      Faktuur.SlaFaktuurTekstOp
      Klanten.SlaGegevensOp
      Voorraad.SlaGegevensOp
      SlaInstellingenOp
   End If
EindeRoutine:
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure slaat de instellingen op.
Public Sub SlaInstellingenOp()
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim BestandH As Integer
Dim Keuze As Long

   ActieveMap = ".\Data\"
   ActiefBestand = "Depot.ins"
   SetAttr ".\Data\Depot.ins", vbNormal
   BestandH = FreeFile()
   Open ".\Data\Depot.ins" For Output Lock Read Write As BestandH
      Print #BestandH, Chr$(Klanten.Afdrukken);
      Print #BestandH, Chr$(Voorraad.Afdrukken);
      Print #BestandH, Chr$(Len(Wachtwoord)); Wachtwoord;
      Print #BestandH, Chr$(Len(BackupMap)); BackupMap;
      Print #BestandH, Chr$(Len(Faktuur.DepotNaam())); Faktuur.DepotNaam();
      Print #BestandH, Chr$(Len(Faktuur.DatumOpmaak())); Faktuur.DatumOpmaak();
      Print #BestandH, Chr$(Len(Faktuur.MaxSubNrs())); Faktuur.MaxSubNrs();
      Print #BestandH, Chr$(Len(Faktuur.NrOpmaak())); Faktuur.NrOpmaak();
      Print #BestandH, Chr$(Len(Faktuur.SubNrTekst())); Faktuur.SubNrTekst();
      Print #BestandH, Chr$(Len(Faktuur.LaatsteDatum())); Faktuur.LaatsteDatum();
      Print #BestandH, Chr$(Len(Faktuur.LaatsteSubNr())); Faktuur.LaatsteSubNr();
      Print #BestandH, Faktuur.SubNrPeriode();
   Close BestandH
   SetAttr ".\Data\Depot.ins", vbHidden Or vbReadOnly Or vbSystem
EindeRoutine:
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stelt een nieuw beheerder wachtwoord in.
Public Sub StelWachtwoordIn(NieuwWachtwoord As String)
On Error GoTo Fout
Dim Keuze As Long

   If Not (IngevoerdWachtwoord = vbNullString And NieuwWachtwoord = vbNullString) Then
      If IngevoerdWachtwoord = Wachtwoord Then
         Wachtwoord = ZetBitsOm(NieuwWachtwoord)
      Else
         MsgBox "Onjuist wachtwoord.", vbExclamation
      End If
   End If
EindeRoutine:
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stuurt de waarde vertegenwoordigd door het opgegeven symbool terug.
Public Function SymboolWaarde(Tekst As String, SymboolIndex As Long) As String
On Error GoTo Fout
Dim Symbool As String
Dim Keuze As Long
Dim Waarde As String

   Waarde = vbNullString
   Symbool = Mid$(Tekst, SymboolIndex, 1)
   Select Case Symbool
      Case "D"
         Waarde = Day(Date)
      Case "J"
         Waarde = Year(Date)
      Case "j"
         Waarde = Mid$(Year(Date), 3)
      Case "K"
         Waarde = Faktuur.KlantNummer()
      Case "M"
         Waarde = Month(Date)
      Case "T"
         Waarde = Faktuur.SubNrTekst()
      Case "#"
         Waarde = Format$(Faktuur.SubNr(), String$(Len(Faktuur.MaxSubNrs()), "0"))
   End Select
 
EindeRoutine:
   SymboolWaarde = Waarde
   Exit Function
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function

'Deze procedure vraagt om het beheerder wachtwoord.
Public Sub VraagWachtwoord()
On Error GoTo Fout
Dim Keuze As Long

   GebruikerIsBeheerder = False

   If Not Wachtwoord = vbNullString Then
      InloggenVenster.Show vbModal
      If Not IngevoerdWachtwoord = vbNullString Then
         If IngevoerdWachtwoord = Wachtwoord Then
            GebruikerIsBeheerder = True
         Else
            MsgBox "Onjuist wachtwoord. Ingelogd als gebruiker.", vbExclamation
         End If
      End If
   End If
   
   DepotBeheerderVenster.Caption = "Depot Beheerder - Ingelogd als "
   If GebruikerIsBeheerder Then
      DepotBeheerderVenster.Caption = DepotBeheerderVenster.Caption & "beheerder"
   Else
      DepotBeheerderVenster.Caption = DepotBeheerderVenster.Caption & "gebruiker"
   End If
   DepotBeheerderVenster.Caption = DepotBeheerderVenster.Caption & "."
EindeRoutine:
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

