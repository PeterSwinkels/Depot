Attribute VB_Name = "KernModule"
'Deze module bevat de kern procedures voor dit en andere programmas.
Option Explicit

'Deze opsomming definieert de voorwaarden waaronder items uit een databron worden verwijderd.
Public Enum VerwijderVoorwaardenE
   VVDubbelNummer   'Definieert de dubbele nummer verwijder voorwaarde.
   VVGeenNummer     'Definieert de geen nummer verwijder voorwaarde.
End Enum

'Deze opsomming definieert de zoekrichtingen.
Public Enum ZoekrichtingE
   ZrAchteruit    'Definieert de achteruit zoekrichting.
   ZrVooruit      'Definieert de vooruit zoekrichting.
End Enum

Public Const BACKUP_LABEL As String = "DEPOT BEHEERDER BACKUP" 'Definieert het label waarmee door dit programma gemaakte backups worden aangeduid.
Public Const EURO_TEKEN As String = "€"                        'Definieert het EURO symbool. (Unicode 0x20AC)
Public Const ONGELDIGE_TEKENS As String = "\/:*?""<>|"         'Definieert tekens die niet zijn toegestaan in faktuur of klant nummers.
Public Const PAD_SCHEIDINGS_TEKEN As String = "\"              'Definieert het teken dat wordt gebruikt om de namen in een pad te scheiden.

Public ActieIsToeVoegen As Boolean     'Geeft aan of gegevens toegevoegd of gewijzigd worden.
Public Buffer(0 To 6) As String        'Bevat ingevoerde gegevens die nog moeten worden verwerkt.
Public IngevoerdWachtwoord As String   'Bevat het door de gebruiker ingevoerde wachtwoord.
Public KlantenAfdrukken As Byte        'Geeft aan welke klant velden worden afgedrukt.
Public VoorraadAfdrukken As Byte       'Geeft aan welke voorraad lijst velden worden afgedrukt.
Public Wachtwoord As String            'Bevat beheerder wachtwoord.

'Deze procedure berekent het opgegeven percentage van het opgegeven getal en stuurt deze terug.
Public Function BerekenProcent(Percentage As Long, Getal As Single) As Single
On Error GoTo Fout
Dim Keuze As Long
Dim Procent As Long

   Procent = (Getal / 100) * Percentage
   
EindeRoutine:
   BerekenProcent = Procent
   Exit Function

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function
'Deze procedure vraagt de gebruiker of een actie na een leesfout afbroken moet worden en stuurt het resultaat terug.
Public Function BevestigAfbrekenNaLeesFout(Pad As String) As Boolean
On Error GoTo Fout
Dim Afbreken As Boolean
Dim Bericht As String
Dim Keuze As Long

   Bericht = "Gegevens in """ & Pad & """ zijn mogelijk niet juist ingelezen tijdens het openen." & vbCr
   Bericht = Bericht & "De mogelijk onjuiste gegevens toch opslaan?"
   Afbreken = (MsgBox(Bericht, vbExclamation Or vbYesNo Or vbDefaultButton2) = vbNo)
   
EindeRoutine:
   BevestigAfbrekenNaLeesFout = Afbreken
   Exit Function

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function

'Deze procedure drukt de opgegeven tekst met de opgegeven eigenschappen af.
Public Sub DrukAf(x As Single, y As Single, Tekst As String, Cursief As Boolean, Vet As Boolean)
On Error GoTo Fout
Dim Buffer As String
Dim Keuze As Long
Dim Lengte As Long
Dim RegelEinde As Long

   Printer.Font.Italic = Cursief
   Printer.Font.Bold = Vet
   
   Buffer = Tekst
   Do Until Buffer = vbNullString
      RegelEinde = InStr(Buffer$, vbCr)
      If RegelEinde > 0 Then RegelEinde = InStr(RegelEinde, Buffer, vbLf)
      Printer.CurrentX = x
      Printer.CurrentY = y
      If RegelEinde > 0 Then Lengte = RegelEinde Else Lengte = Len(Buffer)
      Printer.Print Left$(Buffer, Lengte);
      Buffer = Mid$(Buffer, Lengte + 1)
      If RegelEinde > 0 Then y = y + 1
      If y > Printer.ScaleHeight - 2 Then
         Printer.NewPage
         y = 1
      End If
   Loop
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure drukt de data uit de opgegeven bron af.
Public Sub DrukDataAf(DataBron As Object)
On Error GoTo Fout
Dim Afdrukken() As Boolean
Dim AfgedrukteVelden As Byte
Dim Data As String
Dim DataBreedte As Single
Dim Item As Long
Dim Keuze As Long
Dim Percentage As Single
Dim TotaleVeldBreedte As Long
Dim Veld As Long
Dim VeldX As Single
Dim VeldBreedte As Single

   Afdrukken() = HaalBits(DataBron.Afdrukken())
   AfgedrukteVelden = 0
   TotaleVeldBreedte = 0
   For Veld = 0 To DataBron.AantalVelden() - 1
      If Afdrukken(Veld) Then
         AfgedrukteVelden = AfgedrukteVelden + 1
         TotaleVeldBreedte = TotaleVeldBreedte + DataBron.VeldBreedte(Veld)
      End If
   Next Veld
 
   If AfgedrukteVelden = 0 Then
      MsgBox "Geen velden om af te drukken geselecteerd.", vbInformation
   Else
      StelPrinterIn
      
      Screen.MousePointer = vbHourglass
      
      Percentage = 100 / TotaleVeldBreedte
      Printer.CurrentY = 1
      Printer.Font.Bold = True
      Printer.Font.Size = 12
      Printer.CurrentX = 1
      For Veld = 0 To DataBron.AantalVelden() - 1
         If Afdrukken(Veld) Then
            VeldBreedte = BerekenProcent(Percentage * DataBron.VeldBreedte(Veld), (Printer.ScaleWidth - 2))
            VeldX = Printer.CurrentX
            Data = DataBron.VeldNaam(Veld)
            If Data = vbNullString Then Data = " "
            DataBreedte = VeldBreedte * (Len(Data) / Printer.TextWidth(Data))
            If DataBron.VeldRechtsUitlijnen(Veld) Then
               Data = Right$(Data, DataBreedte - 1)
               Printer.CurrentX = (Printer.CurrentX + (VeldBreedte - 1)) - Printer.TextWidth(Data)
            Else
               Data = Left$(Data, DataBreedte - 1)
            End If
            Printer.Print Data;
            Printer.CurrentX = VeldX + VeldBreedte
         End If
      Next Veld
      Printer.Print
      
      Printer.Font.Bold = False
      For Item = 0 To DataBron.AantalItems() - 1
         Printer.CurrentX = 1
         For Veld = 0 To DataBron.AantalVelden() - 1
            If Afdrukken(Veld) Then
               VeldBreedte = BerekenProcent(Percentage * DataBron.VeldBreedte(Veld), (Printer.ScaleWidth - 2))
               VeldX = Printer.CurrentX
               Data = DataBron.Data(Veld, Item)
               If Data = vbNullString Then Data = " "
               DataBreedte = VeldBreedte * (Len(Data) / Printer.TextWidth(Data))
               If DataBron.VeldRechtsUitlijnen(Veld) Then
                  Data = Right$(Data, DataBreedte - 1)
                  Printer.CurrentX = (Printer.CurrentX + (VeldBreedte - 1)) - Printer.TextWidth(Data)
               Else
                  Data = Left$(Data, DataBreedte - 1)
               End If
               Printer.Print Data;
               Printer.CurrentX = VeldX + VeldBreedte
            End If
         Next Veld
         Printer.Print
         If Printer.CurrentY > Printer.ScaleHeight - 2 Then
         Printer.NewPage
         Printer.CurrentY = 1
         End If
      Next Item
      Printer.EndDoc
   End If
EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure splitst de opgegeven byte op in individuele bits.
Public Function HaalBits(ByteV As Byte) As Boolean()
On Error GoTo Fout
Dim Bit As Long
Dim BitPatroonWaarde As Long
Dim Bits(0 To 7) As Boolean
Dim Keuze As Long

   BitPatroonWaarde = &H1&
   For Bit = LBound(Bits()) To UBound(Bits())
      Bits(Bit) = CBool((ByteV And BitPatroonWaarde) \ BitPatroonWaarde)
      BitPatroonWaarde = BitPatroonWaarde * &H2&
   Next Bit
   
EindeRoutine:
   HaalBits = Bits()
   Screen.MousePointer = vbDefault
   Exit Function
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function

'Deze procedure handelt eventuele fouten af.
Public Function HandelFoutAf(Optional ActieveMap As String = vbNullString, Optional ActiefBestand As String = vbNullString) As Long
Dim Bericht As String
Dim Foutcode As Long
Dim Keuze As Long
Dim Omschrijving As String
Dim VorigeMuisAanwijzer As Integer

   Foutcode = Err.Number
   Omschrijving = Err.Description
  
   On Error GoTo Fout
 
   If Not ActieveMap = vbNullString Then
      ActieveMap = LCase$(VoegScheidingstekenToe(ActieveMap))
   End If

   Omschrijving = Trim$(Omschrijving)
   If Not Right$(Omschrijving, 1) = "." Then
      Omschrijving = Omschrijving & "."
   End If
   
   VorigeMuisAanwijzer = Screen.MousePointer
   Screen.MousePointer = vbDefault
   
   Bericht = Omschrijving & vbCr
   Bericht = Bericht & "Foutcode: " & CStr(Foutcode)
   If Not ActieveMap = vbNullString Then Bericht = Bericht & vbCr & "Map: """ & ActieveMap & """"
   If Not ActiefBestand = vbNullString Then Bericht = Bericht & vbCr & "Bestand: """ & ActiefBestand & """"
   
   Keuze = MsgBox(Bericht, vbExclamation Or vbAbortRetryIgnore Or vbDefaultButton2)
   Select Case Keuze
      Case vbAbort
         Resume BreekProgrammaAf
      Case vbRetry
         Screen.MousePointer = VorigeMuisAanwijzer
      Case vbIgnore
         Resume Next
   End Select
   
   HandelFoutAf = Keuze
   Exit Function

BreekProgrammaAf:
   End
   
Fout:
   Resume BreekProgrammaAf
End Function

'Deze procedure maakt het opgegeven getal op.
Public Function MaakGetalOp(Getal As String) As String
On Error GoTo Fout
Dim Keuze As Long
Dim OpgemaaktGetal As String
Dim Teken As String
Dim TekenIndex As Long

   OpgemaaktGetal = vbNullString
   For TekenIndex = 1 To Len(Getal)
      Teken = Mid$(Getal, TekenIndex, 1)
      If InStr("0123456789.,", Teken) > 0 Then
         If Teken = "," Then Teken = "."
         If Teken = "." Then
            If InStr(TekenIndex + 1, Getal, ".") > 0 Or InStr(TekenIndex + 1, Getal, ",") > 0 Then
               Teken = vbNullString
            End If
         End If
         OpgemaaktGetal = OpgemaaktGetal & Teken
      End If
   Next TekenIndex
   If OpgemaaktGetal = vbNullString Then OpgemaaktGetal = "0"
   On Error GoTo Overloop
   OpgemaaktGetal = Trim$(Str$(CCur(Val(OpgemaaktGetal))))
   On Error GoTo Fout
EindeRoutine:
   MaakGetalOp = OpgemaaktGetal
   Exit Function

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
   Exit Function

Overloop:
   MsgBox "Het getal is te klein of te groot.", vbInformation
   OpgemaaktGetal = "0"
   Resume EindeRoutine
End Function

'Deze procedure voegt de opgegeven bits samen tot een byte en stuurt deze terug.
Public Function PlaatsBits(Bits() As Boolean) As Byte
On Error GoTo Fout
Dim Bit As Long
Dim ByteV As Byte
Dim Keuze As Long
Dim Vermenigvuldiger As Long

   Vermenigvuldiger = &H1&
   ByteV = &H0&
   For Bit = LBound(Bits()) To UBound(Bits())
      ByteV = ByteV Or (Abs(Bits(Bit)) * Vermenigvuldiger)
      Vermenigvuldiger = Vermenigvuldiger * &H2&
   Next Bit
   
EindeRoutine:
   PlaatsBits = ByteV
   Screen.MousePointer = vbDefault
   Exit Function
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function

'Deze procedure rond het opgegeven bedrag af en stuurt het resultaat terug.
Public Function RondAf(Bedrag As String) As String
On Error GoTo Fout
Dim AfgerondBedrag As String
Dim Keuze As Long

   AfgerondBedrag = Bedrag
   AfgerondBedrag = MaakGetalOp(AfgerondBedrag)
   AfgerondBedrag = Format$(Val(AfgerondBedrag), String$(Len(AfgerondBedrag), "#") & "0.00")
   Mid$(AfgerondBedrag, Len(AfgerondBedrag) - 2, 1) = "."

EindeRoutine:
   RondAf = AfgerondBedrag
   Exit Function
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function

'Deze procedure stelt de printer in.
Public Sub StelPrinterIn()
On Error GoTo Fout
Dim Keuze As Long

   Screen.MousePointer = vbHourglass
   
   With Printer
      .KillDoc
      .ColorMode = vbPRCMMonochrome
      .Copies = 1
      .DrawMode = vbCopyPen
      .DrawStyle = vbSolid
      .DrawWidth = 1
      .FillColor = vbWhite
      .Font.Name = "Arial"
      .ForeColor = vbBlack
      .Font.Bold = False
      .Font.Italic = False
      .Font.Strikethrough = False
      .Font.Underline = False
      .FontTransparent = True
      .Orientation = vbPRORPortrait
      .PaperSize = vbPRPSA4
      .ScaleMode = vbCharacters
      .TrackDefault = True
   End With
   
EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure toont informatie over dit programma.
Public Sub ToonProgrammaInformatie()
On Error GoTo Fout
Dim Bericht As String
Dim Keuze As Long

   Bericht = App.Title & vbCr
   Bericht = Bericht & "Door: " & App.CompanyName & vbCr
   Bericht = Bericht & "Versie: " & App.Major & "." & App.Minor & App.Revision & vbCr
   Bericht = Bericht & "***2003***"
   MsgBox Bericht, vbInformation
EindeRoutine:
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze proedure verplaatst een item in de opgegeven data van de opgegven bron naar het opgegeven doel.
Public Sub VerplaatsItem(Data() As String, Bron As Long, Doel As Long)
On Error GoTo Fout
Dim Item As Long
Dim Keuze As Long
Dim Veld As Long

   Screen.MousePointer = vbHourglass
   For Item = Bron + 1 To Doel
      For Veld = LBound(Data(), 1) To UBound(Data(), 1)
         Verwissel Data(Veld, Item), Data(Veld, Item - 1)
      Next Veld
   Next Item
EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure verwijdert indien aanwezig de bestandsnaam uit het opgegeven pad en stuurt het resultaat terug.
Public Function VerwijderBestandsnaam(Pad As String) As String
On Error GoTo Fout
Dim Keuze As Long
Dim PadZonderBestandsnaam As String
Dim Positie As Long
Dim VolgendePositie As Long

   PadZonderBestandsnaam = Pad
   Positie = 0
   Do
      VolgendePositie = InStr(Positie + 1, PadZonderBestandsnaam, PAD_SCHEIDINGS_TEKEN)
      If VolgendePositie = 0 Then Exit Do
      Positie = VolgendePositie
   Loop
   If Positie > 0 Then PadZonderBestandsnaam = Left$(PadZonderBestandsnaam, Positie)
   
EindeRoutine:
   VerwijderBestandsnaam = PadZonderBestandsnaam
   Exit Function
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function

'Deze procedure verwisselt de twee opgegeven waardes.
Public Sub Verwissel(Waarde1 As Variant, Waarde2 As Variant)
On Error GoTo Fout
Dim Keuze As Long
Dim Waarde3 As Variant

   Waarde3 = Waarde1
   Waarde1 = Waarde2
   Waarde2 = Waarde3
EindeRoutine:
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure voegt indien nodig een scheidingsteken toe aan het opgegeven pad en stuurt het resultaat terug.
Public Function VoegScheidingstekenToe(Pad As String) As String
On Error GoTo Fout
Dim Keuze As Long
Dim PadMetScheidingsTeken As String

   PadMetScheidingsTeken = Pad
   If Not Right$(Pad, Len(PAD_SCHEIDINGS_TEKEN)) = PAD_SCHEIDINGS_TEKEN Then
      PadMetScheidingsTeken = PadMetScheidingsTeken & PAD_SCHEIDINGS_TEKEN
   End If
  
EindeRoutine:
   VoegScheidingstekenToe = PadMetScheidingsTeken
   Exit Function
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function

'Deze procedure zet de bits in de opgegeven tekst om.
Public Function ZetBitsOm(Tekst As String) As String
On Error GoTo Fout
Dim Keuze As Long
Dim OmgezetteTekst As String
Dim TekenIndex As Long

   OmgezetteTekst = Tekst
   For TekenIndex = 1 To Len(OmgezetteTekst)
      Mid$(OmgezetteTekst, TekenIndex, 1) = Chr$(&HFF& Xor Asc(Mid$(OmgezetteTekst, TekenIndex, 1)))
   Next TekenIndex
   
EindeRoutine:
   ZetBitsOm = OmgezetteTekst
   Exit Function
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function

'Deze procedure doorzoekt de opgegeven databron naar gegevens die voldoen aan de opgegeven zoekcriteria.
Public Function ZoekItemIndex(DataBron As Object, Selectie As Long, Veld As Long, ZoekTekst As String, ZoekHeleVeld As Boolean, HoofdletterGevoelig As Boolean, Richting As ZoekrichtingE, ByRef Gevonden As Boolean) As Long
On Error GoTo Fout
Dim Data As String
Dim Einde As Long
Dim Item As Long
Dim Keuze As Long
Dim Stap As Long
Dim Start As Long

   If DataBron.AantalItems() = 0 Then
      Gevonden = False
      MsgBox "Geen gegevens om in te zoeken.", vbInformation
   Else
      Select Case Richting
         Case ZrAchteruit
            If Selectie = 0 Then
               Start = DataBron.AantalItems() - 1
            Else
               If Gevonden Then
                  Start = Selectie - 1
               Else
                  Start = Selectie
               End If
            End If
         Case ZrVooruit
            If Selectie < DataBron.AantalItems() - 1 Then
               If Gevonden Then
                  Start = Selectie + 1
               Else
                  Start = Selectie
               End If
            Else
               Start = 0
            End If
      End Select
      
      Gevonden = False
      If HoofdletterGevoelig Then
         ZoekTekst = UCase$(ZoekTekst)
      End If
      
      Select Case Richting
         Case ZrAchteruit
            Einde = 0
         Case ZrVooruit
            Einde = DataBron.AantalItems() - 1
      End Select
   
      Stap = Sgn(Einde - Start)
      If Stap = 0 Then Stap = 1
  
      For Item = Start To Einde Step Stap
         Data = DataBron.Data(Veld, Item)
         If HoofdletterGevoelig Then Data = UCase$(Data)
         If ZoekHeleVeld Then
            Gevonden = (ZoekTekst = Data)
         ElseIf Not ZoekHeleVeld Then
            Gevonden = (InStr(Data, ZoekTekst) > 0)
         End If
         If Gevonden Then Exit For
      Next Item
   End If
EindeRoutine:
   ZoekItemIndex = Item
   Exit Function
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function


