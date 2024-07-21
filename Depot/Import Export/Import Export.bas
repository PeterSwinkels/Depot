Attribute VB_Name = "ImportExport"
'Deze module bevat de kern procedures voor dit programma.
Option Explicit

'Deze opsomming definieert de beschikbare databronnen.
Public Enum DatabronE
   DbKlanten    'Definieert de klanten databron.
   DbVoorraad   'Definieert de voorraad databron.
End Enum

Public ActieIsImporteren As Boolean       'Geeft aan of een bestand geïmporteerd of geëxporteerd wordt.
Public Bestand As String                  'Bevat de naam van een geïmporteerd of geëxporteerd bestand.
Public Databron As Object                 'Bevat de geselecteerde databron.
Public Veldtype(0 To 6) As Long           'Geeft het type van ieder veld in het geïmporteerde of geëxporteerde bestand op.
Public Veldscheidingstekens As String     'Bevat de tekens waarmee de velden in een geïmporteerd of geëxporteerd bestand worden gescheiden.

Public Klanten As New KlantenObject       'Bevat een verwijzing naar het klanten object.
Public Voorraad As New VoorraadObject     'Bevat een verwijzing naar het voorraad object.

'Deze procedure exporteert de gegevens.
Public Sub Exporteer()
On Error GoTo Fout
Dim Afbreken As Boolean
Dim BestandH As Integer
Dim Data As String
Dim Item As Long
Dim Keuze As Long
Dim Veld As Long
 
   If Not Bestand = vbNullString Then
      Afbreken = False
      BestandH = FreeFile()
      Open Bestand For Binary Lock Read Write As BestandH
         If LOF(BestandH) > 0 Then
            Afbreken = (MsgBox(Bestand & " overschrijven?", vbExclamation Or vbYesNo) = vbNo)
         End If
      Close BestandH
      
      If Not Afbreken Then
        Screen.MousePointer = vbHourglass
        BestandH = FreeFile()
        Open Bestand For Output Lock Read Write As BestandH
            For Item = 0 To Databron.AantalItems() - 1
               For Veld = 0 To Databron.AantalVelden() - 1
                  If Veldtype(Veld) > 0 Then
                     Data = Databron.Data(Veldtype(Veld) - 1, Item)
                     Print #BestandH, Data; Veldscheidingstekens;
                  End If
               Next Veld
               Print #BestandH,
            Next Item
        Close BestandH
      End If
   End If
   
EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure importeert de gegevens.
Public Sub Importeer()
On Error GoTo Fout
Dim BestandH As Integer
Dim Data As String
Dim Keuze As Long
Dim Veld As Long
Dim VeldEinde As Long
 
   If Not Bestand = vbNullString Then
      Screen.MousePointer = vbHourglass
      BestandH = FreeFile()
      Open Bestand For Input Lock Read Write As BestandH
         Do Until EOF(BestandH)
            Line Input #BestandH, Data
            Erase Buffer
            For Veld = 0 To Databron.AantalVelden() - 1
               VeldEinde = InStr(Data, Veldscheidingstekens)
               If VeldEinde = 0 Then VeldEinde = Len(Data) + 1
               Buffer(Veld) = Left$(Data, VeldEinde - 1)
               Data = Mid$(Data, VeldEinde + Len(Veldscheidingstekens))
            Next Veld
            Databron.VoegItemToe
            For Veld = 0 To Databron.AantalVelden() - 1
               If Veldtype(Veld) > 0 Then Databron.Data(Veldtype(Veld) - 1, Databron.AantalItems() - 1) = Buffer(Veld)
            Next Veld
         Loop
      Close BestandH
   End If
   
EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure wordt uitgevoerd wanneer dit programma wordt gestart.
Public Sub Main()
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim Afbreken As Boolean
Dim Keuze As Long

   Screen.MousePointer = vbHourglass
   
   ActieveMap = App.Path
   ActiefBestand = vbNullString
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   OpenInstellingen Afbreken
   If Not Afbreken Then
      VraagWachtwoord Afbreken
      If Not Afbreken Then
         GegevensVenster.Show
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

'Deze procedure opent de instellingen.
Public Sub OpenInstellingen(ByRef Afbreken As Boolean)
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim BestandH As Integer
Dim Keuze As Long
Dim Lengte As Long

   Afbreken = False
   ActieveMap = ".\Data\"
   ActiefBestand = "Depot.ins"
   BestandH = FreeFile()
   Open ".\Data\Depot.ins" For Binary Lock Read Write As BestandH
      If LOF(BestandH) > 0 Then
         Klanten.Afdrukken = Asc(Input$(1, BestandH))
         Voorraad.Afdrukken = Asc(Input$(1, BestandH))
         Lengte = Asc(Input$(1, BestandH)): Wachtwoord = Input$(Lengte, BestandH)
      End If
   Close BestandH
   
EindeProcedure:
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then
      Afbreken = True
      Resume EindeProcedure
   End If
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure vraagt om het beheerder wachtwoord.
Public Sub VraagWachtwoord(ByRef Afbreken As Boolean)
On Error GoTo Fout
Dim Keuze As Long

   If Wachtwoord = vbNullString Then
      Afbreken = True
      MsgBox "Geen wachtwoord ingesteld.", vbExclamation
   Else
      WachtwoordVenster.Show vbModal
   
      If IngevoerdWachtwoord = vbNullString Then
         Afbreken = True
      ElseIf Not IngevoerdWachtwoord = Wachtwoord Then
         MsgBox "Onjuist wachtwoord.", vbExclamation
         Afbreken = True
      End If
   End If

EindeRoutine:
   Exit Sub
   
Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

