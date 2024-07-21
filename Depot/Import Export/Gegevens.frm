VERSION 5.00
Begin VB.Form GegevensVenster 
   Caption         =   "Import/Export"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   6375
   ClipControls    =   0   'False
   Icon            =   "Gegevens.frx":0000
   ScaleHeight     =   15.563
   ScaleMode       =   4  'Character
   ScaleWidth      =   53.125
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox KnoppenBalk 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   375
      Left            =   1440
      ScaleHeight     =   375
      ScaleWidth      =   4815
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3240
      Width           =   4815
      Begin VB.CommandButton VerwijderenKnop 
         Caption         =   "&Verwijderen"
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton ToevoegenKnop 
         Caption         =   "&Toevoegen"
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton WijzigenKnop 
         Caption         =   "&Wijzigen"
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton AfdrukkenKnop 
         Caption         =   "&Afdrukken"
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
   End
   Begin ImportExportProgramma.LijstObject ItemsLijst 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _extentx        =   10821
      _extenty        =   5741
   End
   Begin VB.Menu GegevensMenu 
      Caption         =   "&Gegevens"
      Begin VB.Menu NieuwMenu 
         Caption         =   "&Nieuw"
      End
      Begin VB.Menu OpenenMenu 
         Caption         =   "&Openen"
      End
      Begin VB.Menu OpslaanMenu 
         Caption         =   "O&pslaan"
      End
      Begin VB.Menu ExporterenMenu 
         Caption         =   "&Exporteren"
      End
      Begin VB.Menu ImporterenMenu 
         Caption         =   "&Importeren"
      End
   End
   Begin VB.Menu BewerkenMenu 
      Caption         =   "&Bewerken"
      Begin VB.Menu MaakItemsOpMenu 
         Caption         =   "&Maak Items Op"
      End
      Begin VB.Menu SorteerItemsMenu 
         Caption         =   "&Sorteer Items"
      End
      Begin VB.Menu VerwijderMenu 
         Caption         =   "&Verwijder"
         Begin VB.Menu VerwijderDubbeleItemsMenu 
            Caption         =   "&Dubbele Items"
         End
         Begin VB.Menu VerwijderItemsZonderCodeMenu 
            Caption         =   "&Items Zonder Code"
         End
      End
   End
   Begin VB.Menu DatabronMenu 
      Caption         =   "&Databron"
      Begin VB.Menu KlantenMenu 
         Caption         =   "&Klanten"
      End
      Begin VB.Menu VoorraadMenu 
         Caption         =   "&Voorraad"
      End
   End
   Begin VB.Menu ZoekenMenu 
      Caption         =   "&Zoeken"
   End
   Begin VB.Menu InformatieMenu 
      Caption         =   "&Informatie"
   End
End
Attribute VB_Name = "GegevensVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het hoofdvenster.
Option Explicit

'Deze procedure drukt na bevestiging van de gebruiker de gegevens af.
Private Sub AfdrukkenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   If MsgBox("Gegevens afdrukken?", vbQuestion Or vbYesNo) = vbYes Then
      DrukDataAf Databron
   End If
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure opent het export venster.
Private Sub ExporterenMenu_Click()
On Error GoTo Fout
Dim Keuze As Long

   ActieIsImporteren = False
   ImportExportVenster.Show vbModal
   Exporteer
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stelt dit venster in.
Private Sub Form_Initialize()
On Error GoTo Fout
Dim Keuze As Long

   Me.WindowState = vbMaximized
   StelDatabronIn DbKlanten
   
   Databron.ResetGegevens
   Databron.OpenGegevens

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

   AfdrukkenKnop.ToolTipText = "Drukt de gegevens af."
   ToevoegenKnop.ToolTipText = "Voegt een item toe."
   VerwijderenKnop.ToolTipText = "Verwijdert een item."
   WijzigenKnop.ToolTipText = "Wijzigt een item."
   
   Me.Width = Screen.Width / 1.5
   Me.Height = Screen.Height / 1.5
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure past de objecten in dit venster aan de nieuwe afmetingen aan.
Private Sub Form_Resize()
On Error Resume Next
   
   ItemsLijst.Width = Me.ScaleWidth - 2
   ItemsLijst.Height = Me.ScaleHeight - 3
   KnoppenBalk.Left = Me.ScaleWidth - KnoppenBalk.Width - 1
   KnoppenBalk.Top = Me.ScaleHeight - 2
   
   ItemsLijst.WerkLijstBij
End Sub

'Deze procedure sluit alle vensters wanneer dit venster wordt gesloten.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Fout
Dim FormV As Variant
Dim Keuze As Long

   For Each FormV In Forms
      Unload FormV
   Next FormV
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure opent het importvenster.
Private Sub ImporterenMenu_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   ActieIsImporteren = True
   ImportExportVenster.Show vbModal
   Importeer
   ItemsLijst.WerkLijstBij
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure opent het programma informatievenster.
Private Sub InformatieMenu_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   ToonProgrammainformatie

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stelt de klanten lijst in als databron.
Private Sub KlantenMenu_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   StelDatabronIn DbKlanten
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft opdracht om de gegevens op te maken.
Private Sub MaakItemsOpMenu_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   Databron.MaakItemsOp
   ItemsLijst.WerkLijstBij

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze proedure verwijdert na bevestiging van de gebruiker alle gegevens uit een databron.
Private Sub NieuwMenu_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   If MsgBox("Huidige gegevens verwijderen?", vbQuestion Or vbYesNo) = vbYes Then
      Databron.ResetGegevens
      ItemsLijst.WerkLijstBij
   End If
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze proedure verwijdert na bevestiging van de gebruiker alle gegevens uit een databron en laadt de opgeslagen gegevens.
Private Sub OpenenMenu_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   If MsgBox("Huidige gegevens verwijderen en de opgeslagen gegevens laden?", vbQuestion Or vbYesNo) = vbYes Then
      Databron.ResetGegevens
      Databron.OpenGegevens
      ItemsLijst.WerkLijstBij
   End If
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure bewaart na bevestiging van de gebruiker de gegevens.
Private Sub OpslaanMenu_Click()
On Error GoTo Fout
Dim Keuze As Long
    
    If MsgBox("Oude gegevens overschrijven?", vbExclamation Or vbYesNo) = vbYes Then
       Databron.SlaGegevensOp
    End If
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om de gegevens te sorteren.
Private Sub SorteerItemsMenu_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   Databron.SorteerItems
   ItemsLijst.WerkLijstBij
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de opgegeven databron vast.
Private Sub StelDatabronIn(DataBronNr As Long)
On Error GoTo Fout
Dim Keuze As Long
   
   Select Case DataBronNr
      Case DbKlanten
         KlantenMenu.Checked = True
         VoorraadMenu.Checked = False
         Set Databron = Klanten
      Case DbVoorraad
         KlantenMenu.Checked = False
         VoorraadMenu.Checked = True
         Set Databron = Voorraad
   End Select
   
   Set ItemsLijst.Databron = Databron
   
   ItemsLijst.WerkLijstBij
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure voegt een item toe aan de gegevens uit een databron.
Private Sub ToevoegenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   Databron.StelStandaardWaardesIn
   
   ActieIsToeVoegen = True
   InvoerVenster.Show vbModal
   
   If Not Buffer(LBound(Buffer())) = vbNullString Then
      Databron.VoegItemToe
      ItemsLijst.Selectie = Databron.AantalItems() - 1
      Databron.WijzigItem ItemsLijst.Selectie()
      ItemsLijst.WerkLijstBij
   End If
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om de dubbele gegevens te verwijderen uit een databron.
Private Sub VerwijderDubbeleItemsMenu_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   Databron.VerwijderItems VVDubbelNummer
   ItemsLijst.WerkLijstBij
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om de geselecteerde gegevens te verwijderen uit een databron.
Private Sub VerwijderenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   Databron.VerwijderItem ItemsLijst.Selectie()
   ItemsLijst.WerkLijstBij
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om de gegevens zonder nummer te verwijderen uit een databron.
Private Sub VerwijderItemsZonderCodeMenu_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   Databron.VerwijderItems VVGeenNummer
   ItemsLijst.WerkLijstBij
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert de voorraad gegevens databron.
Private Sub VoorraadMenu_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   StelDatabronIn DbVoorraad
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure wijzigt een item in een databron.
Private Sub WijzigenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long
Dim Veld As Long

   If Not Databron.AantalItems() = 0 Then
      For Veld = 0 To Databron.AantalVelden() - 1
         Buffer(Veld) = Databron.Data(Veld, ItemsLijst.Selectie())
      Next Veld
      
      ActieIsToeVoegen = False
      InvoerVenster.Show vbModal
      
      If Not Buffer(LBound(Buffer())) = vbNullString Then
         Databron.WijzigItem ItemsLijst.Selectie()
         ItemsLijst.WerkLijstBij
      End If
   End If
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure toont het zoekvenster.
Private Sub ZoekenMenu_Click()
On Error GoTo Fout
Dim Keuze As Long
   
   ZoekVenster.Show
   ZoekVenster.ZOrder
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

