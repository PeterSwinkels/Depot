VERSION 5.00
Begin VB.Form FaktuurInvoerVenster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artikel"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox LeverbaarVeld 
      Alignment       =   1  'Right Justify
      Caption         =   "Leverbaar:"
      Height          =   195
      Left            =   2520
      TabIndex        =   5
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1155
   End
   Begin VB.CheckBox UitVoorraadVeld 
      Alignment       =   1  'Right Justify
      Caption         =   "Uit Voorraad:"
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.TextBox ArtikelNaamVeld 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      MaxLength       =   250
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox StukprijsVeld 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      MaxLength       =   250
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox ArtikelnrVeld 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2160
      MaxLength       =   250
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox AantalVeld 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2160
      MaxLength       =   250
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton ActieKnop 
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton AnnulerenKnop 
      Cancel          =   -1  'True
      Caption         =   "&Annuleren"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label VeldNaam 
      Alignment       =   1  'Right Justify
      Caption         =   "Artikel:"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label VeldNaam 
      Alignment       =   1  'Right Justify
      Caption         =   "Stukprijs:"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label VeldNaam 
      Alignment       =   1  'Right Justify
      Caption         =   "Artikelnr.:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label VeldNaam 
      Alignment       =   1  'Right Justify
      Caption         =   "Aantal:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "FaktuurInvoerVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het artikel invoer venster voor fakturen.
Option Explicit

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub AantalVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   AantalVeld.SelStart = 0
   AantalVeld.SelLength = Len(AantalVeld.Text)
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure plaatst de ingevoerde gegevens in een buffer en sluit dit venster.
Private Sub ActieKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   Buffer(FkArtikelNr) = UCase$(ArtikelnrVeld.Text)
   Buffer(FkAantal) = MaakGetalOp(AantalVeld.Text)
   Buffer(FkAtikelNaam) = ArtikelNaamVeld.Text
   Buffer(FkStukprijs) = MaakGetalOp(StukprijsVeld.Text)
   
   If UitVoorraadVeld.Value = vbChecked Then
      Buffer(FkUitVoorraad) = "J"
   ElseIf UitVoorraadVeld.Value = vbUnchecked Then
      Buffer(FkUitVoorraad) = "N"
   End If
   
   If LeverbaarVeld.Value = vbChecked Then
      Buffer(FkLeverbaar) = "J"
   ElseIf LeverbaarVeld.Value = vbUnchecked Then
      Buffer(FkLeverbaar) = "N"
   End If
   
   Unload Me
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure breekt het invoeren van de gegevens af.
Private Sub AnnulerenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   Erase Buffer
   Unload Me
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub ArtikelNaamVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   ArtikelNaamVeld.SelStart = 0
   ArtikelNaamVeld.SelLength = Len(ArtikelNaamVeld.Text)
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub ArtikelnrVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   ArtikelnrVeld.SelStart = 0
   ArtikelnrVeld.SelLength = Len(ArtikelnrVeld.Text)
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

   AantalVeld.ToolTipText = "Het aantal artikelen."
   If ActieIsToeVoegen Then
      ActieKnop.Caption = "&Toevoegen"
      ActieKnop.ToolTipText = "Voegt het artikel toe."
   Else
      ActieKnop.Caption = "&Wijzigen"
      ActieKnop.ToolTipText = "Wijzigt het artikel."
   End If
   AnnulerenKnop.ToolTipText = "Sluit dit venster."
   ArtikelNaamVeld.ToolTipText = "De artikel naam."
   ArtikelnrVeld.ToolTipText = "Het artikelnummer."
   LeverbaarVeld.ToolTipText = "Geeft aan of het artikel leverbaar is."
   StukprijsVeld.ToolTipText = "De artikel stukprijs."
   UitVoorraadVeld.ToolTipText = "Geeft aan of het artikel uit de voorraad afkomstig is."

   ArtikelnrVeld.Text = Buffer(FkArtikelNr)
   AantalVeld.Text = Buffer(FkAantal)
   ArtikelNaamVeld.Text = Buffer(FkAtikelNaam)
   StukprijsVeld.Text = Buffer(FkStukprijs)
   If Faktuur.IsParticulier() Then
      VeldNaam(FkStukprijs).Enabled = True
      StukprijsVeld.Enabled = True
   End If
   
   If Buffer(FkUitVoorraad) = "J" Then
      UitVoorraadVeld.Value = vbChecked
   ElseIf Buffer(FkUitVoorraad) = "N" Then
      UitVoorraadVeld.Value = vbUnchecked
   End If
   
   If Buffer(FkLeverbaar) = "J" Then
      LeverbaarVeld.Value = vbChecked
   ElseIf Buffer(FkLeverbaar) = "N" Then
      LeverbaarVeld.Value = vbUnchecked
   End If
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub StukPrijsVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   StukprijsVeld.SelStart = 0
   StukprijsVeld.SelLength = Len(StukprijsVeld.Text)
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stelt de invoer velden in afhankelijk van of een artikel wel of niet uit de voorraad komt.
Private Sub UitVoorraadVeld_Click()
On Error GoTo Fout
Dim Keuze As Long

   Select Case UitVoorraadVeld.Value
      Case vbChecked
         VeldNaam(FkAtikelNaam).Enabled = False
         ArtikelNaamVeld.Enabled = False
         If Not Faktuur.IsParticulier() Then
            VeldNaam(FkStukprijs).Enabled = False
            StukprijsVeld.Enabled = False
         End If
      Case vbUnchecked
         VeldNaam(FkAtikelNaam).Enabled = True
         VeldNaam(FkStukprijs).Enabled = True
         ArtikelNaamVeld.Enabled = True
         StukprijsVeld.Enabled = True
   End Select
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

