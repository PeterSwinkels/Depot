VERSION 5.00
Begin VB.Form VoorraadInvoerVenster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artikel"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10.125
   ScaleMode       =   4  'Character
   ScaleWidth      =   38.125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox AantalErafVeld 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      MaxLength       =   250
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox AantalErbijVeld 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      MaxLength       =   250
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton AnnulerenKnop 
      Cancel          =   -1  'True
      Caption         =   "&Annuleren"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton ActieKnop 
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox StukprijsVeld 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      MaxLength       =   250
      TabIndex        =   4
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox ArtikelNaamVeld 
      Height          =   285
      Left            =   2280
      MaxLength       =   250
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox ArtikelnrVeld 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      MaxLength       =   250
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label AantalErafLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Aantal Eraf:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label StukprijsLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Stukprijs:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label AantalErBijLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Aantal Erbij:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label ArtikelLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Artikel:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label ArtikelNrLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Artikelnr.:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "VoorraadInvoerVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het voorraad invoer venster.
Option Explicit

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub AantalErAfVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   AantalErAfVeld.SelStart = 0
   AantalErAfVeld.SelLength = Len(AantalErAfVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub AantalErBijVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   AantalErBijVeld.SelStart = 0
   AantalErBijVeld.SelLength = Len(AantalErBijVeld.Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waardes vast.
Private Sub ActieKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   Buffer(VrArtikelnr) = UCase$(ArtikelnrVeld.Text)
   Buffer(VrNaam) = ArtikelNaamVeld.Text
   Buffer(VrAftotaal) = MaakGetalOp(AantalErAfVeld.Text)
   Buffer(VrBijTotaal) = MaakGetalOp(AantalErBijVeld.Text)
   Buffer(VrStukprijs) = StukprijsVeld.Text

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

   ArtikelnrVeld.Text = Buffer(VrArtikelnr)
   ArtikelNaamVeld.Text = Buffer(VrNaam)
   AantalErAfVeld.Text = "0"
   AantalErBijVeld.Text = "0"
   StukprijsVeld.Text = Buffer(VrStukprijs)

   If ActieIsToeVoegen Then
      ActieKnop.Caption = "&Toevoegen"
   Else
      ActieKnop.Caption = "&Wijzigen"
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

