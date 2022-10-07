VERSION 5.00
Begin VB.Form ImportExportVenster 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   ClipControls    =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   17.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   40.25
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox VeldTekenLijst 
      Height          =   315
      ItemData        =   "Import Export.frx":0000
      Left            =   2760
      List            =   "Import Export.frx":0002
      TabIndex        =   4
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox BestandVeld 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   2760
      Width           =   3735
   End
   Begin VB.CommandButton AnnulerenKnop 
      Cancel          =   -1  'True
      Caption         =   "&Annuleren"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton ActieKnop 
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.DirListBox MappenLijst 
      Height          =   1890
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.DriveListBox StationLijst 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   2280
      Width           =   3255
   End
   Begin VB.FileListBox BestandenLijst 
      Height          =   1845
      Hidden          =   -1  'True
      Left            =   1440
      System          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.ComboBox VeldTypeLijst 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Text            =   "(Leeg)"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label VeldenScheidenMetDezeTekensLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Velden scheiden met deze tekens:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label BestandLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Bestand:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label VeldVolgordeLabel 
      Caption         =   "Veld volgorde:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "ImportExportVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het import/export venster.
Option Explicit

'Deze procedure legt de invoer vast.
Private Sub HaalInvoer()
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim Keuze As Long

   If Not BestandVeld.Text = vbNullString Then
      If InStr(BestandVeld.Text, PAD_SCHEIDINGS_TEKEN) = 0 Then
         BestandVeld.Text = VoegScheidingstekenToe(MappenLijst.Path) & BestandVeld.Text
      End If
      BestandVeld.Text = LCase$(BestandVeld.Text)
      Bestand = BestandVeld.Text
      ActieveMap = MappenLijst.Path
      ActiefBestand = Bestand
   End If
   
   Select Case VeldTekenLijst.ListIndex
      Case 1
         VeldScheidTekens = " "
      Case 2
         VeldScheidTekens = vbTab
      Case Else
         VeldScheidTekens = VeldTekenLijst.Text
   End Select
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de invoer vast en sluit dit venster.
Private Sub ActieKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   HaalInvoer
   Unload Me
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure sluit dit venster.
Private Sub AnnulerenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   Bestand = vbNullString
   Unload Me
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure voegt het pad toe aan de naam van het geselecteerde bestand.
Private Sub BestandenLijst_Click()
On Error GoTo Fout
Dim Keuze As Long

   BestandVeld.Text = VoegScheidingstekenToe(MappenLijst.Path) & BestandenLijst.List(BestandenLijst.ListIndex)
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om de invoer vast te leggen.
Private Sub BestandVeld_LostFocus()
On Error GoTo Fout
Dim Keuze As Long

   HaalInvoer
   
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
Dim Veld As Long
Dim VeldTypeNr As Long

   AnnulerenKnop.ToolTipText = "Sluit dit venster."
   BestandenLijst.ToolTipText = "De bestanden in de geselecteerde map."
   BestandVeld.ToolTipText = "Het te exporteren/importeren bestand."
   MappenLijst.ToolTipText = "De mappen in het geselecteerde station."
   StationLijst.ToolTipText = "De beschikbare stations."
   VeldTekenLijst.ToolTipText = "Geeft aan welke tekens de velden van elkaar scheiden."
   
   If ActieIsImporteren Then
      ActieKnop.ToolTipText = "Importeert de gegevens."
      Me.Caption = "Importeren"
   Else
      ActieKnop.ToolTipText = "Exporteert de gegevens."
      Me.Caption = "Exporteren"
   End If
   
   ActieKnop.Caption = "&" & Me.Caption
   
   BestandVeld.Text = Bestand
   
   MappenLijst.Path = VerwijderBestandsnaam(Bestand)
   
   For Veld = 0 To DataBron.AantalVelden() - 1
      If Veld > 0 Then Load VeldTypeLijst(Veld)
      VeldTypeLijst(Veld).Clear
      VeldTypeLijst(Veld).AddItem "(Leeg)"
      VeldTypeLijst(Veld).ToolTipText = "Kent een veld toe aan de gegevens."
      VeldTypeLijst(Veld).Top = (Veld * 1.25) + 2
      VeldTypeLijst(Veld).Visible = True
      For VeldTypeNr = 0 To DataBron.AantalVelden() - 1
         VeldTypeLijst(Veld).AddItem DataBron.VeldNaam(VeldTypeNr)
      Next VeldTypeNr
      VeldTypeLijst(Veld).ListIndex = VeldType(Veld)
   Next Veld
   
   VeldTekenLijst.Clear
   VeldTekenLijst.AddItem "(Andere Tekens)"
   VeldTekenLijst.AddItem "Spatie"
   VeldTekenLijst.AddItem "Tab"
   
   Select Case VeldScheidTekens
      Case " "
         VeldTekenLijst.ListIndex = 1
      Case vbTab
         VeldTekenLijst.ListIndex = 2
      Case Else
         If Not VeldScheidTekens = vbNullString Then
            VeldTekenLijst.ListIndex = 0
            VeldTekenLijst.List(0) = VeldScheidTekens
         End If
   End Select
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure toont de bestanden in de geselecteerde map.
Private Sub MappenLijst_Change()
On Error GoTo Fout
Dim Keuze As Long

   BestandenLijst.Path = MappenLijst.Path
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure toont de mappen in het geselecteerde station.
Private Sub StationLijst_Change()
On Error GoTo Fout
Dim Keuze As Long

   MappenLijst.Path = StationLijst.Drive
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om de invoer vast te leggen.
Private Sub VeldTekenLijst_LostFocus()
On Error GoTo Fout
Dim Keuze As Long

   HaalInvoer
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procudure legt het geselecteerde veld type vast.
Private Sub VeldTypeLijst_Click(Index As Integer)
On Error GoTo Fout
Dim Keuze As Long

   VeldType(Index) = VeldTypeLijst(Index).ListIndex
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

