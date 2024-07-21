VERSION 5.00
Begin VB.UserControl LijstObject 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   ScaleHeight     =   15
   ScaleMode       =   0  'User
   ScaleWidth      =   40
   Begin VB.PictureBox UitvoerVenster 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawMode        =   6  'Mask Pen Not
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   15.063
      ScaleMode       =   4  'Character
      ScaleWidth      =   38.125
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.VScrollBar SchuifBalk 
      Height          =   3615
      Left            =   4560
      Max             =   0
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "LijstObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Deze module bevat het lijstobject.
Option Explicit
Private AantalZichtbaar As Long        'Bevat het aantal zichtbare items in de lijst.
Private BovensteItemV As Long          'Bevat het bovenste zichtbare item in de lijst.
Private DatabronV As Object            'Bevat de bron van de gegevens in de lijst.
Private DatabronIngesteld As Boolean   'Geeft aan of de bron van de gegevens is geselecteerd.
Private SelectieV As Long              'Bevat het geselecteerde item in de lijst.

'Deze procedure legt het bovenste item in de lijst vast.
Public Property Let BovensteItem(NieuwBovensteItem As Long)
On Error GoTo Fout
Dim Keuze As Long

   SchuifBalk.Value = NieuwBovensteItem

EindeRoutine:
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven databron vast.
Public Property Set Databron(NieuweDataBron As Object)
On Error GoTo Fout
Dim Keuze As Long

   Set DatabronV = NieuweDataBron
   DatabronIngesteld = True

EindeRoutine:
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure stuurt het geselecteerde item terug.
Public Property Get Selectie() As Long
On Error GoTo Fout
Dim Keuze As Long

EindeRoutine:
   Selectie = SelectieV
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure legt de opgegeven selectie vast.
Public Property Let Selectie(NieuweSelectie As Long)
On Error GoTo Fout
Dim Keuze As Long

   SelectieV = NieuweSelectie
   WerkLijstBij

EindeRoutine:
   Exit Property

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Property

'Deze procedure toont de gegevens in de lijst.
Public Sub WerkLijstBij()
On Error GoTo Fout
Dim Item As Long
Dim Data As String
Dim DataBreedte As Single
Dim Keuze As Long
Dim Percentage As Single
Dim Veld As Long
Dim VeldBreedte As Single
Dim VeldX As Single
Dim VeldY As Single

   If DatabronIngesteld Then
      If SelectieV >= DatabronV.AantalItems() Then SelectieV = DatabronV.AantalItems() - 1
      If SelectieV < 0 And DatabronV.AantalItems() > 0 Then SelectieV = 0
      SchuifBalk.Max = DatabronV.AantalItems()

      Percentage = 100 / DatabronV.TotaleVeldBreedte()

      Screen.MousePointer = vbHourglass
      UitvoerVenster.Cls
      UitvoerVenster.Font.Bold = True
      For Veld = 0 To DatabronV.AantalVelden() - 1
         VeldX = UitvoerVenster.CurrentX
         Data = DatabronV.VeldNaam(Veld)
         VeldBreedte = BerekenProcent(Percentage * DatabronV.VeldBreedte(Veld), UitvoerVenster.ScaleWidth)
         DataBreedte = VeldBreedte * (Len(Data) / UitvoerVenster.TextWidth(Data))
         If DatabronV.VeldRechtsUitlijnen(Veld) Then
            Data = Right$(Data, DataBreedte)
            UitvoerVenster.CurrentX = (UitvoerVenster.CurrentX + VeldBreedte - 1) - UitvoerVenster.TextWidth(Data)
         Else
            Data = Left$(Data, DataBreedte)
         End If
         UitvoerVenster.Print Data;
         UitvoerVenster.CurrentX = VeldX + VeldBreedte
      Next Veld

      UitvoerVenster.CurrentY = 1
      UitvoerVenster.Font.Bold = False
      For Item = BovensteItemV To BovensteItemV + AantalZichtbaar
         If Item = DatabronV.AantalItems() Then Exit For
         UitvoerVenster.CurrentX = 0
         VeldY = UitvoerVenster.CurrentY
         For Veld = 0 To DatabronV.AantalVelden() - 1
            VeldX = UitvoerVenster.CurrentX
            Data = DatabronV.Data(Veld, Item)
            VeldBreedte = BerekenProcent(Percentage * DatabronV.VeldBreedte(Veld), UitvoerVenster.ScaleWidth)
            If DatabronV.VeldRechtsUitlijnen(Veld) Then
               Data = Right$(Data, DataBreedte)
               UitvoerVenster.CurrentX = (UitvoerVenster.CurrentX + VeldBreedte - 1) - UitvoerVenster.TextWidth(Data)
            Else
               Data = Left$(Data, DataBreedte)
            End If
            UitvoerVenster.Print Data;
            UitvoerVenster.CurrentX = VeldX + VeldBreedte
         Next Veld
         If Item = SelectieV Then UitvoerVenster.Line (0, UitvoerVenster.CurrentY)-(UitvoerVenster.ScaleWidth, UitvoerVenster.CurrentY + 1), , BF
         UitvoerVenster.CurrentY = VeldY + 1
      Next Item
   End If

EindeRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure past de lijst aan wanneer de schuifbalkknop wordt verplaatst.
Private Sub SchuifBalk_Change()
On Error GoTo Fout
Dim Keuze As Long

   BovensteItemV = SchuifBalk.Value
   WerkLijstBij

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure past de selectie aan wanneer de gebruiker door de lijst scrolt met de navigatie toetsen.
Private Sub UitvoerVenster_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Fout
Dim Keuze As Long

   Select Case KeyCode
      Case vbKeyUp
         If SelectieV > 0 Then SelectieV = SelectieV - 1
      Case vbKeyDown
         If SelectieV < DatabronV.AantalItems() - 1 Then SelectieV = SelectieV + 1
   End Select

   WerkLijstBij

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure past de selectie aan wanneer de gebruiker met de muis in de lijst klikt.
Private Sub UitvoerVenster_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Fout
Dim Keuze As Long
Dim NieuweSelectie As Long

   If DatabronIngesteld Then
      NieuweSelectie = (CLng(y - 1) + BovensteItemV)
      If NieuweSelectie <= DatabronV.AantalItems() Then
         SelectieV = NieuweSelectie
         WerkLijstBij
      End If
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stelt dit object in.
Private Sub UserControl_Initialize()
On Error GoTo Fout
Dim Keuze As Long

   AantalZichtbaar = 0
   BovensteItemV = 0
   Set DatabronV = Nothing
   DatabronIngesteld = False
   SelectieV = 0

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure past de objecten in dit object aan de nieuwe afmetingen aan.
Private Sub UserControl_Resize()
On Error Resume Next

   AantalZichtbaar = Int(ScaleHeight) - 2
   SchuifBalk.Left = ScaleWidth - 2
   SchuifBalk.Height = ScaleHeight
   UitvoerVenster.Width = ScaleWidth - 2
   UitvoerVenster.Height = ScaleHeight
End Sub

