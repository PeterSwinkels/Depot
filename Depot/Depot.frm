VERSION 5.00
Begin VB.MDIForm DepotBeheerderVenster 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Depot Beheerder"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "Depot.frx":0000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu UitloggenMenu 
      Caption         =   "&Uitloggen"
   End
   Begin VB.Menu ZoekenMenu 
      Caption         =   "&Zoeken"
   End
   Begin VB.Menu HandleidingMenu 
      Caption         =   "&Handleiding"
   End
   Begin VB.Menu BackupMenu 
      Caption         =   "&Backup"
   End
   Begin VB.Menu InformatieMenu 
      Caption         =   "&Informatie"
   End
End
Attribute VB_Name = "DepotBeheerderVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het hoofdvenster van dit programma.
Option Explicit

'Deze procedure opent het backup venster.
Private Sub BackupMenu_Click()
On Error GoTo Fout
Dim Keuze As Long

   If Beheerder() Then
      BackupVenster.Show
      BackupVenster.ZOrder
   End If
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure opent de handleiding.
Private Sub HandleidingMenu_Click()
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim Keuze As Long
Dim Pad As String

   ActieveMap = Pad
   ActiefBestand = "Handleiding.hta"
   If Dir$(Pad & "Handleiding.hta") = vbNullString Then
      MsgBox "De handleiding kan niet worden gevonden.", vbExclamation
   Else
      Pad = VoegScheidingstekenToe(App.Path)
      Shell "Mshta.exe """ & Pad & "Handleiding.hta""", vbNormalFocus
   End If
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om de programma informatie te tonen.
Private Sub InformatieMenu_Click()
On Error GoTo Fout
Dim Keuze As Long

   ToonProgrammaInformatie
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stelt dit venster in.
Private Sub MDIForm_Initialize()
On Error GoTo Fout
Dim Keuze As Long

   Me.WindowState = vbMaximized
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure stelt dit venster in.
Private Sub MDIForm_Load()
On Error GoTo Fout
Dim Keuze As Long

   Me.Width = Screen.Width / 1.5
   Me.Height = Screen.Height / 1.5
   
   MenuVenster.Show
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om de gegevens te bewaren wanneer dit venster wordt gesloten.
Private Sub MDIForm_Unload(Cancel As Integer)
On Error GoTo Fout
Dim Keuze As Long

   SlaGegevensOp
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure logt de gebruiker of beheerder uit.
Private Sub UitloggenMenu_Click()
On Error GoTo Fout
Dim Keuze As Long

   If Wachtwoord = vbNullString Then
      MsgBox "Geen wachtwoord ingesteld.", vbExclamation
   Else
      VraagWachtwoord
   End If
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure toont het zoek venster.
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

