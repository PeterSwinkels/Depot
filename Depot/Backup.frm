VERSION 5.00
Begin VB.Form BackupVenster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup"
   ClientHeight    =   1215
   ClientLeft      =   2970
   ClientTop       =   1845
   ClientWidth     =   3975
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1215
   ScaleWidth      =   3975
   Begin VB.CommandButton ZetBackupTerugKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Zet Backup Terug"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox BackupmapVeld 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton MaakBackupKnop 
      Appearance      =   0  'Flat
      Caption         =   "&Maak Backup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label BackuplocatieLabel 
      Caption         =   "Backuplocatie:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "BackupVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het backup venster.
Option Explicit

'Deze procedure kopieert bestanden van de opgegeven bron map naar de opgegeven doel map.
Private Sub KopieerBestanden(Bronmap As String, Doelmap As String, ByRef Afbreken As Boolean)
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim Bestand As String
Dim Keuze As Long

   If BackupmapVeld.Text = vbNullString Then
      MsgBox "Geen backup map opgegeven", vbInformation
   Else
      Bronmap = VoegScheidingstekenToe(Bronmap)
      Doelmap = VoegScheidingstekenToe(Doelmap)

      Afbreken = False

      If Dir$(Doelmap, vbArchive Or vbDirectory Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = vbNullString Then
         ActieveMap = Doelmap
         MkDir Doelmap
      Else
         Afbreken = (MsgBox("Alle bestanden in " & Doelmap & vbCr & "zullen verwijderd worden. Doorgaan?", vbExclamation Or vbYesNo Or vbDefaultButton2) = vbNo)
      End If

      If Not Afbreken Then
         Screen.MousePointer = vbHourglass

         VerwijderBestanden Doelmap

         ActieveMap = Bronmap
         Bestand = Dir$(Bronmap & "*.*", vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
         Do Until Bestand = vbNullString
            ActiefBestand = Bestand

            ActieveMap = Bronmap
            FileCopy Bronmap & Bestand, Doelmap & Bestand

            ActieveMap = Doelmap
            SetAttr Doelmap & Bestand, (GetAttr(Bronmap & Bestand) And (vbAlias Or vbArchive Or vbDirectory Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem Or vbVolume))

            Bestand = Dir$()
            If UCase$(Bestand) = "BACKUP.DAT" Then Bestand = Dir$()
         Loop
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

'Deze procedure maakt een backup label bestand in de opgegeven doel map.
Private Sub MaakBackupLabel(Doelmap As String)
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim BestandH As Integer
Dim Keuze As Long

   If Not Doelmap = vbNullString Then
      Doelmap = VoegScheidingstekenToe(Doelmap)
      ActieveMap = Doelmap
      ActiefBestand = "Backup.dat"
      BestandH = FreeFile()
      Open Doelmap & "Backup.dat" For Output Lock Read Write As BestandH
         Print #BestandH, BACKUP_LABEL;
      Close BestandH
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure controleert of er een backup aanwezig is in de opgegeven map.
Private Function MapIsBackup(Bronmap As String) As Boolean
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim Data As String
Dim BestandH As Integer
Dim IsBackup As Boolean
Dim Keuze As Long

   IsBackup = False
   If Not Bronmap = vbNullString Then
      Bronmap = VoegScheidingstekenToe(Bronmap)
      ActieveMap = Bronmap
      ActiefBestand = "Backup.dat"
      BestandH = FreeFile()
      If Not Dir$(Bronmap & "Backup.dat", vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = vbNullString Then
         Open Bronmap & "Backup.dat" For Binary Lock Read Write As BestandH
            Data = Input$(LOF(BestandH), BestandH)
            IsBackup = (Data = BACKUP_LABEL)
         Close BestandH
      End If
   End If

EindeRoutine:
   MapIsBackup = IsBackup
   Exit Function

Fout:
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Function

'Deze procedure verwijdert de bestanden in de opgegeven map.
Private Sub VerwijderBestanden(Map As String)
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim Bestand As String
Dim Keuze As Long

   ActieveMap = Map
   Bestand = Dir$(Map & "*.*", vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
   Do Until Bestand = vbNullString
      ActiefBestand = Bestand
      SetAttr Map & Bestand, vbNormal
      Kill Map & Bestand
      Bestand = Dir$()
   Loop

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub BackupMapVeld_GotFocus()
On Error GoTo Fout
Dim Keuze As Long

   BackupmapVeld.SelStart = 0
   BackupmapVeld.SelLength = Len(BackupmapVeld.Text)

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

   BackupmapVeld.ToolTipText = "De map waar de backup wordt geplaatst."
   MaakBackupKnop.ToolTipText = "Plaatst de backup in de opgegeven map."
   ZetBackupTerugKnop.ToolTipText = "Vervangt de huidige de gegevens met de backup in de opgegeven map."

   Me.Left = (DepotBeheerderVenster.Width / 2) - (Me.Width / 2)
   Me.Top = (DepotBeheerderVenster.Height / 3) - (Me.Height / 2)

   BackupmapVeld.Text = BackupMap

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om een backup te maken.
Private Sub MaakBackupKnop_Click()
On Error GoTo Fout
Dim Afbreken As Boolean
Dim Keuze As Long

   If Beheerder() Then
      BackupMap = BackupmapVeld.Text

      SlaGegevensOp

      KopieerBestanden ".\Data\", BackupMap, Afbreken
      If Not Afbreken Then
         MaakBackupLabel BackupMap
         MsgBox "De backup is gemaakt.", vbInformation
      End If
   End If

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure geeft de opdracht om een backup terug te zetten.
Private Sub ZetBackupTerugKnop_Click()
On Error GoTo Fout
Dim ActiefBestand As String
Dim ActieveMap As String
Dim Afbreken As Boolean
Dim Keuze As Long

   If Beheerder() Then
      If MapIsBackup(BackupmapVeld.Text) Then
         KopieerBestanden BackupmapVeld.Text, ".\Data\", Afbreken

         MsgBox "De Depot Beheerder wordt nu opnieuw gestart.", vbInformation
         ActieveMap = CurDir$()
         ActiefBestand = "Depot.exe"
         Shell "Depot.exe", vbNormalFocus
      Else
         MsgBox "Geen backup aanwezig.", vbExclamation
      End If
   End If

EindeRoutine:
   GegevensOpslaan = False
   Unload DepotBeheerderVenster
   Exit Sub

Fout:
   Keuze = HandelFoutAf(ActieveMap, ActiefBestand)
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

