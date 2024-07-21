VERSION 5.00
Begin VB.Form PrinterVenster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printer"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   ClipControls    =   0   'False
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   32.125
   Begin VB.CheckBox VoorraadVeld 
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.CheckBox KlantVeld 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.PictureBox KnoppenBalk 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   375
      Left            =   1440
      ScaleHeight     =   375
      ScaleWidth      =   2415
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   2415
      Begin VB.CommandButton AnnulerenKnop 
         Cancel          =   -1  'True
         Caption         =   "&Annuleren"
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton WijzigenKnop 
         Caption         =   "&Wijzigen"
         Default         =   -1  'True
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label VoorraadLabel 
      Caption         =   "Voorraad:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label KlantenLabel 
      Caption         =   "Klanten:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label DrukDezeVeldenAfLabel 
      Caption         =   "Druk deze velden af:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "PrinterVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het printer venster.
Option Explicit

'Deze procedure sluit dit venster.
Private Sub AnnulerenKnop_Click()
On Error GoTo Fout
Dim Keuze As Long

   Unload Me

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
Dim Afdrukken() As Boolean
Dim Keuze As Long
Dim Veld As Long

   AnnulerenKnop.ToolTipText = "Sluit dit venster."
   WijzigenKnop.ToolTipText = "Wijzigt de printer instellingen."

   Afdrukken() = HaalBits(Klanten.Afdrukken)
   For Veld = 0 To Klanten.AantalVelden - 1
      If Veld > 0 Then
         Load KlantVeld(Veld)
         KlantVeld(Veld).Top = (Veld * 1.5) + 2.5
         KlantVeld(Veld).Visible = True
         If KlantVeld(Veld).Top > Me.ScaleHeight - 5 Then
            Me.Height = Me.Height + 256
         End If
      End If
      KlantVeld(Veld).Caption = Klanten.VeldNaam(Veld)
      KlantVeld(Veld).TabIndex = KlantVeld(0).TabIndex + Veld
      KlantVeld(Veld).ToolTipText = "Geeft aan of dit veld word afgedrukt."
      KlantVeld(Veld).Value = -Afdrukken(Veld)
   Next Veld

   Afdrukken() = HaalBits(Voorraad.Afdrukken)
   For Veld = 0 To Voorraad.AantalVelden - 1
      If Veld > 0 Then
         Load VoorraadVeld(Veld)
         VoorraadVeld(Veld).Top = (Veld * 1.5) + 2.5
         VoorraadVeld(Veld).Visible = True
         If VoorraadVeld(Veld).Top > Me.ScaleHeight - 5 Then
            Me.Height = Me.Height + 256
         End If
      End If
      VoorraadVeld(Veld).Caption = Voorraad.VeldNaam(Veld)
      VoorraadVeld(Veld).TabIndex = VoorraadVeld(0).TabIndex + Veld
      VoorraadVeld(Veld).ToolTipText = "Geeft aan of dit veld word afgedrukt."
      VoorraadVeld(Veld).Value = -Afdrukken(Veld)
   Next Veld

   KnoppenBalk.Top = Me.ScaleHeight - 2
   Me.Left = (DepotbeheerderVenster.Width / 2) - (Me.Width / 2)
   Me.Top = (DepotbeheerderVenster.Height / 3) - (Me.Height / 2)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure legt de ingevoerde waardes vast.
Private Sub WijzigenKnop_Click()
On Error GoTo Fout
Dim Afdrukken(0 To 7) As Boolean
Dim Keuze As Long
Dim Veld As Long

   For Veld = 0 To Klanten.AantalVelden - 1
      Afdrukken(Veld) = -KlantVeld(Veld).Value
   Next Veld
   Klanten.Afdrukken = PlaatsBits(Afdrukken())

   For Veld = 0 To Voorraad.AantalVelden - 1
      Afdrukken(Veld) = -VoorraadVeld(Veld).Value
   Next Veld
   Voorraad.Afdrukken = PlaatsBits(Afdrukken())

   Unload Me

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

