VERSION 5.00
Begin VB.Form InvoerVenster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   38.125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox KnoppenBalk 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   375
      Left            =   2160
      ScaleHeight     =   375
      ScaleWidth      =   2415
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   2415
      Begin VB.CommandButton ActieKnop 
         Default         =   -1  'True
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton AnnulerenKnop 
         Cancel          =   -1  'True
         Caption         =   "&Annuleren"
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox InvoerVeld 
      Height          =   285
      Index           =   0
      Left            =   2280
      MaxLength       =   250
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label VeldNaam 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "InvoerVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het invoer venster.
Option Explicit

'Deze procedure legt de invoer vast.
Private Sub ActieKnop_Click()
On Error GoTo Fout
Dim Keuze As Long
Dim Veld As Long

   For Veld = 0 To DataBron.AantalVelden() - 1
      Buffer(Veld) = InvoerVeld(Veld).Text
   Next Veld
 
   Unload Me
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure breekt de invoer af en sluit dit venster.
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

'Deze procedure stelt dit venster in.
Private Sub Form_Load()
On Error GoTo Fout
Dim Keuze As Long
Dim Veld As Long

   AnnulerenKnop.ToolTipText = "Sluit dit venster."

   Me.Height = (DataBron.AantalVelden() * 512) + 64
   KnoppenBalk.Top = (DataBron.AantalVelden() * 1.5) + 0.5
   
   For Veld = 0 To DataBron.AantalVelden() - 1
      If Veld > 0 Then
         Load InvoerVeld(Veld)
         Load VeldNaam(Veld)
      End If
      InvoerVeld(Veld).Text = Buffer(Veld)
      InvoerVeld(Veld).TabIndex = Veld
      InvoerVeld(Veld).Top = (Veld * 1.5) + 0.5
      InvoerVeld(Veld).Visible = True
      VeldNaam(Veld) = DataBron.VeldNaam(Veld)
      VeldNaam(Veld).ToolTipText = "De inhoud van dit veld."
      VeldNaam(Veld).Top = (Veld * 1.5) + 0.5
      VeldNaam(Veld).Visible = True
   Next Veld
   
   If ActieIsToeVoegen Then
      ActieKnop.Caption = "&Toevoegen"
      ActieKnop.ToolTipText = "Voegt dit item toe."
   Else
      ActieKnop.Caption = "&Wijzigen"
      ActieKnop.ToolTipText = "Wijzigt dit item."
   End If
   
EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

'Deze procedure selecteert automatisch de inhoud van het veld.
Private Sub InvoerVeld_GotFocus(Index As Integer)
On Error GoTo Fout
Dim Keuze As Long

   InvoerVeld(Index).SelStart = 0
   InvoerVeld(Index).SelLength = Len(InvoerVeld(Index).Text)

EindeRoutine:
   Exit Sub

Fout:
   Keuze = HandelFoutAf()
   If Keuze = vbIgnore Then Resume EindeRoutine
   If Keuze = vbRetry Then Resume
End Sub

