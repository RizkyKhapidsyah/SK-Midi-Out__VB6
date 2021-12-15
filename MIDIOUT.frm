VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "MIDI OUT"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAutoSend 
      Caption         =   "AutoSend"
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdALLOFF 
      Caption         =   "Reset All"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtChn 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   10
      Text            =   "1"
      Top             =   1080
      Width           =   375
   End
   Begin VB.ComboBox cmbEvent 
      Height          =   315
      Left            =   2760
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmbSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtVal2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3960
      TabIndex        =   4
      Text            =   "100"
      Top             =   1200
      Width           =   375
   End
   Begin VB.ListBox lstMidiOut 
      Height          =   1620
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtVal1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3960
      TabIndex        =   1
      Text            =   "60"
      Top             =   840
      Width           =   375
   End
   Begin ComctlLib.Slider sldValue 
      Height          =   2175
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   3836
      _Version        =   327682
      Orientation     =   1
      Max             =   127
      SelStart        =   20
      TickStyle       =   3
      TickFrequency   =   0
      Value           =   20
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "click text box to change"
      Height          =   555
      Left            =   3840
      TabIndex        =   12
      Top             =   1560
      Width           =   765
   End
   Begin VB.Label lblChn 
      Alignment       =   2  'Center
      Caption         =   "Chn"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblEventType 
      AutoSize        =   -1  'True
      Caption         =   "Event Type"
      Height          =   195
      Left            =   3000
      TabIndex        =   9
      Top             =   240
      Width           =   825
   End
   Begin VB.Label lblVal2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Val 2"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3480
      TabIndex        =   6
      Top             =   1200
      Width           =   360
   End
   Begin VB.Label lblVal1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Val 1 >"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3405
      TabIndex        =   5
      Top             =   840
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select MIDI Out Port"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1470
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EventParm As Byte
Const ColorSelect = 1
Const ColorNoSelect = 1

Private Sub cmbEvent_Click()
Dim x As Byte
For x = 0 To 3
If cmbEvent.ListIndex = x Then txtVal2.Enabled = True
Next x
For x = 4 To 7
If cmbEvent.ListIndex = x Then txtVal2.Enabled = False
Next x
End Sub

Private Sub cmbSend_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call midioutmsg(Me.cmbEvent.ItemData(cmbEvent.ListIndex), Val(Me.txtChn - 1), Val(Me.txtVal1), Val(Me.txtVal2))
End Sub

Private Sub cmbSend_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If cmbEvent.ListIndex = 1 Then _
Call midioutmsg(Me.cmbEvent.ItemData(0), Me.txtChn - 1, Me.txtVal1, Me.txtVal2)
End Sub

Private Sub Form_Load()
Call mdlGeneral.midi_ListOutdevs(Me.lstMidiOut)
Me.cmbEvent.AddItem "Note Off"
Me.cmbEvent.ItemData(cmbEvent.NewIndex) = &H80
Me.cmbEvent.AddItem "Note On"
Me.cmbEvent.ItemData(cmbEvent.NewIndex) = &H90
Me.cmbEvent.AddItem "Aftertouch"
Me.cmbEvent.ItemData(cmbEvent.NewIndex) = &HA0
Me.cmbEvent.AddItem "Control Change"
Me.cmbEvent.ItemData(cmbEvent.NewIndex) = &HB0
Me.cmbEvent.AddItem "Program Change"
Me.cmbEvent.ItemData(cmbEvent.NewIndex) = &HC0
Me.cmbEvent.AddItem "Chan Pressure"
Me.cmbEvent.ItemData(cmbEvent.NewIndex) = &HD0
Me.cmbEvent.AddItem "Pitch Bend"
Me.cmbEvent.ItemData(cmbEvent.NewIndex) = &HE0
Me.cmbEvent.ListIndex = 0
Me.lstMidiOut.ListIndex = 0
EventParm = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call mdlGeneral.midi_outStatus(lstMidiOut.ListIndex - 1, False) ' Close Selected Port
End Sub

Private Sub cmdALLOFF_Click()
Call midioutmsg(Me.cmbEvent.ItemData(3), txtChn - 1, &H78, &H7F)
End Sub

Private Sub lstMidiOut_Click()
Call mdlGeneral.midi_outStatus(lstMidiOut.ListIndex - 1, True) ' Open Selected Port
End Sub

Private Sub sldValue_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If chkAutoSend = 1 And Button = 1 Then _
Call midioutmsg(Me.cmbEvent.ItemData(cmbEvent.ListIndex), txtChn - 1, txtVal1, txtVal2)

End Sub

Private Sub sldValue_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If chkAutoSend = 1 Then Call midioutmsg(Me.cmbEvent.ItemData(3), txtChn - 1, &H7B, &H7F)
If chkAutoSend = 1 Then Call midioutmsg(Me.cmbEvent.ItemData(3), txtChn - 1, &H78, &H7F)
If chkAutoSend = 1 Then Call midioutmsg(Me.cmbEvent.ItemData(3), txtChn - 1, &H79, &H7F)
End Sub

Private Sub sldValue_Scroll()
If EventParm = 1 Then txtVal1 = 127 - sldValue
If EventParm = 2 Then txtVal2 = 127 - sldValue
End Sub

Private Sub txtChn_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
If Val(txtChn) > 16 Then txtChn = 16
If Val(txtChn) < 1 Then txtChn = 1
End If
End Sub

Private Sub txtVal1_GotFocus()
EventParm = 1
lblVal1.Caption = "Val 1 >"
lblVal2.Caption = "Val 1"
End Sub

Private Sub txtVal2_GotFocus()
EventParm = 2
lblVal2.Caption = "Val 1 >"
lblVal1.Caption = "Val 1"
End Sub

