VERSION 5.00
Begin VB.Form frmRadio 
   Caption         =   "Radio Ann"
   ClientHeight    =   0
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2160
   Icon            =   "frmRadio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   0
   ScaleWidth      =   2160
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuRadio 
      Caption         =   "Radio"
      Begin VB.Menu mnuChannel 
         Caption         =   "Channel"
         Index           =   0
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close Radio"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmRadio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**Type that holds Radiochannels
'**The name given to it and its adress.
Private Type Channels
    Name As String
    Adress As String
End Type

'**Type Array of all radiochannels
Private Channel() As Channels

'**Function that retreives all Radiochannels in the ini file
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Function that returns alla radiochannels Chr(0) delimited in the ini file
Public Function ReadIniSection(Filename As String, Section As String) As String
Dim RetVal As String * 4096, v As Long
    v = GetPrivateProfileSection(Section, RetVal, 4096, Filename)
    ReadIniSection = Left(RetVal, v - 1)
End Function


Private Sub Form_Load()
FixMenu 'load the menu
AddSystray Me, "Radio Ann" 'Put it to the systray
Me.Hide 'Hide the form
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
CleanUp 'Dont use DirectShow any more
RemoveSystray 'Tell the system to remove me from the systray
End Sub

'**Check if the mouse clicks the icon in the systray
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim rtn As Long
'Get the message
If Me.ScaleMode = vbPixels Then
    rtn = X
  Else
    rtn = X / Screen.TwipsPerPixelX
End If

'If it is the right message popup the menu
Select Case rtn
  Case WM_LBUTTONDOWN
     Me.PopupMenu Me.mnuRadio
  Case WM_RBUTTONDOWN
     Me.PopupMenu Me.mnuRadio
End Select

End Sub

'**Loads the menu
Private Sub FixMenu()
Dim i As Integer, var As Variant, var2 As Variant, s As String
'Get all channels
s = ReadIniSection(App.Path & "\RadioChannels.ini", "Channels")
ReDim Channel(0)

'Make an array with all channels
var = Split(s, Chr(0))
For i = 0 To UBound(var)
    If i > 0 Then Load mnuChannel(i) 'Load a new menu item
    ReDim Preserve Channel(i) 'Allocate more space to the channel array
    var2 = Split(var(i), "=") 'Separate the name of the radiochannel and its adress
    
    Channel(i).Name = var2(0)   'Add the name to the channel type array
    Channel(i).Adress = var2(1) 'Add the adress to the channel type array
    
    mnuChannel(i).Caption = Channel(i).Name 'put the name of the channel to the menu item
    mnuChannel(i).Visible = True 'Show it
Next
End Sub

'Change Channel or close
Private Sub mnuChannel_Click(Index As Integer)
Dim Namn As String, kanal As String
Namn = Channel(Index).Name
kanal = Channel(Index).Adress
'Open the channel and change the mouse tooltip to its name to be seen when it is over the icon in the systray
If OpenRadio(kanal) Then ModifySystray Me, Namn
End Sub

'Close The radiochanel
Private Sub mnuClose_Click()
ModifySystray Me, "Radio Ann"
StopRadio
End Sub

'Close the application
Private Sub mnuExit_Click()
Unload Me
End Sub
