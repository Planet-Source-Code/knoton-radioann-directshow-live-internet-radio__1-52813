Attribute VB_Name = "modRadio"
Option Explicit
'**Declares for DirectShow audio
Private BasicAudio As IBasicAudio
Private MediaControl As IMediaControl

'**Open Radio with given adress
Public Function OpenRadio(RadioPath As String) As Boolean
On Local Error GoTo ErrHandler
Call CleanUp    'Clean up any previous action of Directshow

'Create a new Directshow object
Set MediaControl = New FilgraphManager

'Tell it to render the Radio channel adress
Call MediaControl.RenderFile(RadioPath)

'Create at audio object
Set BasicAudio = MediaControl
'Set audio properties
BasicAudio.Volume = 0
BasicAudio.Balance = 0

'Run the radiochannel
MediaControl.Run
OpenRadio = True

Exit Function
ErrHandler:
CleanUp
End Function

'Stop rendering the radio
Public Sub StopRadio()
If (ObjPtr(MediaControl) > 0) Then Call MediaControl.Stop
End Sub

'Deallocate the Directshow objects
Public Sub CleanUp()
On Local Error GoTo ErrHandler
If ObjPtr(MediaControl) > 0 Then MediaControl.Stop
If ObjPtr(BasicAudio) > 0 Then Set BasicAudio = Nothing
If ObjPtr(MediaControl) > 0 Then Set MediaControl = Nothing
    
Exit Sub
ErrHandler:
Err.Clear
End Sub
