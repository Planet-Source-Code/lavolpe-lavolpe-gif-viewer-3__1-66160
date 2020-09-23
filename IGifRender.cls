VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IGifRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' This Interface must be used if the cGIFViewer class will create and use its own timer
' for displaying animated GIFs. The timer will call an event each time a frame is to
' be drawn so that your project can pass the class the target/destination DC to accept
' the drawn image. Other optional flags are also provided.

' To include the interface in your project, simply add the following statements in
' your form/class/module Declarations section:  Implements IGifRender

Public Enum RenderMessage
    msgRendered = 0
    msgProgress = 1
    msgLoopsEnded = 2
End Enum

Public Sub GetRenderDC(ByVal ViewerID As Long, ByVal FrameIndex As Long, _
                    ByRef destDC As Long, Optional ByRef hwndRefresh As Long, _
                    Optional ByRef bAutoRedraw As Boolean = False, _
                    Optional ByRef PostNotify As Boolean = False)

' ViewerID :: the ViewerID parameter the user passed to cGIFViewer.LoadGif
' FrameIndex :: the animation frame index to be rendered. Frames are 1-bound, not 0-bound
' destDC :: the destination DC to render the frame
' hwndRefresh :: this allows the class to refresh your window for you
' PostNotify :: Calls an IGifRender.Rendered event after the frame has been rendered
'               passing the msgRendered message.
'             ' This is useful if the destDC is a memory DC to inform you rendering is done.
End Sub

Public Sub Rendered(ByVal ViewerID As Long, ByVal FrameIndex As Long, _
                    ByVal Message As RenderMessage, ByVal MsgValue As Long)

' ViewerID :: the ViewerID parameter the user passed to cGIFViewer.LoadGif
' FrameIndex :: the frame index just rendered; Frames are 1-bound, not 0-bound
' Message:: one of the three RenderMessages
'   msgRendered - frame was drawn, msgValue is the DC frame was rendered to
'   msgProgress - frame is loading, msgValue is the Percent between 0 to 100
'       Note: when msgValue=0 then cGifViewer.SetAnimationBkg can be called
'       enabling the 1st frame to immediately be displayed even though the
'       remaining frames (if any) have not yet been processed
'   msgLoopsEnded - frame loops expired, msgValue is the number of loops that were completed
' MsgValue:: as just stated above

' if this event is not used, you must at least view it in your VB form/uc/class
' otherwise you'll get an error when you first try to run your project

' This event will inform you, during automation only, when specific rendering events
' occur. By monitoring this event, you can continue whatever action is needed.
' Let's say the frame is being rendered to a memory DC, then this event will inform
'   you that the rendering was completed.

End Sub
