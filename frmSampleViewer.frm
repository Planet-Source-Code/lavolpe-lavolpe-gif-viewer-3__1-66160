VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSampleViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample Multi-GIF Viewer"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkNoPopup 
      Caption         =   "Prevent Popup WIndow"
      Height          =   240
      Left            =   3225
      TabIndex        =   13
      Top             =   4830
      Width           =   2925
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Toggle BackColor"
      Height          =   420
      Left            =   315
      TabIndex        =   4
      Top             =   4365
      Width           =   2250
   End
   Begin VB.PictureBox picBkg 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   135
      Picture         =   "frmSampleViewer.frx":0000
      ScaleHeight     =   1095
      ScaleWidth      =   1305
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   210
      Width           =   1305
      Begin VB.Label Label1 
         Caption         =   "Render to PicBox Example"
         Height          =   615
         Index           =   1
         Left            =   90
         TabIndex        =   11
         Top             =   270
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Shape shpMarker 
         Height          =   1095
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   1305
      End
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2625
      Top             =   4335
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboStretchMode 
      Height          =   315
      ItemData        =   "frmSampleViewer.frx":43C4
      Left            =   3195
      List            =   "frmSampleViewer.frx":43D4
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4005
      Width           =   2895
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   ">"
      Height          =   405
      Index           =   1
      Left            =   2100
      TabIndex        =   3
      ToolTipText     =   "Play/Animate"
      Top             =   3930
      Width           =   465
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   ">>"
      Height          =   405
      Index           =   3
      Left            =   1500
      TabIndex        =   2
      ToolTipText     =   "Step"
      Top             =   3930
      Width           =   465
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "| |"
      Height          =   405
      Index           =   2
      Left            =   900
      TabIndex        =   1
      ToolTipText     =   "Pause"
      Top             =   3930
      Width           =   465
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "O"
      Height          =   405
      Index           =   0
      Left            =   330
      TabIndex        =   0
      ToolTipText     =   "Stop"
      Top             =   3930
      Width           =   465
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1275
      Left            =   4605
      TabIndex        =   8
      Top             =   2595
      Width           =   1605
      Begin VB.Image imgFrame 
         Height          =   375
         Left            =   390
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Render to Frame Example"
         Height          =   465
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   345
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Shape shpMarker 
         Height          =   1005
         Index           =   11
         Left            =   105
         Top             =   210
         Width           =   1425
      End
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Clear ALL"
      Enabled         =   0   'False
      Height          =   420
      Index           =   1
      Left            =   4905
      TabIndex        =   7
      Top             =   4365
      Width           =   1230
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load a GIF File"
      Height          =   420
      Index           =   0
      Left            =   3165
      TabIndex        =   6
      Top             =   4365
      Width           =   1725
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   $"frmSampleViewer.frx":4419
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   2
      Left            =   510
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Shape shpMarker 
      Height          =   1095
      Index           =   10
      Left            =   3150
      Top             =   2730
      Width           =   1305
   End
   Begin VB.Shape shpMarker 
      Height          =   1095
      Index           =   9
      Left            =   1635
      Top             =   2730
      Width           =   1305
   End
   Begin VB.Shape shpMarker 
      Height          =   1095
      Index           =   8
      Left            =   120
      Top             =   2730
      Width           =   1305
   End
   Begin VB.Shape shpMarker 
      Height          =   1095
      Index           =   7
      Left            =   4665
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Shape shpMarker 
      Height          =   1095
      Index           =   6
      Left            =   3135
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Shape shpMarker 
      Height          =   1095
      Index           =   5
      Left            =   1635
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Shape shpMarker 
      Height          =   1095
      Index           =   4
      Left            =   120
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Shape shpMarker 
      Height          =   1095
      Index           =   3
      Left            =   4680
      Top             =   210
      Width           =   1305
   End
   Begin VB.Shape shpMarker 
      Height          =   1095
      Index           =   2
      Left            =   3150
      Top             =   210
      Width           =   1305
   End
   Begin VB.Shape shpMarker 
      Height          =   1095
      Index           =   1
      Left            =   1650
      Top             =   210
      Width           =   1305
   End
End
Attribute VB_Name = "frmSampleViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Example of viewing multiple GIFs.
'-----------------------------------------
' For any GIF, ensure 3 things are done, in this order:
' 1. Load the GIF (cGifViewer.LoadGIF)
' 2. Set the background properites (cGifViewer.SetAnimationBkg)
' 3. Start playing the GIF (cGifViewer.AnimationState = gfaPlaying)

' This form was a test for my viewer class. Previous viewer classes either used too
' many resources by saving each animated frame as a stdPicture or were way to slow
' by saving each frame as a separate file. My second version of the viewer used one
' stdPicture created on the fly as each frame was about to be rendered. This worked
' well and was very fast until you loaded more than 6 or so in a project. Then the
' constant creation of dozens or even hundreds of stdPictures every second, slowed
' down the project to a crawl, even when compiled; way too CPU intensive

' That brings me to my third version; the one you downloaded as cGIFViewer. It uses
' a unique method where it stores the entire GIF as a single bitmap much like the
' film strip of a movie. Not only is this very GDI heap friendly (2 GDI objects:
' 1 dc/1 bmp per GIF), it is also extremely fast to render each frame. Now I can
' display dozens of GIFs on a form without bringing the project to a crawl. I
' have not stress tested it though; my max test was using 50 animated, transparent
' GIFs that ranged from 4 to 40 frames each -- no noticable slow down

' The IGifRender interface allows the GIF class to be used with "With Events"-like
' events. Each GIF you load, you provide a reference for it. That reference is passed
' in the 2 IGifRender events so you know which GIF is being referenced/processed.

' SPECIAL NOTES & PROBLEMS FOR CONTROLS WITHOUT A .HDC PROPERTY

' Notice that one of the samples is drawn on a VB Frame. VB Frame's have no .HDC
' property, but the GIF class will accept an .HWND (See IGifRender.Render event). The
' class will then get the hDC from that hWnd. Of course, since VB Frames don't have
' an AutoRedraw property, anything drawn on the Frame's DC will disappear when the
' frame is covered/uncovered by other windows or frame is refreshed. In cases like
' these and only if your Form has its AutoRedraw set to False, simply referesh the
' GIFs when the form's Paint event fires. But if AutoRedraw is True, you won't know
' if the VB Frame was erased because you don't get the Paint event. In these cases,
' when the GIF is animating, do nothing as the next frame will be drawn as expected.
' But if the frame is single-framed or the GIF has ended its loops, then ask the class
' to return the GIF frame as a stdPicture, for example. You can then assign the GIF
' frame to an Image Control that you placed in your VB Frame/container and it will always
' be visible or you can change the GIF's loop count to infinite so it is always animating,
' or use a picturebox in the frame with the picturebox.AutoRedraw=True. However, not
' using VB Frames for hosting animated GIFs may probably be the best solution. One of the
' suggested solutions is provided as the MakeFrameImagePermanent routine

Private myGifViewer() As cGIFViewer ' array of gif viewers
Implements IGifRender               ' required when using the gif viewer

Private Sub cmdAction_Click(Index As Integer)
    '0=stop; 1=start; 2=pause; 3=step; 4=refresh
    Dim g As Long
    If myGifViewer(0) Is Nothing Then Exit Sub      ' no viewers loaded yet
    For g = shpMarker.LBound To shpMarker.UBound
        If Not myGifViewer(g) Is Nothing Then
            myGifViewer(g).AnimationState = Index
        End If
    Next
    ' Special handling for VB Frames,see IGifRender_Rendered & following routine for more
    Call MakeFrameImagePermanent(shpMarker.UBound)
    
End Sub

Private Sub cmdLoad_Click(Index As Integer)

    ' Load a new viewer or clear all viewers
    Dim Slot As Long
    If Index = 0 Then
        
        Dim sFileName As String
        
        On Error GoTo EH
        dlgFile.ShowOpen
        DoEvents
        
        For Slot = shpMarker.LBound To shpMarker.UBound
            ' find a free viewer to use
            If myGifViewer(Slot) Is Nothing Then Exit For
        Next
        
        Set myGifViewer(Slot) = New cGIFViewer
        If myGifViewer(Slot).LoadGIF(dlgFile.FileName, Slot, Me, Me.hWnd) < 1 Then
        
            myGifViewer(Slot).UnloadGIF
            Set myGifViewer(Slot) = Nothing
            MsgBox "Failed to load that image.", vbExclamation + vbOKOnly
            
        Else
            'TIP. If wanting to display the first image immediately, you must
            ' call SetAnimationBkg in the IGifRender_Rendered event, otherwise
            ' you would call it here, after the successful loading of the GIF
             ' ... myGifViewer(Slot).SetAnimationBkg gfdBkgFromDC, Me.hDC, ....
        
            ' now we start animating the viewer
            myGifViewer(Slot).AnimationState = gfaPlaying
            
            ' prevent loading any more gifs if all viewers are used
            If Slot = shpMarker.UBound Then cmdLoad(0).Enabled = False
            cmdLoad(1).Enabled = True   ' allow removing all viewers
            
            ' show the most current file in a fuller size
            If chkNoPopup = 0 Then frmViewer.ViewGIF dlgFile.FileName, vbWhite
        End If
        
    Else
        ' remove all viewers
        For Slot = shpMarker.LBound To shpMarker.UBound
            If Not myGifViewer(Slot) Is Nothing Then Set myGifViewer(Slot) = Nothing
        Next
        cmdLoad(0).Enabled = True   ' re-enable loading viewers
        cmdLoad(1).Enabled = False  ' disable clearing viewers
        Me.Cls
        Frame1.Refresh
        picBkg.Cls
        Set imgFrame = Nothing
    End If
    
EH:
End Sub

Private Sub Command1_Click()
    ' example of changing colors during animation
    
    Dim v As Long, fs As Long
    ' first pause any current animation
    If Not myGifViewer(0) Is Nothing Then Call cmdAction_Click(gfaPaused)
    
    ' change our DC backcolor
    If Me.BackColor = vbButtonFace Then
        Me.BackColor = vbCyan
        fs = vb3DShadow  ' for fuller-size aniamtion
    Else
        Me.BackColor = vbButtonFace
        fs = vbWhite    ' for fuller-size aniamtion
    End If
    Frame1.BackColor = Me.BackColor
    
    ' now restart animation after having the viewer change its backcolor
    If Not myGifViewer(0) Is Nothing Then
        myGifViewer(0).AnimationState = gfaPlaying
        ' skip the picturebox example when changing bkg colors
        For v = shpMarker.LBound + 1 To shpMarker.UBound
            If Not myGifViewer(v) Is Nothing Then
                myGifViewer(v).SetAnimationBkgColor Me.BackColor
                myGifViewer(v).AnimationState = gfaPlaying
            End If
        Next
    End If
    
    For v = 0 To Forms.Count - 1
        If Forms(v).Name = "frmViewer" Then
            frmViewer.ChangeBackColor fs
            Exit For
        End If
    Next
        
End Sub

Private Sub Form_Load()

    Me.ScaleMode = vbPixels         ' class uses pixels, easier calcs if our form is pixels
    Me.picBkg.ScaleMode = vbPixels
    cboStretchMode.ListIndex = 3    ' set stretch mode combobox
    ' create array of empty GIF viewers
    ReDim myGifViewer(shpMarker.LBound To shpMarker.UBound)
    ' set up our common dialog
    With dlgFile
        .CancelError = True
        .Flags = cdlOFNExplorer Or cdlOFNFileMustExist
        .DialogTitle = "Select GIF File"
        .Filter = "GIFs|*.GIF"
    End With
End Sub

Private Sub Form_Paint()
    ' When the form's AutoRedraw is false, this event will be fired whenever
    ' another window covers/uncovers our form or if our form is resized
    
    ' Let's refresh current GIF frames
    Dim f As Long
    For f = shpMarker.LBound To shpMarker.UBound
        If Not myGifViewer(f) Is Nothing Then
            myGifViewer(f).AnimationState = gfaRefresh
        End If
    Next
End Sub

Private Sub Form_Resize()

    ' Save your user some CPU cycles; don't animate if minimized!
    ' If you are subclassing, you can also determine if your window is hidden behind
    ' other windows & you might consider pausing animation when that happens too
    
    If Me.WindowState = vbMinimized Then
        Call cmdAction_Click(gfaPaused) ' pause animation; no need to draw if we are minimized
    Else
        If Not myGifViewer(shpMarker.LBound) Is Nothing Then
            If myGifViewer(shpMarker.LBound).AnimationState = gfaPaused Then
                Call cmdAction_Click(gfaPlaying) ' re-start from pause position
            End If
        End If
    End If
End Sub

Private Sub IGifRender_GetRenderDC(ByVal ViewerID As Long, ByVal FrameIndex As Long, _
        destDC As Long, Optional hwndRefresh As Long, Optional bAutoRedraw As Boolean = False, Optional PostNotify As Boolean = False)
    
    If ViewerID = shpMarker.UBound Then
        destDC = Frame1.hWnd    ' VB Frames have no .HDC, but class will get it from the hWnd
        hwndRefresh = Frame1.hWnd
    ElseIf ViewerID = 0 Then     ' VB picture box example
        destDC = picBkg.hdc
        hwndRefresh = picBkg.hWnd
        bAutoRedraw = picBkg.AutoRedraw
    Else
        destDC = Me.hdc
        hwndRefresh = Me.hWnd
        bAutoRedraw = Me.AutoRedraw
    End If
    'PostNotify = False ' don't need the msgRendered message; otherwise
                        ' set it to True to get that message in IGifRender_Rendered
End Sub

Private Sub IGifRender_Rendered(ByVal ViewerID As Long, ByVal FrameIndex As Long, ByVal Message As RenderMessage, ByVal MsgValue As Long)
    
    ' Three types of messages can be received (more can be added in future rewrites):
    Select Case Message
        Case msgProgress ' occurs during GIF load only and is always received if you
                         ' passed the IGifRender parameter to LoadGIF
                         
                         
            Me.Caption = "Loading... " & FrameIndex & ", " & MsgValue & "% complete"
            If MsgValue = 100 Then
                Me.Caption = "Sample Multi-GIF Viewer"
                
            ElseIf MsgValue = 0 Then ' this means the GIF has been parsed but not processed yet
            
                ' we want the viewer to cache our background for flicker free drawing
                ' And also we want viewer to scale our image for us too; basically,
                ' we want the viewer to be self-sufficient. So supply it with what it
                ' needs...
                
                ' TIP #1: By calling .SetAnimationBkg here, then the 1st frame will immediately
                ' be rendered when it is done being processed. Otherwise, .SetAnimationBkg should
                ' be called after LoadGIF and the 1st frame will be rendered after all frames are processed
                
                ' Tip #2: If this is what you want, ensure the DC being drawn to has AutoRedraw=False, otherwise
                ' the frame will still be rendered by VB may not update the screen immediately
                With shpMarker(ViewerID)
                    If ViewerID = shpMarker.LBound Then ' Picture Box example
                        myGifViewer(ViewerID).SetAnimationBkg gfdBkgFromDC, picBkg.hdc, _
                            .Left + 1, .Top + 1, .Width - 2, .Height - 2, cboStretchMode.ListIndex Or gfsCentered
    
                    ElseIf ViewerID = shpMarker.UBound Then ' Frame example
                        ' VB Frame scalemode is twips regardless of parent container(form)
                        myGifViewer(ViewerID).SetAnimationBkg gfdSolidColor, Frame1.BackColor, _
                            .Left \ Screen.TwipsPerPixelX + 1, .Top \ Screen.TwipsPerPixelY + 1, _
                            .Width \ Screen.TwipsPerPixelX - 2, .Height \ Screen.TwipsPerPixelY - 2, cboStretchMode.ListIndex Or gfsCentered
    
                    Else
                        myGifViewer(ViewerID).SetAnimationBkg gfdSolidColor, Me.BackColor, _
                            .Left + 1, .Top + 1, .Width - 2, .Height - 2, cboStretchMode.ListIndex Or gfsCentered
                    End If
                End With
            End If
        
        Case msgRendered ' informs that a frame was rendered and provides target DC
                         ' as msgValue
            
        Case msgLoopsEnded ' informs that a GIF has stopped animating due to loop threshhold
                           ' was met. This also gets fired when a single-frame GIF is drawn.
                           ' One can use this message to restart animation or possibly to
                           ' get the frame as a stdPicture to assign to an image control?
                           ' The msgValue is the number of loops completed.
            If ViewerID = shpMarker.UBound Then Call MakeFrameImagePermanent(ViewerID)
    End Select

End Sub


Private Sub MakeFrameImagePermanent(ByVal vwrID As Long)

    If myGifViewer(vwrID) Is Nothing Then Exit Sub
    ' this GIF was animating on the VB frame, but frame's don't have autoredraw,
    ' so if form is refreshed we lose the image. This example shows how to
    ' preserve the image using an image control after animation is done...
    
    ' When the VB Frame's parent container has AutoRedraw=False, then simply refreshing
    ' the VB Frame will do the trick; otherwise you'll need something a little more
    ' permanent:
    
    If Frame1.Parent.AutoRedraw = False Then Exit Sub
    '^^ If so, our form (the frame's parent) is refreshing images when it gets
    '   a Paint event so we don't need to run this routine
    
    
    Dim X As Long, Y As Long, CX As Long, CY As Long, Index As Long
    
    imgFrame.Visible = False
    
    Select Case myGifViewer(vwrID).AnimationState
        Case gfaPaused, gfaStepping, gfaStopped
            Index = myGifViewer(vwrID).AnimationIndex ' make last animated frame permanent
        Case gfaRefresh ' nothing to do, we should already have a permanent image
            If imgFrame.Picture Is Nothing Then
                Index = myGifViewer(vwrID).AnimationIndex ' make last animated frame permanent
            Else
                Exit Sub
            End If
        Case gfaPlaying    ' animating already; remove our permanent image
            Set imgFrame.Picture = Nothing
            Frame1.Refresh
            Exit Sub
    End Select
    
    With Screen
        ' remember; VB Frame controls use twips, set max bounding rectangle for image
        Y = shpMarker(vwrID).Top \ .TwipsPerPixelY + 1
        X = shpMarker(vwrID).Left \ .TwipsPerPixelX + 1
        CX = shpMarker(vwrID).Width \ .TwipsPerPixelX - 2
        CY = shpMarker(vwrID).Height \ .TwipsPerPixelY - 2
        
        ' allow viewer to scale and adjust coords for us...
        myGifViewer(vwrID).ScaleToDestination Index, X, Y, CX, CY, gfsShrinkScaleToFit Or gfsCentered, 0, 0
        ' size & position the image control
        imgFrame.Move X * .TwipsPerPixelX, Y * .TwipsPerPixelY, _
            CX * .TwipsPerPixelX, CY * .TwipsPerPixelY
    End With
    ' finish the job
    imgFrame.Stretch = True
    Set imgFrame.Picture = myGifViewer(vwrID).FrameImage(Index, gfiPicGIF)
    imgFrame.Visible = True
    Frame1.Refresh

End Sub
