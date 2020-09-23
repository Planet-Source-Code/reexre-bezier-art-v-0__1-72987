VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fMAIN 
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   538
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1008
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fBezier 
      Caption         =   "Bezier"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   9840
      TabIndex        =   14
      Top             =   1200
      Width           =   2295
      Begin VB.HScrollBar sNCP 
         Height          =   255
         Left            =   120
         Max             =   10
         Min             =   3
         TabIndex        =   15
         Top             =   840
         Value           =   5
         Width           =   1455
      End
      Begin VB.Label labNCP 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Number of Bezier Control Points"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S T A R T"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9840
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Frame fVideo 
      Caption         =   "Video"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   12480
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
      Begin VB.CommandButton cmdBuildAVI 
         Caption         =   "Make AVI (Special)"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Make AVI with Frames in ""frames"" folder"
         Top             =   4080
         Width           =   2055
      End
      Begin VB.CommandButton cmdStopAndMakeAVI 
         Caption         =   "Stop and Make AVI"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   3120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chPLAY 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Auto PLAY Avi"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   8
         ToolTipText     =   "AutoPlay AVI when It's Created"
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CommandButton cmdPLAY 
         Caption         =   "Player ..."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Select Your AVI Player (.exe)"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtFPS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "24"
         ToolTipText     =   "Output FPS"
         Top             =   360
         Width           =   495
      End
      Begin VB.ComboBox cmbEXTRA 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         TabIndex        =   5
         Text            =   "Combo1"
         ToolTipText     =   "EXTRA FRAMES"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox tSeconds 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Text            =   "60"
         ToolTipText     =   "Video Length (Seconds)"
         Top             =   840
         Width           =   735
      End
      Begin VB.HScrollBar sFrameStep 
         Height          =   255
         Left            =   120
         Max             =   15
         Min             =   1
         TabIndex        =   3
         Top             =   1800
         Value           =   10
         Width           =   1455
      End
      Begin VB.Label LabFS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Get Video Frame Every [N] pictures"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Labelfps 
         Caption         =   "Video FPS"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label LabelVL 
         Caption         =   "Video Lenght"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   10560
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   4815
      Left            =   120
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "fMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private BE    As New clsBezier


Private i     As Long
Private J     As Long


Private P()   As mPoint

Private QD    As Double
Private D     As Double
Private s     As Double


Private NP    As Long





Private RealFrame As Long
Private VideoLenFrames As Long

Public OutputAVIName As String
Public AVIPLAYER As String

Private TextToPrint As String
Private PrintX As Single


Private Sub StopAndMakeAVI()
    BUILD_AVI
    If chPLAY.Value = Checked Then
        If (AVIPLAYER <> "") Then
            If OutputAVIName <> "" Then
                Shell AVIPLAYER & " " & Chr$(34) & OutputAVIName, vbNormalFocus
            End If
        Else
            MsgBox "No Avi Player Selected!", vbCritical
        End If

    End If
End Sub

Private Sub cmdBuildAVI_Click()
StopAndMakeAVI
End Sub

Private Sub cmdStopAndMakeAVI_Click()
    RealFrame = VideoLenFrames

End Sub

Private Sub Command1_Click()
Dim R         As Single
Dim G         As Single
Dim B         As Single


Dim dx        As Single
Dim dy        As Single
Dim F         As Single
Dim FX        As Single
Dim Fy        As Single

 TextToPrint = "BEZIER_ART  v.0     Software by Reexre [Roberto Mior]"
 TextToPrint = TextToPrint & "     Bezier Control points:" & sNCP & "     "
 TextToPrint = TextToPrint & "     Video " & PIC.Width & "X" & PIC.Height & " at " & txtFPS & " FPS"

    If Dir(App.Path & "\frames\*.bmp") <> "" Then Kill App.Path & "\frames\*.bmp"
    cmdStopAndMakeAVI.Visible = True
    PrintX = PIC.Width + 100

    Randomize Timer

    BE.InitCurve sNCP '5    '8    '12 '5

    ReDim P(0 To BE.GetNumVerts - 1)


    For i = 0 To BE.GetNumVerts - 1
        BE.SetPointCoords i, 10 + Rnd * (PIC.Width - 20), 10 + Rnd * (PIC.Height - 20)

        P(i).X = BE.GetPointX(i)
        P(i).Y = BE.GetPointY(i)
        P(i).vX = Rnd * 2 - 1
        P(i).vY = Rnd * 2 - 1

    Next


    NP = BE.GetNumVerts

    BE.InitTarget PIC.Image.Handle


    VideoLenFrames = Val(txtFPS) * Val(tSeconds) * sFrameStep
    'Stop

    For RealFrame = 1 To VideoLenFrames
        Me.Caption = "Real Frame:" & RealFrame & "    Video Frame:" & RealFrame \ sFrameStep & "    Video Seconds:" & (RealFrame \ sFrameStep) / Val(txtFPS)


        DoEvents

        For i = 0 To BE.GetNumVerts - 2
            For J = i + 1 To BE.GetNumVerts - 1

                'Stop
                dx = BE.GetPointX(i) - BE.GetPointX(J)
                dy = BE.GetPointY(i) - BE.GetPointY(J)
                QD = QuadDistance(dx, dy)
                D = Sqr(QD)

                'Stop

                F = Force(QD)

                If D < 40 Then
                    s = -1
                Else
                    s = 1
                End If

                FX = F * Sgn(dx) * s * 2
                Fy = F * Sgn(dy) * s * 2

                'FX = f * dx * S * 0.1
                'Fy = f * dy * S * 0.1


                P(i).vX = P(i).vX - FX    '* P(I).kF
                P(i).vY = P(i).vY - Fy    '* P(I).kF
                P(J).vX = P(J).vX + FX    '* P(J).kF
                P(J).vY = P(J).vY + Fy    '* P(J).kF
                'Stop

                'R = (Sin(RealFrame / 80) + 1) * 100 + 55
                'G = (Sin(RealFrame / 90) + 1) * 100 + 55
                'B = (Sin(RealFrame / 170) + 1) * 100 + 55
                
                R = (Cos(RealFrame / 100) + 1) * 255
                G = (Cos(RealFrame / 127) + 1) * 255
                B = (Cos(RealFrame / 143) + 1) * 255

            Next J
        Next

        'Stop

        BE.Render PIC, 500, R, G, B


        For i = 0 To BE.GetNumVerts - 1
            P(i).X = P(i).X + P(i).vX
            P(i).Y = P(i).Y + P(i).vY
            If P(i).X < 5 Then P(i).X = 5: P(i).vX = -P(i).vX
            If P(i).Y < 5 Then P(i).Y = 5: P(i).vY = -P(i).vY
            If P(i).X > PIC.Width - 5 Then P(i).X = PIC.Width - 5: P(i).vX = -P(i).vX
            If P(i).Y > PIC.Height - 5 Then P(i).Y = PIC.Height - 5: P(i).vY = -P(i).vY


            P(i).vX = P(i).vX * 0.999
            P(i).vY = P(i).vY * 0.999



            BE.SetPointCoords i, P(i).X, P(i).Y


        Next i

        If RealFrame Mod sFrameStep = 0 Then

            PIC.CurrentX = PrintX
            PrintX = PrintX - Val(txtFPS) / 8
            If PrintX < -Len(TextToPrint) * 20 Then PrintX = PIC.Width + 100


            PIC.CurrentY = PIC.Height - 20
            PIC.Print TextToPrint
            '            Stop

            SavePicture PIC.Image, App.Path & "\frames\P" & Format(RealFrame, "000000") & ".bmp"
        End If

    Next RealFrame

    StopAndMakeAVI

End Sub





Private Sub Form_Load()

ProcessPrioritySet , , ppidle    'ppbelownormal ' So While is Computing You Can to Other

PIC.Height = 360 '270 '360
    PIC.Width = Int(PIC.Height * 16 / 9)


    If Dir(App.Path & "\frames\", vbDirectory) = "" Then MkDir App.Path & "\frames\"


    cmbEXTRA.AddItem "0"
    cmbEXTRA.AddItem "1"
    cmbEXTRA.ListIndex = 0
    If Dir(App.Path & "\Player.txt") <> "" Then
        Open App.Path & "\Player.txt" For Input As 22
        Input #22, AVIPLAYER
        Close 22
    End If

    LabFS = sFrameStep
    labNCP = sNCP
    

   

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End

End Sub

Private Sub chPLAY_Click()
    If chPLAY And (AVIPLAYER = "") Then MsgBox "Select Avi Player": cmdPLAY_Click
End Sub
Private Sub cmdPLAY_Click()
    With CMD
        .filename = ""
        .InitDir = "RealFrame:\"
        .Filter = "AVI Player|*.EXE"    ';*.mpg"
        .DialogTitle = "Select AVI PLAYER"
    End With
    CMD.Action = 1

    If CMD.filename <> "" Then
        AVIPLAYER = CMD.filename
        Open App.Path & "\Player.txt" For Output As 22
        Print #22, AVIPLAYER
        Close 22
    End If
End Sub
Public Sub BUILD_AVI()

    OutputAVIName = ""

    Dim fPATH As String

    fPATH = App.Path & "\frames\"

    Dim s     As String

    Dim fLIST() As String
    Dim RealFrame As Long

    s = Dir(fPATH & "*.bmp")

    If s = "" Then Exit Sub

    ReDim Preserve fLIST(1)
    Do
        fLIST(RealFrame) = fPATH & s
        RealFrame = RealFrame + 1
        ReDim Preserve fLIST(0 To RealFrame)
        s = Dir
    Loop While s <> ""

    '----------------------------------------------------------------------------------------
    Dim file  As cFileDlg
    Dim InitDir As String
    Dim szOutputAVIFile As String
    Dim res   As Long
    Dim pfile As Long    'ptr PAVIFILE
    Dim bmp   As cDIB
    Dim ps    As Long    'ptr PAVISTREAM
    Dim psCompressed As Long    'ptr PAVISTREAM
    Dim strhdr As AVI_STREAM_INFO
    Dim BI    As BITMAPINFOHEADER
    Dim opts  As AVI_COMPRESS_OPTIONS
    Dim pOpts As Long
    Dim i     As Long
    Dim I2    As Long

    Dim EXTRA As Integer    'Extra Frame

    Debug.Print
    Set file = New cFileDlg
    'get an avi filename from user
    With file
        .InitDirectory = App.Path & "\VIDEO\"
        .DefaultExt = "avi"
        .DlgTitle = "Choose a filename to save AVI to..."
        .Filter = "AVI Files|*.avi"
        .OwnerHwnd = fMAIN.hWnd
    End With
    szOutputAVIFile = "MyAVI.avi"
    If file.VBGetSaveFileName(szOutputAVIFile) <> True Then Exit Sub


    OutputAVIName = szOutputAVIFile
    'Stop

    '    Open the file for writing
    res = AVIFileOpen(pfile, szOutputAVIFile, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then GoTo error

    'Get the first bmp in the list for setting format
    Set bmp = New cDIB
    If bmp.CreateFromFile(fLIST(1)) <> True Then
        MsgBox "Could not load first bitmap file in list!", vbExclamation, App.title
        GoTo error
    End If
    'Stop

    '   Fill in the header for the video stream
    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)    '// stream type video
        .fccHandler = 0&    '// default AVI handler
        .dwScale = 1
        .dwRate = Val(txtFPS) * (Val(cmbEXTRA) + 1)    '// fps
        .dwSuggestedBufferSize = bmp.SizeImage    '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, bmp.Width, bmp.Height)    '// rectangle for stream
    End With

    'validate user input
    If strhdr.dwRate < 1 Then strhdr.dwRate = 1
    If strhdr.dwRate > 30 Then strhdr.dwRate = 30

    '   And create the stream
    res = AVIFileCreateStream(pfile, ps, strhdr)
    If (res <> AVIERR_OK) Then GoTo error

    'get the compression options from the user
    'Careful! this API requires a pointer to a pointer to a UDT
    pOpts = VarPtr(opts)
    res = AVISaveOptions(Me.hWnd, _
                         ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, _
                         1, _
                         ps, _
                         pOpts)    'returns TRUE if User presses OK, FALSE if Cancel, or error code
    If res <> 1 Then    'In RealFrame TRUE = 1
        Call AVISaveOptionsFree(1, pOpts)
        GoTo error
    End If

    'make compressed stream
    res = AVIMakeCompressedStream(psCompressed, ps, opts, 0&)
    If res <> AVIERR_OK Then GoTo error

    'set format of stream according to the bitmap
    With BI
        .biBitCount = bmp.BitCount
        .biClrImportant = bmp.ClrImportant
        .biClrUsed = bmp.ClrUsed
        .biCompression = bmp.Compression
        .biHeight = bmp.Height
        .biWidth = bmp.Width
        .biPlanes = bmp.Planes
        .biSize = bmp.SizeInfoHeader
        .biSizeImage = bmp.SizeImage
        .biXPelsPerMeter = bmp.XPPM
        .biYPelsPerMeter = bmp.YPPM
    End With

    'set the format of the compressed stream
    res = AVIStreamSetFormat(psCompressed, 0, ByVal bmp.PointerToBitmapInfo, bmp.SizeBitmapInfo)
    If (res <> AVIERR_OK) Then GoTo error

    '   Now write out each video frame
    I2 = 0
    For i = 0 To RealFrame - 1

        Me.Caption = "Creating AVI file...  Frame " & i & " of " & RealFrame - 1
        DoEvents


        For EXTRA = 0 To Val(cmbEXTRA)

            bmp.CreateFromFile (fLIST(i))    'load the bitmap (ignore errors)


            res = AVIStreamWrite(psCompressed, _
                                 I2, _
                                 1, _
                                 bmp.PointerToBits, _
                                 bmp.SizeImage, _
                                 AVIIF_KEYFRAME, _
                                 ByVal 0&, _
                                 ByVal 0&)
            If res <> AVIERR_OK Then GoTo error
            'Show user feedback
            'imgPreview.Picture = LoadPicture(lstDIBList.Text)
            'imgPreview.Refresh
            'lblStatus = "Frame number " & i & " saved"
            'lblStatus.Refresh
            I2 = I2 + 1

        Next EXTRA


    Next
    Me.Caption = "Avi file  Created!"



error:
    '   Now close the file
    Set file = Nothing
    Set bmp = Nothing

    If (ps <> 0) Then Call AVIStreamClose(ps)

    If (psCompressed <> 0) Then Call AVIStreamClose(psCompressed)

    If (pfile <> 0) Then Call AVIFileClose(pfile)

    Call AVIFileExit

    If (res <> AVIERR_OK) Then
        MsgBox "There was an error writing the file.", vbInformation, App.title
    End If

    'Stop


End Sub

Private Sub sFrameStep_Change()
    LabFS = sFrameStep
End Sub

Private Sub sFrameStep_Scroll()
    LabFS = sFrameStep
End Sub

Private Sub sNCP_Change()
labNCP = sNCP

End Sub

Private Sub sNCP_Scroll()
labNCP = sNCP
End Sub
