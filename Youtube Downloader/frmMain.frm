VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   Caption         =   "Video Downloader "
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   6975
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtClip 
      Height          =   285
      Left            =   5700
      TabIndex        =   17
      Top             =   2190
      Visible         =   0   'False
      Width           =   675
   End
   Begin ComctlLib.StatusBar STBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   16
      Top             =   3150
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7937
            MinWidth        =   7937
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6174
            MinWidth        =   6174
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Downloading Information"
      Height          =   2115
      Left            =   90
      TabIndex        =   5
      Top             =   930
      Width           =   4245
      Begin ComctlLib.ProgressBar PB1 
         Height          =   285
         Left            =   1290
         TabIndex        =   14
         Top             =   1680
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   503
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label Label4 
         Caption         =   "Progress"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1710
         Width           =   1005
      End
      Begin VB.Label lblPercent 
         Caption         =   "???"
         Height          =   255
         Left            =   1290
         TabIndex        =   13
         Top             =   1350
         Width           =   2805
      End
      Begin VB.Label Label123 
         Caption         =   "Percent"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1380
         Width           =   1005
      End
      Begin VB.Label lblSaved 
         Caption         =   "???"
         Height          =   255
         Left            =   1290
         TabIndex        =   11
         Top             =   1050
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Saved"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Label lblRemaining 
         Caption         =   "???"
         Height          =   255
         Left            =   1290
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Remaining"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label lblVidSize 
         Caption         =   "???"
         Height          =   255
         Left            =   1290
         TabIndex        =   7
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Video Size"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   390
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "&Download Video"
      Height          =   1995
      Left            =   4410
      TabIndex        =   4
      Top             =   990
      Width           =   2475
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   5070
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4500
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtURL 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Text            =   "http://www.youtube.com/watch?v=Rfr9bhSmfXc"
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label lblVidName 
      Caption         =   "???"
      Height          =   255
      Left            =   1290
      TabIndex        =   3
      Top             =   540
      Width           =   5595
   End
   Begin VB.Label Label2 
      Caption         =   "Video Name"
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   540
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Youtube URL"
      Height          =   225
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   1035
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ResetControls()
    txtURL.Text = ""
    lblVidName.Caption = ""
    lblVidID = ""
    lblVidSize = ""
    lblRemaining = ""
    lblSaved = ""
    lblPercent = ""
    lblSpeed = ""
    PB1.Value = 0
End Sub

Private Sub cmdDownload_Click()
    If InStr(1, txtURL.Text, "youtube") Then
        DownloadVideo GetVideoInfo(txtURL.Text, Inet1), VideoName & ".flv"
    End If
End Sub

Private Sub txtURL_Change()
    txtClip.Text = Clipboard.GetText(vbCFText)
    txtURL.Text = Left(txtClip.Text, 42)
End Sub

