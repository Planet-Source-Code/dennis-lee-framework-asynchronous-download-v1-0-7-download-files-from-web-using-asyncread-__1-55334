VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":08CA
   ScaleHeight     =   3345
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin AsyncDownloadFramework.Downloader Downloader1 
      Left            =   600
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   600
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":243E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstDownload 
      Height          =   2655
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "SmallImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   2716
      EndProperty
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    cmdDownload.Enabled = True
    cmdCancel.Enabled = False
    Me.Downloader1.CancelAllDownload

End Sub

Private Sub cmdDownload_Click()

  Dim sURL As String
  Dim sFilename As String
  Dim sDescription As String

    cmdCancel.Enabled = True
    cmdDownload.Enabled = False
   
    lstDownload.ListItems.Clear

    sURL = "http://download.microsoft.com/download/ie6sp1/finrel/6_sp1/W98NT42KMeXP/EN-US/ie6setup.exe"
    sFilename = "ie6setup.exe"
    sDescription = "Internet Explorer 6 Setup File"
    Download_File sURL, App.Path & "\" & sFilename, sDescription

    sURL = "http://optusnet.dl.sourceforge.net/sourceforge/vnc-tight/tightvnc-1.2.9-setup.exe"
    sFilename = "tightvnc-1.2.9-setup.exe"
    sDescription = "TightVNC 1.2.9 Setup File"
    Download_File sURL, App.Path & "\" & sFilename, sDescription

    sURL = "http://www.planet-source-code.com/vb/images/PscLogo1.jpg"
    sFilename = "PscLogo1.jpg"
    sDescription = "Planet Source Code image"
    Download_File sURL, App.Path & "\" & sFilename, sDescription

    sURL = "http://www.winzip.com/index.htm"
    sFilename = "index.html"
    sDescription = "Winzip homepage"
    Download_File sURL, App.Path & "\" & sFilename, sDescription

    sURL = "http://download.winzip.com/winzip90.exeX" '<--- do that on purpose to show error
    sFilename = "winzip90.exe"
    sDescription = "Winzip 9.0"
    Download_File sURL, App.Path & "\" & sFilename, sDescription

End Sub

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub Downloader1_DownloadAllComplete(FileNotDownload() As String)

  Dim i As Integer

    Debug.Print "Finished all download"
    cmdDownload.Enabled = True
    cmdCancel.Enabled = False

    If UBound(FileNotDownload) > 0 Then
        For i = 1 To UBound(FileNotDownload)
            Debug.Print "File not downloaded: " & FileNotDownload(i)
        Next i
    End If

End Sub

Private Sub Form_Load()

    Me.Caption = App.Title & " v" & App.Major & "." & App.Minor & " Build " & App.Revision
    Debug.Print

    lstDownload.ListItems.Clear

End Sub

Private Sub Downloader1_DownloadComplete(MaxBytes As Long, SaveFile As String)

  Dim i As Integer

    Debug.Print "Completed " & SaveFile & ", Size = " & MaxBytes

    With lstDownload
        For i = 1 To .ListItems.Count
            If .ListItems(i).Key = SaveFile Then
                .ListItems(i).SubItems(1) = "Completed"
            End If
        Next i
    End With

End Sub

Private Sub Downloader1_DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)

  Dim i As Integer
  Dim RemBytes As Long

    With lstDownload
        For i = 1 To .ListItems.Count
            If .ListItems(i).Key = SaveFile Then
                RemBytes = MaxBytes - CurBytes
                If RemBytes < 2 ^ 20 Then
                    .ListItems(i).SubItems(1) = Format((MaxBytes - CurBytes) / 2 ^ 10, "#0.0 KB") & _
                               " (" & Format(CurBytes / MaxBytes, "#0.0%") & ")"
                  Else
                    .ListItems(i).SubItems(1) = Format((MaxBytes - CurBytes) / 2 ^ 20, "#0.00 MB") & _
                               " (" & Format(CurBytes / MaxBytes, "#0.0%") & ")"
                End If
            End If
        Next i

    End With

End Sub

Private Sub Downloader1_DownloadError(SaveFile As String)

  Dim i As Integer

    Debug.Print "Error downloading " & SaveFile

    With lstDownload
        For i = 1 To .ListItems.Count
            If .ListItems(i).Key = SaveFile Then
                .ListItems(i).SubItems(1) = "Error"
            End If
        Next i

    End With

End Sub

Private Function Download_File(URL As String, SaveFile As String, Description As String)

    lstDownload.ListItems.Add , SaveFile, Description, , 1

    Me.Downloader1.BeginDownload URL, SaveFile

End Function

Private Function GetFilename(URL As String) As String

  Dim i As Integer

    For i = Len(URL) To 1 Step -1
        If Mid(URL, i, 1) = "/" Then
            GetFilename = Right(URL, Len(URL) - i)
            Exit For
        End If
    Next i

End Function
