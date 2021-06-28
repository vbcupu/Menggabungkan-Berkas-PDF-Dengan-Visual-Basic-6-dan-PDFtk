VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGabungkanFIle 
      Caption         =   "Gabungkan File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtFileOutput 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   5535
   End
   Begin VB.TextBox txtFolderAsal 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.CommandButton cmdPilihFolder 
      Caption         =   "Pilih Folder Asal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_EDITBOX = &H10
Private Const BIF_NEWDIALOGSTYLE = &H40
 
Private Const MAX_PATH = 260
 
Private Declare Function SHBrowseForFolder Lib _
        "shell32" (lpbi As BrowseInfo) As Long
 
Private Declare Function SHGetPathFromIDList Lib _
        "shell32" (ByVal pidList As Long, ByVal lpBuffer _
        As String) As Long
 
Private Declare Function lstrcat Lib "kernel32" _
        Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
        lpString2 As String) As Long
 
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Sub cmdGabungkanFIle_Click()
    Shell ("pdftk " & txtFolderAsal.Text & "/*.pdf cat output " & App.Path & "/" & txtFileOutput.Text & ".pdf")
    Call MsgBox("Tersimpan di " & App.Path & "/output/" & txtFileOutput.Text & ".pdf")
End Sub

Private Sub cmdPilihFolder_Click()
   
'=============================
Dim lpIDList As Long
Dim sBuffer As String
Dim sTitle As String
Dim tBrowseInfo As BrowseInfo
 
    sTitle = "Find Directory"
    With tBrowseInfo
            .hWndOwner = Me.hWnd
            .lpszTitle = lstrcat(sTitle, "")
            .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN _
                        + BIF_EDITBOX + BIF_NEWDIALOGSTYLE
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
            sBuffer = Space(MAX_PATH)
            SHGetPathFromIDList lpIDList, sBuffer
            sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
            txtFolderAsal.Text = sBuffer
    End If
 

End Sub

Private Sub Form_Load()
    txtFileOutput.Text = ""
End Sub
