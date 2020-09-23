VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bitmap To JPEG    By: Eyal Perry"
   ClientHeight    =   3084
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   3936
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3084
   ScaleWidth      =   3936
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save As JPEG"
      Height          =   492
      Left            =   2040
      TabIndex        =   2
      Top             =   2520
      Width           =   1812
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Bitmap File"
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   1812
   End
   Begin VB.PictureBox picMain 
      Height          =   2292
      Left            =   120
      ScaleHeight     =   2244
      ScaleWidth      =   3684
      TabIndex        =   0
      Top             =   120
      Width           =   3732
      Begin MSComDlg.CommonDialog dlg 
         Left            =   360
         Top             =   600
         _ExtentX        =   677
         _ExtentY        =   677
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BmpToJpeg Lib "Bmp2Jpeg.dll" (ByVal BmpFilename As String, ByVal JpegFilename As String, ByVal CompressQuality As Integer) As Integer

Private Sub cmdLoad_Click()
Dim filename As String
On Error Resume Next
'Set the filter to bitmap pictures
dlg.Filter = "Bitmap Files | *.bmp"
dlg.filename = ""
dlg.ShowOpen
'Get the file name
filename = dlg.filename
'Load the picture
picMain.Picture = LoadPicture(filename)
End Sub

'For saving into jpg format you must have
'the Bmp2Jpeg.dll file in your windows folder

Private Sub cmdSave_Click()
Dim filename As String
On Error Resume Next
'Set the filter to JPEG files
dlg.Filter = "JPEG files | *.jpg"
dlg.filename = ""
dlg.ShowOpen
'Get the file name without the ".jpg"
filename = Left(dlg.filename, Len(dlg.filename) - 4)
'Saving the file as Bitmap
SavePicture picMain.Picture, filename & ".bmp"
'Change the bitmap file to the jpg file with the Bmp2Jpg.dll
'100 is the Compress Quality
BmpToJpeg filename & ".bmp", filename & ".jpg", 100
'Deleting the bitmap file
Kill filename & ".bmp"
End Sub
