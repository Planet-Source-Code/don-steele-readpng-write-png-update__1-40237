VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Load PNG Example"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw Color Bars"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save PNG File"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load PNG File"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox pctMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2610
      Left            =   240
      ScaleHeight     =   2610
      ScaleWidth      =   3255
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDraw_Click()

    Dim i As Long
    For i = 0 To pctMain.Width - 15 Step 15
        pctMain.Line (i, 0)-Step(0, pctMain.Height / 3), RGB(i / pctMain.Width * 255, 0, 0)
        pctMain.Line (i, pctMain.Height / 3)-Step(0, pctMain.Height / 3), RGB(0, i / pctMain.Width * 255, 0)
        pctMain.Line (i, pctMain.Height / 3 * 2)-Step(0, pctMain.Height / 3), RGB(0, 0, i / pctMain.Width * 255)
    Next

End Sub

Private Sub cmdLoad_Click()
    'Note that the picture box should be set AutoRedraw = True

    If Not LoadPNGFile(App.Path & "\example.png", pctMain) Then
        MsgBox "There was an error loading the example.png file", vbCritical, "Error"
    End If
        

End Sub

Private Sub cmdSave_Click()

    If Dir(App.Path & "\output.png") <> "" Then
        MsgBox "The output file, """ & App.Path & "\output.png"" already exists, save aborted.", vbCritical, "Error"
        Exit Sub
    End If

    If Not SavePNGFile(pctMain, App.Path & "\output.png") Then
        MsgBox "There was an error saving the output file", vbCritical, "Error"
    Else
        MsgBox """" & App.Path & "\output.png"" has been saved.", vbInformation, "Example"
    End If

End Sub

