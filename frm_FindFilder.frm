VERSION 5.00
Begin VB.Form frm_FindFolder 
   Caption         =   "DonkBuilt Find Folder Example"
   ClientHeight    =   1500
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1500
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6105
      TabIndex        =   1
      Top             =   630
      Width           =   1095
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Find Folder"
      Height          =   375
      Left            =   6105
      TabIndex        =   0
      Top             =   210
      Width           =   1095
   End
   Begin VB.Label lblFolder 
      Caption         =   " "
      Height          =   330
      Left            =   180
      TabIndex        =   2
      Top             =   1080
      Width           =   7035
   End
End
Attribute VB_Name = "frm_FindFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'   Copyright © 2001 DonkBuilt Software
'   Written by Allen S. Donker
'   All rights reserved.

'***************************************************************
'   Demonstrates the use of mod_SelectFolder which
'   uses the windows API's to open a browse folder
'   dialog box
'***************************************************************


Private Sub cmdBrowse_Click()
On Error GoTo ErrH

Dim sPath As String

            '   Open the Browse Folder dialog box and
            '   return the folder selected
    sPath = SelectFolder(Me, "Select folder")
  
  
            '   If no folder was selected, exit here
    If Len(sPath) = 0 Then
        lblFolder.Caption = "No folder selected"
        Exit Sub
    Else
        lblFolder.Caption = sPath
    End If

Exit Sub
    
ErrH:
    MsgBox Err.Number & Chr(10) & Err.Description
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub
