VERSION 5.00
Begin VB.Form frmNoteHelp 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notepad"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7605
   Icon            =   "frmNoteHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraNoteHelp 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4335
      Left            =   120
      MouseIcon       =   "frmNoteHelp.frx":030A
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.Timer tmrNote 
         Interval        =   500
         Left            =   120
         Top             =   3840
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00FFC0C0&
         Caption         =   "OK"
         Height          =   375
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Gm Softwares Pvt. Ltd."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   240
         Left            =   600
         TabIndex        =   4
         Top             =   1320
         Width           =   1980
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Jammu (J&&K)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   3
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Anoop Chargotra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   2
         Top             =   1080
         Width           =   1530
      End
      Begin VB.Image imgNoteHelp 
         Height          =   3450
         Left            =   0
         Picture         =   "frmNoteHelp.frx":0614
         Top             =   2040
         Width           =   12000
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "NOTEPAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   1890
      End
   End
End
Attribute VB_Name = "frmNoteHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim boolValid3 As Boolean

Private Sub cmdOk_Click()

 ' This subroutine Unloads the Credits Form
   Unload frmNoteHelp

End Sub

Private Sub tmrNote_Timer()
  
 ' This subroutine flashes the Notepad in the caption
   If boolValid3 Then
      lblCaption.ForeColor = vbBlue
   Else
      lblCaption.ForeColor = &HFFC0C0
   End If
   boolValid3 = Not boolValid3
  
End Sub
