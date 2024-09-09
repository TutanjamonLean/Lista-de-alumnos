VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   14115
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   4680
      TabIndex        =   7
      Top             =   1920
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   5
      Left            =   2760
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   4
      Left            =   2760
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   3
      Left            =   2760
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   2
      Left            =   2760
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   1
      Left            =   2760
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   4320
      TabIndex        =   0
      Top             =   6480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case Index
        
        Case Is = 0, Is = 1
            If KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii >= 97 And KeyAscii <= 122 Then
                
                KeyAscii = KeyAscii
                
            Else
                
                KeyAscii = 0
                
            End If
            
        Case Is <> 0, Is <> 1
            If KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 47 And KeyAscii <= 57) Then
                
                KeyAscii = KeyAscii
                
            Else
                
                KeyAscii = 0
                
            End If
            
    End Select
    
End Sub
