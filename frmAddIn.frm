VERSION 5.00
Begin VB.Form frmAddIn 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Close All VB Windows"
   ClientHeight    =   840
   ClientLeft      =   2130
   ClientTop       =   1605
   ClientWidth     =   3765
   Icon            =   "frmAddIn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3285
      Top             =   840
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   615
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Close All Windoiws"
      Height          =   465
      Left            =   435
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Closing all windows..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   217
      TabIndex        =   2
      Top             =   225
      Width           =   3330
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE

Public Connect As Connect

Option Explicit

Private Sub CancelButton_Click()
    Connect.Hide
End Sub

Private Sub Form_Activate()

'OKButton_Click

End Sub

Private Sub Form_Resize()
Timer1.Enabled = True
End Sub

Private Sub OKButton_Click()
    Dim i As Integer
    Dim p As Integer
    
   
    i = 0
    While i < VBInstance.Windows.Count
        i = i + 1
    
        If VBInstance.Windows(i).Type = vbext_wt_Designer Or VBInstance.Windows(i).Type = vbext_wt_CodeWindow Then
           VBInstance.Windows(i).Close
           i = 0
        End If
        
    Wend
    
    Unload Me
    
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
  OKButton_Click
End Sub
