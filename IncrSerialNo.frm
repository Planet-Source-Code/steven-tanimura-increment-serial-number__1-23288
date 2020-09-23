VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form IncrSerialNo 
   Caption         =   "Increment Serial Number"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid msgResults 
      Height          =   2055
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   8
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.TextBox txtPrevSerial 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblPrevNbr 
      Caption         =   "Inital Serial Number"
      Height          =   280
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "IncrSerialNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sPrevSerial As String

Private Sub txtPrevSerial_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13, 9
        sPrevSerial = txtPrevSerial.Text
        genSerial
    End Select
End Sub

Private Sub txtPrevSerial_LostFocus()
    sPrevSerial = txtPrevSerial.Text
End Sub


Private Sub genSerial()
    Dim i As Integer
    Dim j As Integer
    Dim sTemp As String
    
    sTemp = sPrevSerial
    
    With msgResults
        For i = 1 To 8
            For j = 1 To 4
                sTemp = IncrSerial(sTemp)
                .TextMatrix(i - 1, j - 1) = sTemp
            Next j
        Next i
    End With
End Sub
