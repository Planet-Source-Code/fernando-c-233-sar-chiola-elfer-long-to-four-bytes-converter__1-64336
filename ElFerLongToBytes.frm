VERSION 5.00
Begin VB.Form Main 
   Caption         =   "ElFerLongToBytes"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Byte4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Byte3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Byte2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Byte1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Convert 
      Caption         =   "Convert"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox LongNumber 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Long, ByVal ByteLen As Long)
Const MAXLONG = &H7FFFFFFF 'Max Long at VB is 2147483647
Public Function LongToBytes(MyLong As Long) As String
    Dim MemoryAddres As Long
    Dim sBuffer As String
    MemoryAddres = VarPtr(MyLong) 'Get the Memory addres of Long
    If MemoryAddres <> 0 Then
          sBuffer = Space$(4)     'String of four spaces
          CopyMemory ByVal sBuffer, ByVal MemoryAddres, 4  'Read the memory as string
          LongToBytes = sBuffer
    End If
End Function




Private Sub Convert_Click()
        Dim OneLong As Long
        If Val(LongNumber.Text) > MAXLONG Then
            MsgBox LongNumber + " is to long. :D"
        Else
            OneLong = Val(LongNumber.Text)
            StringLong = LongToBytes(OneLong)
            Byte1.Text = Asc(Mid$(StringLong, 1, 1))
            Byte2.Text = Asc(Mid$(StringLong, 2, 1))
            Byte3.Text = Asc(Mid$(StringLong, 3, 1))
            Byte4.Text = Asc(Mid$(StringLong, 4, 1))
        End If
End Sub


