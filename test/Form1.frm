VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3396
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3396
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   348
      Left            =   672
      TabIndex        =   1
      Text            =   "D:\TEMP\WebCatalog_2015_04_15.bak"
      Top             =   756
      Width           =   6648
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   852
      Left            =   672
      TabIndex        =   0
      Top             =   1344
      Width           =   1692
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ApiEmptyByteArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal VarType As VbVarType = vbByte, Optional ByVal Low As Long = 0, Optional ByVal Count As Long = 0) As Byte()

Public Function ReadBinaryFile(sFile As String) As Byte()
    Dim baBuffer()      As Byte
    Dim nFile           As Integer
    Dim vErr            As Variant
    
    On Error GoTo EH
    baBuffer = ApiEmptyByteArray()
    nFile = FreeFile
    Open sFile For Binary Access Read As nFile
    If LOF(nFile) > 0 Then
        ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
        Get nFile, , baBuffer
    End If
    Close nFile
    ReadBinaryFile = baBuffer
    Exit Function
EH:
    Close nFile
End Function

Private Sub Command1_Click()
    Dim baFile()        As Byte
    Dim baCompressed()  As Byte
    Dim baUncompressed() As Byte
    Dim lIdx            As Long
    Dim dblTimer        As Double
    Dim dblTotal        As Double
    
    baFile = ReadBinaryFile(Text1.Text)
    baCompressed = PithyCompress(baFile)
    Debug.Print "Compression ratio: " & Format((UBound(baCompressed) + 1) * 100# / (UBound(baFile) + 1), "0.00") & "% (" & Format(UBound(baFile) + 1, "#,0") & " -> " & Format(UBound(baCompressed) + 1, "#,0") & ")"
    baUncompressed = PithyDecompress(baCompressed)
    For lIdx = 0 To UBound(baFile)
        If baFile(lIdx) <> baUncompressed(lIdx) Then
            Debug.Print "Difference at index " & lIdx
        End If
    Next
    dblTimer = Timer
    For lIdx = 1 To 100
        baCompressed = PithyCompress(baFile)
    Next
    dblTimer = (Timer - dblTimer) * 1000
    dblTotal = (UBound(baFile) + 1) * CDbl(lIdx - 1) / 1024# / 1024#
    Debug.Print "Compression speed: " & Format(dblTotal, "0.00") & "MB in " & Format(dblTimer, "0") & " ms -> " & Format(dblTotal * 1000 / dblTimer, "0.00") & " MB/s"

    dblTimer = Timer
    For lIdx = 1 To 100
        baUncompressed = PithyDecompress(baCompressed)
    Next
    dblTimer = (Timer - dblTimer) * 1000
    dblTotal = (UBound(baUncompressed) + 1) * CDbl(lIdx - 1) / 1024# / 1024#
    Debug.Print "Decompression speed: " & Format(dblTotal, "0.00") & "MB in " & Format(dblTimer, "0") & " ms -> " & Format(dblTotal * 1000 / dblTimer, "0.00") & " MB/s"
End Sub
