Attribute VB_Name = "Module1"
Option Explicit

Public Function PithyCompress(baData() As Byte, Optional ByVal CompressionLevel As Long = 9) As Byte()
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    
    lSize = vbpithy_MaxCompressedLength(UBound(baData) + 1)
    ReDim baBuffer(0 To lSize) As Byte
    lSize = vbpithy_Compress(VarPtr(baData(0)), UBound(baData) + 1, VarPtr(baBuffer(0)), UBound(baBuffer) + 1, CompressionLevel)
    If lSize > 0 Then
        ReDim Preserve baBuffer(0 To lSize - 1)
        PithyCompress = baBuffer
    End If
End Function

Public Function PithyDecompress(baCompressed() As Byte) As Byte()
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    
    Call vbpithy_GetDecompressedLength(VarPtr(baCompressed(0)), UBound(baCompressed) + 1, lSize)
    If lSize > 0 Then
        ReDim baBuffer(0 To lSize - 1) As Byte
        Call vbpithy_Decompress(VarPtr(baCompressed(0)), UBound(baCompressed) + 1, VarPtr(baBuffer(0)), UBound(baBuffer) + 1)
        PithyDecompress = baBuffer
    End If
End Function
