Attribute VB_Name = "mdPithy"
Option Explicit

Private Declare Function pithy_Compress Lib "debug_pithy.dll" (ByVal uncompressed As OLE_HANDLE, ByVal uncompressedLength As Long, ByVal compressedOut As OLE_HANDLE, ByVal compressedOutLength As Long, ByVal compressionLevel As Long) As Long
Private Declare Function pithy_Decompress Lib "debug_pithy.dll" (ByVal compressed As OLE_HANDLE, ByVal compressedLength As Long, ByVal decompressedOut As OLE_HANDLE, ByVal decompressedOutLength As Long) As Long
Private Declare Function pithy_MaxCompressedLength Lib "debug_pithy.dll" (ByVal inputLength As Long) As Long
Private Declare Function pithy_GetDecompressedLength Lib "debug_pithy.dll" (ByVal compressed As OLE_HANDLE, ByVal compressedLength As Long, decompressedOutLengthResult As Long) As Long

Public Function vbpithy_Compress(ByVal uncompressed As OLE_HANDLE, ByVal uncompressedLength As Long, ByVal compressedOut As OLE_HANDLE, ByVal compressedOutLength As Long, ByVal compressionLevel As Long) As Long
    vbpithy_Compress = pithy_Compress(uncompressed, uncompressedLength, compressedOut, compressedOutLength, compressionLevel)
End Function

Public Function vbpithy_Decompress(ByVal compressed As OLE_HANDLE, ByVal compressedLength As Long, ByVal decompressedOut As OLE_HANDLE, ByVal decompressedOutLength As Long) As Long
    vbpithy_Decompress = pithy_Decompress(compressed, compressedLength, decompressedOut, decompressedOutLength)
End Function

Public Function vbpithy_MaxCompressedLength(ByVal inputLength As Long) As Long
    vbpithy_MaxCompressedLength = pithy_MaxCompressedLength(inputLength)
End Function

Public Function vbpithy_GetDecompressedLength(ByVal compressed As OLE_HANDLE, ByVal compressedLength As Long, decompressedOutLengthResult As Long) As Long
    vbpithy_GetDecompressedLength = pithy_GetDecompressedLength(compressed, compressedLength, decompressedOutLengthResult)
End Function

