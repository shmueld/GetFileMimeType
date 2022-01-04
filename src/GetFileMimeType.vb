Imports System.IO
Imports Microsoft.VisualBasic

Public Class GetFilesMime
       Private Shared Function CompereArrays(arr1 As Byte(), arr2 As Byte()) As Boolean
              If (arr1 Is Nothing OrElse arr2 Is Nothing) Then
                     Return False
              End If
              If (arr1.Length <> arr2.Length) Then
                     Return False
              End If

              For i As Integer = 0 To arr1.Length - 1
                     If (arr1(i) <> arr2(i)) Then
                           Return False
                     End If
              Next i
              Return True

       End Function

       Private Shared Function GetBytesFromArray(ByRef arr As Byte(), ByVal byteNo As Integer) As Byte()
              Dim destArr As Byte() = New Byte(byteNo - 1) {}

              If (arr.Length <= byteNo) Then
                     Return destArr
              End If

              Array.Copy(arr, 0, destArr, 0, byteNo)

              Return destArr
       End Function

       Public Shared Function GetMineType(FS As Stream, Filename As String, extension As String) As String
              Dim br As BinaryReader = New BinaryReader(FS)
              Dim bytes As Byte() = br.ReadBytes(FS.Length)
              Return GetMineType(bytes, Filename, extension)
       End Function
       Public Shared Function GetMineType(file As Byte(), fileName As String, extension As String) As String
              Dim output As String = "application/octet-stream" 'Default UNKNOWN MIME TYPE

              'Ensure that the filename isn't empty or null
              If (String.IsNullOrEmpty(fileName)) Then
                     Return output
              End If

              'Get the MIME Type
              If (CompereArrays(GetBytesFromArray(file, 2), MimeTypes.BMP)) Then
                     Return "image/bmp"
              ElseIf (CompereArrays(GetBytesFromArray(file, 8), MimeTypes.DOC)) Then
                     Return "application/msword"
              ElseIf (CompereArrays(GetBytesFromArray(file, 2), MimeTypes.EXE_DLL)) Then
                     Return "application/x-msdownload" 'both use same mime type
              ElseIf (CompereArrays(GetBytesFromArray(file, 4), MimeTypes.GIF)) Then
                     Return "image/gif"
              ElseIf (CompereArrays(GetBytesFromArray(file, 4), MimeTypes.ICO)) Then
                     Return "image/x-icon"
              ElseIf (CompereArrays(GetBytesFromArray(file, 3), MimeTypes.JPG)) Then
                     Return "image/jpeg"
              ElseIf (CompereArrays(GetBytesFromArray(file, 3), MimeTypes.MP3)) Then
                     Return "audio/mpeg"
              ElseIf (CompereArrays(GetBytesFromArray(file, 14), MimeTypes.OGG)) Then
                     If (extension.ToUpper = ".OGX") Then
                           Return "application/ogg"
                     ElseIf (extension.ToUpper = ".OGA") Then
                           Return "audio/ogg"
                     Else
                           Return "video/ogg"
                     End If
              ElseIf (CompereArrays(GetBytesFromArray(file, 7), MimeTypes.PDF)) Then
                     Return "application/pdf"
              ElseIf (CompereArrays(GetBytesFromArray(file, 16), MimeTypes.PNG)) Then
                     Return "image/png"
              ElseIf (CompereArrays(GetBytesFromArray(file, 7), MimeTypes.RAR)) Then
                     Return "application/x-rar-compressed"
              ElseIf (CompereArrays(GetBytesFromArray(file, 3), MimeTypes.SWF)) Then
                     Return "application/x-shockwave-flash"
              ElseIf (CompereArrays(GetBytesFromArray(file, 4), MimeTypes.TIFF)) Then
                     Return "image/tiff"
              ElseIf (CompereArrays(GetBytesFromArray(file, 11), MimeTypes.TORRENT)) Then
                     Return "application/x-bittorrent"
              ElseIf (CompereArrays(GetBytesFromArray(file, 5), MimeTypes.TTF)) Then
                     Return "application/x-font-ttf"
              ElseIf (CompereArrays(GetBytesFromArray(file, 4), MimeTypes.WAV_AVI)) Then
                     If (extension.ToUpper = ".AVI") Then
                           Return "video/x-msvideo"
                     Else
                           Return "audio/x-wav"
                     End If
              ElseIf (CompereArrays(GetBytesFromArray(file, 16), MimeTypes.WMV_WMA)) Then
                     If (extension.ToUpper = ".WMA") Then
                           Return "audio/x-ms-wma"
                     Else
                           Return "video/x-ms-wmv"
                     End If
              ElseIf (CompereArrays(GetBytesFromArray(file, 4), MimeTypes.ZIP_DOCX)) Then
                     If (extension.ToUpper = ".DOCX") Then
                           output = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                     ElseIf (extension.ToUpper = ".XLSX") Then
                           output = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                     Else
                           output = "application/x-zip-compressed"
                     End If
              End If
              Return output
       End Function
End Class
