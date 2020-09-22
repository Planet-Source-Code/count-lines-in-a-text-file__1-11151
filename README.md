<div align="center">

## Count Lines in a Text File


</div>

### Description

This code will count all of the lines in a text File. This code will work on any sized file, and is quicker than using LineInput or the FSO.
 
### More Info
 
A Valid FileName

After being asked many times how you count the number of lines in a text file, i decided to write a function that would do it as quickly as possible.

This function reads the file in in chunks. When it reads a chunk in, it counts the number of lines in that chunk, before reading the next chunk, and so on.

You may find a different chunk size works better for you, feel free to experiment with it.

The number of lines contained in the passed filename.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Advanced
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/count-lines-in-a-text-file__1-11151/archive/master.zip)





### Source Code

```
Function lineCount(myInFile As String) As Long
 Dim lFileSize As Long, lChunk As Long
 Dim bFile() As Byte
 Dim lSize As Long
 Dim strText As String
 'the size of the chunk to read in. You can experiment
 'with this to see what works fastest.
 lSize = CLng(1024) * 10
 'size the array to the chunk size
 ReDim bFile(lSize - 1) As Byte
 Open myInFile For Binary As #1
 'get the file size
 lFileSize = LOF(1)
 'set the chunk number to 1
 lChunk = 1
 Do While (lSize * lChunk) < lFileSize
  'get the data from the in file
  Get #1, , bFile
  strText = StrConv(bFile, vbUnicode)
  'get the line count for this chunk
  lineCount = lineCount + searchText(strText)
  'increment the chunk count
  lChunk = lChunk + 1
 Loop
 'redim the array to the remaining size
 ReDim bFile((lFileSize - (lSize * (lChunk - 1))) - 1) As Byte
 'get the remaining data
 Get #1, , bFile
 strText = StrConv(bFile, vbUnicode)
 'get line count for this chunk
 lineCount = lineCount + searchText(strText)
 'close the file
 Close #1
 lineCount = lineCount + 1
End Function
Private Function searchText(strText As String) As Long
 Static blPossible As Boolean
 Dim lp1 As Long
 'if we have a possible line count
 If blPossible = True Then
  'if the fist charcter is chr(10) then we have a new line
  If Left$(strText, 1) = Chr(10) Then
  searchText = searchText + 1
  End If
 End If
 blPossible = False
 'loop through counting vbCrLf's
 lp1 = 1
 Do
  lp1 = InStr(lp1, strText, vbCrLf)
  If lp1 <> 0 Then
  searchText = searchText + 1
  lp1 = lp1 + 2
  End If
 Loop Until lp1 = 0
 'if the last character is a chr(13) then we may have a
 'new line, so we mark it as possible
 If Right$(strText, 1) = Chr(13) Then
  blPossible = True
 End If
End Function
```

