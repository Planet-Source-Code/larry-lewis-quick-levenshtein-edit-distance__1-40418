<div align="center">

## Quick Levenshtein Edit Distance


</div>

### Description

Levenshtein edit distance is a measure of the similarity between two strings. Edit distance is the the minimum number of character deletions, insertions, substitutions, and transpositions required to transform string1 into string2. In essence, the function is used to perform a fuzzy or approximate string search. This is very handy for trying to find the "correct" string for one that has been entered incorrectly, mistyped, etc. The code has been optimized to find strings that are very similar. A "limit" parameter is provided so the function will quickly reject strings that contain more than k mismatches.
 
### More Info
 
s as string, t as string, limit as integer 'maximum edit distance

This code takes character transpositions into account when calculating edit distance (If desired, the transposition code can be commented out). The limit parameter provides a significant performance gain (over 10x faster) over standard implementations when searching for highly similar strings.

Returns the integer edit distance (minimum number of character deletions, insertions, substitutions, and transpositions) required to transform string1 into string2 where edit distance is <= limit. Otherwise returns (len(s) + len(t)).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Larry Lewis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/larry-lewis.md)
**Level**          |Advanced
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/larry-lewis-quick-levenshtein-edit-distance__1-40418/archive/master.zip)





### Source Code

```
Public Function LD(ByVal s As String, ByVal t As String, ByVal limit As Long) As Long
' Levenshtein edit distance is a measure of the similarity between two strings.
' Edit distance is the the minimum number of character deletions, insertions,
' substitutions, and transpositions required to transform string1 into string2.
' Includes transposed characters and many times faster than the original
' LEVENSHTEIN function (especially when limit is low)
' Returns EditDistance where EditDistance is <= limit otherwise returns len(@s) + len(@t)
' Based on code by: Michael Gilleland http://www.merriampark.com/ld.htm
' Author: Larry Lewis
Dim d() As Integer ' matrix
Dim k As Integer
Dim m As Integer ' length of t
Dim n As Integer ' length of s
Dim i As Integer ' iterates through s
Dim j As Integer ' iterates through t
Dim s_i As String ' ith character of s
Dim t_j As String ' jth character of t
Dim lendif As Integer
Dim MinDist As Integer, y As Integer, z As Integer
Dim smallLen As Integer
LD = Len(s$) + Len(t$)
'Remove leftmost matching portion of strings
While Left$(s, 1) = Left$(t, 1) And Len(s) > 0 And Len(t) > 0
 s = Right$(s, Len(s) - 1)
 t = Right$(t, Len(t) - 1)
Wend
 ' Find shorter string
 n = Len(s)
 m = Len(t)
 smallLen = n
 If m < n Then smallLen = m
 'Stop if string lengths differ by more than LIMIT
 lendif = n - m
 If Abs(lendif) > limit Then
  Exit Function
 End If
 If n = 0 Then
 If m <= limit Then LD = m
 Exit Function
 End If
 If m = 0 Then
 If n <= limit Then LD = n
 Exit Function
 End If
 ReDim d(0 To n, 0 To m) As Integer
 ' Initialize matrix
 For i = 0 To n
 d(i, 0) = i
 Next i
 For j = 1 To m
 d(0, j) = j
 Next j
 ' Main Loop - Levenshtein
 ' Try to traverse the matrix so that pruning cells are calculated ASAP
 For k = 1 To smallLen
  s_i = Mid$(s, k, 1)
  For j = 1 To k
   t_j = Mid$(t, j, 1)
   ' Evaluate cell
   MinDist = d(k - 1, j - 1)
   If s_i <> t_j Then MinDist = MinDist + 1
   y = d(k - 1, j) + 1
   z = d(k, j - 1) + 1
   If y < MinDist Then MinDist = y
   If z < MinDist Then MinDist = z
   d(k, j) = MinDist
   ' Check Transposition
   If j > 1 Then
   If k > 1 Then
    If MinDist - d(k - 2, j - 2) = 2 Then
     If Mid$(s, k - 1, 1) = t_j Then
      If s_i = Mid$(t, j - 1, 1) Then
      d(k, j) = d(k - 2, j - 2) + 1
      End If
     End If
    End If
   End If
   End If
   ' Limit Pruning - Stop processing if we are already over the limit
   If k = j + lendif Then
   If k < smallLen Then
    If d(k, j) > limit Then
     Erase d
     Exit Function
    End If
   ElseIf j < smallLen Then
    If d(k, j) > limit Then
     Erase d
     Exit Function
    End If
   End If
   End If
  Next j
  If j <= m Then
  t_j = Mid$(t, j, 1)
  For i = 1 To k
   s_i = Mid$(s, i, 1)
   'Evaluate
   MinDist = d(i - 1, j - 1)
   If s_i <> t_j Then MinDist = MinDist + 1
   y = d(i - 1, j) + 1
   z = d(i, j - 1) + 1
   If y < MinDist Then MinDist = y
   If z < MinDist Then MinDist = z
   d(i, j) = MinDist
   ' Check Transposition
   If i > 1 Then
    If MinDist - d(i - 2, j - 2) = 2 Then
     If Mid$(s, i - 1, 1) = t_j Then
      If s_i = Mid$(t, j - 1, 1) Then
      d(i, j) = d(i - 2, j - 2) + 1
      End If
     End If
    End If
   End If
   ' Limit Pruning - Stop processing if we are already over the limit
   If i = j + lendif Then
   If i < smallLen Then
    If d(i, j) > limit Then
     Erase d
     Exit Function
    End If
   ElseIf j < smallLen Then
    If d(i, j) > limit Then
     Erase d
     Exit Function
    End If
   End If
   End If
  Next i
  End If
 Next k
 'process remaining rightmost portion of matrix if any
 For i = n + 1 To m
 t_j = Mid$(t, i, 1)
 For j = 1 To n
  s_i = Mid$(s, j, 1)
  ' Evaluate
  MinDist = d(j - 1, i - 1)
  If s_i <> t_j Then MinDist = MinDist + 1
  y = d(j - 1, i) + 1
  z = d(j, i - 1) + 1
  If y < MinDist Then MinDist = y
  If z < MinDist Then MinDist = z
  d(j, i) = MinDist
  ' Check For Transposition
   If j > 1 Then
    If MinDist - d(j - 2, i - 2) = 2 Then
     If Mid$(s, j - 1, 1) = t_j Then
     If s_i = Mid$(t, i - 1, 1) Then
      d(j, i) = d(j - 2, i - 2) + 1
     End If
     End If
    End If
   End If
  ' Limit Pruning - Stop processing if we are already over the limit
  If j = i + lendif Then
   If (j < n Or i < m) Then
   If d(j, i) > limit Then
    Erase d
    Exit Function
   End If
   End If
  End If
 Next j
 Next i
 'process remaining lower portion of matrix if any
 For i = m + 1 To n
 s_i = Mid$(s, i, 1)
 For j = 1 To m
  t_j = Mid$(t, j, 1)
  'Evaluate
  MinDist = d(i - 1, j - 1)
  If s_i <> t_j Then MinDist = MinDist + 1
  y = d(i - 1, j) + 1
  z = d(i, j - 1) + 1
  If y < MinDist Then MinDist = y
  If z < MinDist Then MinDist = z
  d(i, j) = MinDist
  ' Check for Transposition
   If j > 1 Then
   If MinDist - d(i - 2, j - 2) = 2 Then
    If Mid$(s, i - 1, 1) = t_j Then
     If s_i = Mid$(t, j - 1, 1) Then
      d(i, j) = d(i - 2, j - 2) + 1
     End If
    End If
   End If
   End If
  ' Limit Pruning - Stop processing if we are already over the limit
  If i = j + lendif Then
   If (i < n Or j < m) Then
   If d(i, j) > limit Then
    Erase d
    Exit Function
   End If
   End If
  End If
 Next j
 Next i
 LD = d(n, m)
 Erase d
 Exit Function
dump:
 Dim ss As String
 For i = 1 To n
 For j = 1 To m
 ss = ss & d(i, j) & " "
 Next j
 Debug.Print ss
 ss = ""
 Next i
End Function
```

