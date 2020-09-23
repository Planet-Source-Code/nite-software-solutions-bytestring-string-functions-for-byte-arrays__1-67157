Attribute VB_Name = "ModByteString"
'ByteString module v1.0.0
'Author: Danny Elkins (DigiRev@hotmail.com or DanGuitar3@hotmail.com)
'Created: November 17th, 2006
'Last updated: November 22nd, 2006

'Description:
'------------
'o I created this module to try and make it as easy to work with
'  byte arrays as it is with strings, while still maintaining the speed of byte arrays.

'o Feel free to use this in any of your programs, both free and commercial.

'Notes:
'------
'o I pass most parameters as ByRef to conserve memory.
'o All returned arrays are 0 based.

Option Explicit

'Our own CompareMethod enum without the vbDatabaseCompare.
Public Enum bsCompareMethod
    bsTextCompare = 0
    bsBinaryCompare = 1
End Enum

'Used for fast copying of data.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Gets the lower bounds (LBound) of a byte array without returning an error.
'I didn't want to have to put error handling into every sub/function.
Private Function SafeLBound(TheArray() As Byte) As Long
    On Error GoTo ErrorHandler
    
    SafeLBound = LBound(TheArray())
    
    Exit Function
    
ErrorHandler:
    
End Function

'Gets the upper bounds (UBound) of a byte array without returning an error.
'I didn't want to have to put error handling into every sub/function.
Private Function SafeUBound(TheArray() As Byte) As Long
    On Error GoTo ErrorHandler
    
    SafeUBound = UBound(TheArray())
    
    Exit Function
    
ErrorHandler:
    
End Function

'Converts a byte array to a string.
'It's best to just use StrConv() function to begin with instead of this function.
Public Function bsToString(TheArray() As Byte) As String
    bsToString = StrConv(TheArray(), vbUnicode)
End Function

'Converts a string to a byte array.
'It's best to just use the StrConv() function to begin with instead of this function.
Public Function bsFromString(TheString As String) As Byte()
    bsFromString = StrConv(TheString, vbFromUnicode)
End Function

'Equivalent of Right() function for strings.
'Returns a string.
Public Function bsStrRight(TheArray() As Byte, ByVal Length As Long) As String
    Dim lonLen As Long, bytRet() As Byte
    
    If Length = 0 Then Exit Function
    
    'Get the length of the array.
    lonLen = SafeUBound(TheArray())
    
    If Length > lonLen Then
        ReDim bytRet(0 To lonLen) As Byte
        
        CopyMemory bytRet(0), TheArray(0), lonLen + 1
        
        bsStrRight = StrConv(bytRet(), vbUnicode)
        
        Erase bytRet()
    Else
            
        ReDim bytRet(0 To Length - 1) As Byte
        
        CopyMemory bytRet(0), TheArray((lonLen - Length) + 1), (Length + 1)
        
        bsStrRight = StrConv(bytRet(), vbUnicode)
        
        Erase bytRet()
    End If
    
End Function

'Equivalent of Right() function for strings.
'Returns a byte array.
Public Function bsRight(TheArray() As Byte, ByVal Length As Long) As Byte()
    Dim lonLen As Long, bytRet() As Byte
    
    If Length = 0 Then Exit Function
    
    'Get the length of the array.
    lonLen = SafeUBound(TheArray())
    
    If Length > lonLen Then
        ReDim bytRet(0 To lonLen) As Byte
        
        CopyMemory bytRet(0), TheArray(0), lonLen + 1
        
        bsRight = bytRet()
        
        Erase bytRet()
    Else
            
        ReDim bytRet(0 To Length - 1) As Byte
        
        CopyMemory bytRet(0), TheArray((lonLen - Length) + 1), (Length + 1)
        
        bsRight = bytRet()
        
        Erase bytRet()
    End If
    
End Function

'Equivalent of Left() function for strings.
'Returns a string.
Public Function bsStrLeft(TheArray() As Byte, ByVal Length As Long) As String
    Dim lonLen As Long, bytRet() As Byte
    
    If Length = 0 Then Exit Function
    
    'Get the length of the array.
    lonLen = SafeUBound(TheArray())
    
    'Check if length is larger than the size of the string.
    If Length > lonLen Then
        'Allocate some memory to store the returned string.
        ReDim bytRet(0 To lonLen) As Byte
        
        'Copy that memory directly to the byte array (without having to loop).
        CopyMemory bytRet(0), TheArray(0), lonLen + 1
        
        bsStrLeft = StrConv(bytRet(), vbUnicode)
        
        'Clean up.
        Erase bytRet()
    Else
        'Allocate some memory to store the returned string.
        ReDim bytRet(0 To Length - 1) As Byte
        
        'Copy that memory directly to the byte array (without having to loop).
        CopyMemory bytRet(0), TheArray(0), Length
        
        bsStrLeft = StrConv(bytRet(), vbUnicode)
        
        'Clean up.
        Erase bytRet()
    End If
    
End Function

'Equivalent of Left() function for strings.
'Returns a byte array.
Public Function bsLeft(TheArray() As Byte, ByVal Length As Long) As Byte()
    Dim lonLen As Long, bytRet() As Byte
    
    If Length = 0 Then Exit Function
    
    'Get the length of the array.
    lonLen = SafeUBound(TheArray())
    
    'Check if length is larger than the size of the string.
    If Length > lonLen Then
        'Allocate some memory to store the returned string.
        ReDim bytRet(0 To lonLen) As Byte
        
        'Copy that memory directly to the byte array (without having to loop).
        CopyMemory bytRet(0), TheArray(0), lonLen + 1
        
        bsLeft = bytRet()
        
        'Clean up.
        Erase bytRet()
    Else
        'Allocate some memory to store the returned string.
        ReDim bytRet(0 To Length - 1) As Byte
        
        'Copy that memory directly to the byte array (without having to loop).
        CopyMemory bytRet(0), TheArray(0), Length
        
        bsLeft = bytRet()
        
        'Clean up.
        Erase bytRet()
    End If
    
End Function

'Equivalent of Mid() function for strings.
'Returns a string.
Public Function bsStrMid(TheArray() As Byte, ByVal Start As Long, Optional ByVal Length As Long = -1) As String
    Dim lonLen As Long, bytRet() As Byte
    Dim lonRetLen As Long
    
    If Length = 0 Then Exit Function
    
    lonLen = SafeUBound(TheArray())
    
    'Return nothing if start is greater than length.
    'This is how the Mid() function behaves.
    If (Start - 1) > lonLen Then Exit Function
    
    If Length = -1 Or (Length - 1) > lonLen Then
        lonRetLen = lonLen - (Start - 1)
        
        ReDim bytRet(0 To lonRetLen) As Byte
        
        CopyMemory bytRet(0), TheArray(Start - 1), lonRetLen + 1
        
        bsStrMid = StrConv(bytRet(), vbUnicode)
        
        Erase bytRet()
    Else
        ReDim bytRet(0 To Length - 1) As Byte
        
        CopyMemory bytRet(0), TheArray(Start - 1), Length
        
        bsStrMid = StrConv(bytRet(), vbUnicode)
        
        Erase bytRet()
    End If
    
End Function

'Equivalent of Mid() function for strings.
'Returns a byte array.
Public Function bsMid(TheArray() As Byte, ByVal Start As Long, Optional ByVal Length As Long = -1) As Byte()
    Dim lonLen As Long, bytRet() As Byte
    Dim lonRetLen As Long
    
    If Length = 0 Then Exit Function
    
    lonLen = SafeUBound(TheArray())
    
    'Return nothing if start is greater than length.
    'This is how the Mid() function behaves.
    If (Start - 1) > lonLen Then Exit Function
    
    If Length = -1 Or (Length - 1) > lonLen Then
        lonRetLen = lonLen - (Start - 1)
        
        ReDim bytRet(0 To lonRetLen) As Byte
        
        CopyMemory bytRet(0), TheArray(Start - 1), lonRetLen + 1
        
        bsMid = bytRet()
        
        Erase bytRet()
    Else
    
        ReDim bytRet(0 To Length - 1) As Byte
        
        CopyMemory bytRet(0), TheArray(Start - 1), Length
        
        bsMid = bytRet()
        
        Erase bytRet()
    End If
    
End Function

'Compares a byte array to a string.
'Returns TRUE if they are the same.
Private Function CompareByteToStr(TheArray() As Byte, CompareTo As String, Optional ByVal Compare As bsCompareMethod = bsBinaryCompare) As Boolean
    Dim strCompareTo As String
    
    strCompareTo = StrConv(TheArray(), vbUnicode)
    
    If Compare = bsBinaryCompare Then
        CompareByteToStr = (strCompareTo = CompareTo)
    Else
        CompareByteToStr = (LCase$(strCompareTo) = LCase$(CompareTo))
    End If
    
End Function

'Equivalent of InStrRev() function for strings.
'Compares with a string.
Public Function bsInStrRev(ByVal Start As Long, TheArray() As Byte, ToFind As String, Optional Compare As bsCompareMethod = bsBinaryCompare) As Long
    Dim lonLen As Long, lonLoop As Long
    Dim bytStartFindBin As Byte, bytStartFindStr As Byte
    Dim bytCurLC As Byte, bytCompArr() As Byte
    Dim lonLenFind As Long, lonStart As Long
    
    If Start <= 0 Then
        lonStart = SafeUBound(TheArray()) + 1
    Else
        lonStart = Start
    End If
    
    lonLen = SafeUBound(TheArray())
    lonLenFind = Len(ToFind)

    If (lonStart - 1) > lonLen Then Exit Function
    
    ReDim bytCompArr(0 To (lonLenFind - 1)) As Byte
    
    'First byte of string we want to find.
    bytStartFindBin = CByte(Asc(Left$(ToFind, 1)))
    'First byte of string we want to find (for case-insensitive search).
    bytStartFindStr = CByte(Asc(LCase$(Left$(ToFind, 1))))
    
    'lonStart looping through the byte array (backwards).
    For lonLoop = (lonStart - 1) To 0 Step -1
        
        'Already done.
        'If lonLenFind > ((lonLen - lonLoop) + 1) Then Exit For
        
        'Case-sensitive search.
        If Compare = bsBinaryCompare Then
            
            If TheArray(lonLoop) = bytStartFindBin Then
                
                CopyMemory bytCompArr(0), TheArray(lonLoop), (lonLenFind)
                
                If CompareByteToStr(bytCompArr(), ToFind, Compare) = True Then
                    bsInStrRev = lonLoop + 1
                    Exit For
                End If
                
            End If
        
        'Case-insensitive search.
        ElseIf Compare = bsTextCompare Then
            bytCurLC = CByte(Asc(LCase$(Chr$(CLng(TheArray(lonLoop))))))
            
            If bytCurLC = bytStartFindStr Then
                
                CopyMemory bytCompArr(0), TheArray(lonLoop), (lonLenFind)
                
                If CompareByteToStr(bytCompArr(), ToFind, Compare) = True Then
                    bsInStrRev = lonLoop + 1
                    Exit For
                End If
                
            End If
        
        End If
    
    Next lonLoop
    
End Function


'Equivalent of InStr() function for strings.
'Compares with a string.
Public Function bsInStrString(ByVal Start As Long, TheArray() As Byte, ToFind As String, Optional Compare As bsCompareMethod = bsBinaryCompare) As Long
    Dim lonLen As Long, lonLoop As Long
    Dim bytStartFindBin As Byte, bytStartFindStr As Byte
    Dim bytCurLC As Byte, bytCompArr() As Byte
    Dim lonLenFind As Long
    
    If Start = 0 Then
        Err.Raise 5, App.EXEName, "Invalid procedure call or argument"
        Exit Function
    End If
    
    lonLen = SafeUBound(TheArray())
    lonLenFind = Len(ToFind)
    
    'Start position is greater than length of byte array. Return 0.
    'This is how the InStr() function behaves.
    If (Start - 1) > lonLen Then Exit Function
    
    ReDim bytCompArr(0 To (lonLenFind - 1)) As Byte
    
    'First byte of string we want to find.
    bytStartFindBin = CByte(Asc(Left$(ToFind, 1)))
    'First byte of string we want to find (for case-insensitive search).
    bytStartFindStr = CByte(Asc(LCase$(Left$(ToFind, 1))))
    
    'Start looping through the byte array.
    For lonLoop = (Start - 1) To lonLen
        
        'Already done.
        If lonLenFind > ((lonLen - lonLoop) + 1) Then Exit For
        
        'Case-sensitive search.
        If Compare = bsBinaryCompare Then
            
            If TheArray(lonLoop) = bytStartFindBin Then
                
                CopyMemory bytCompArr(0), TheArray(lonLoop), (lonLenFind)
                
                If CompareByteToStr(bytCompArr(), ToFind, Compare) = True Then
                    bsInStrString = lonLoop + 1
                    Exit For
                End If
                
            End If
        
        'Case-insensitive search.
        ElseIf Compare = bsTextCompare Then
            bytCurLC = CByte(Asc(LCase$(Chr$(CLng(TheArray(lonLoop))))))
            
            If bytCurLC = bytStartFindStr Then
                
                CopyMemory bytCompArr(0), TheArray(lonLoop), (lonLenFind)
                
                If CompareByteToStr(bytCompArr(), ToFind, Compare) = True Then
                    bsInStrString = lonLoop + 1
                    Exit For
                End If
                
            End If
        
        End If
    
    Next lonLoop
    
End Function

'Appends some data tot he end of the byte.
'Appends a string.
Public Sub bsAppendToByte(TheArray() As Byte, ToAppend As String)
    Dim lonLen As Long, lonStart As Long
    Dim bytAppend() As Byte
    
    lonLen = SafeUBound(TheArray())
    lonStart = SafeLBound(TheArray())
    bytAppend() = StrConv(ToAppend, vbFromUnicode)
    
    ReDim Preserve TheArray(lonStart To (lonLen + Len(ToAppend))) As Byte
    
    CopyMemory TheArray(lonLen + 1), bytAppend(0), Len(ToAppend)
End Sub

'Equivalent to the IsNumeric() function for strings.
'Except it does not accept D and E as a numeric value.
Public Function bsIsNumeric(TheArray() As Byte) As Boolean
    Dim lonLoop As Long, lonBnd As Long
    Dim lonStart As Long, bolFound As Boolean
    
    lonStart = SafeLBound(TheArray())
    lonBnd = SafeUBound(TheArray())
    
    For lonLoop = lonStart To lonBnd
        
        Select Case TheArray(lonLoop)
        '48 49 50 51 52 53 54 55 56 57
            Case 48 To 57 'Case Not 48 To 57 returned overflow error? :/
                '
            Case Else
                bolFound = True
                Exit For
        
        End Select
    
    Next lonLoop
    
    bsIsNumeric = Not bolFound
End Function
