VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' By Chris Lucas
' Updated 1-14-02:  Made stable against all ANSI Input except ligatures (Æ,æ,Œ,œ,ß,Þ, and þ)
'
' ************************************** WARNING ***************************************
' If your strings have ligatures in them and you need vbTextCompare, this code will fail
' use Visual Basic's InStr instead.
' **************************************************************************************
' If your doing binary comparision this is 100% stable against all input(ANSI and UNICODE)

Option Explicit

Private Declare Function ArrPtr& Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any)
Private Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal nBytes&)

Private Header1(5) As Long
Private Header2(5) As Long
Private SafeArray1() As Integer
Private SafeArray2() As Integer
Private LUT(8482) As Long

Private Sub Class_Initialize()
    Dim i As Long
    
    ' Set up our template for looking at strings
    Header1(0) = 1              ' Number of dimensions
    Header1(1) = 2              ' Bytes per element (integer = 2)
    Header1(4) = &H7FFFFFFF     ' Array size, 2.1+ billion should cover us
  
    ' Force SafeArray1 to use Header1 as its own header
    RtlMoveMemory ByVal ArrPtr(SafeArray1), VarPtr(Header1(0)), 4
    
    ' Set up our template for look at search text
    Header2(0) = 1                 ' Number of dimensions
    Header2(1) = 2                 ' Bytes per element (integer = 2)
    Header2(4) = &H7FFFFFFF        ' Array size, 2.1+ billion should cover us
  
    ' Force SafeArray2 to use Header2 as its own header
    RtlMoveMemory ByVal ArrPtr(SafeArray2), VarPtr(Header2(0)), 4
    
    ' Set up a look up table for ANSI characters that have a non-zero UNICODE component
    For i = 0 To 255: LUT(i) = i: Next i
    LUT(8364) = 128
    LUT(8218) = 130
    LUT(402) = 131
    LUT(8222) = 132
    LUT(8230) = 133
    LUT(8224) = 134
    LUT(8225) = 135
    LUT(710) = 136
    LUT(8240) = 137
    LUT(352) = 138
    LUT(8249) = 139
    LUT(338) = 140
    LUT(381) = 142
    LUT(8216) = 145
    LUT(8217) = 146
    LUT(8220) = 147
    LUT(8221) = 148
    LUT(8226) = 149
    LUT(8211) = 150
    LUT(8212) = 151
    LUT(732) = 152
    LUT(8482) = 153
    LUT(353) = 154
    LUT(8250) = 155
    LUT(339) = 156
    LUT(382) = 158
    LUT(376) = 159

End Sub

Private Sub Class_Terminate()
    ' Make SafeArray1 and SafeArray2 once again use
    ' their own headers
    
    ' If this code doesn't run the IDE will crash
    RtlMoveMemory ByVal ArrPtr(SafeArray1), 0&, 4
    RtlMoveMemory ByVal ArrPtr(SafeArray2), 0&, 4
End Sub

Friend Function InString(Start As Long, String1 As String, String2 As String, _
 Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
 
    Static DT&(255), OldPtr&, PatLen&, PatLen2&
    Dim i&, j&, tmp&, alt&, TextLen&
    
    ' Prepare headers
    Header1(3) = StrPtr(String1): Header2(3) = StrPtr(String2)
    
    ' Test for special cases
    TextLen = Len(String1): If TextLen = 0 Then Exit Function
    
    ' Distance Table Setup - only go through this if we are searching for a new string
    If Header2(3) <> OldPtr Then
        ' Cache the new pointer for next time
        OldPtr = Header2(3)
        
        PatLen = Len(String2): PatLen2 = PatLen - 1
        If PatLen = 1 And Compare = vbBinaryCompare Then
            ' All bow to the master, I can't beat native code
            ' on this particular call
            InString = InStr(Start, String1, String2, Compare)
            Exit Function
        End If
        
        ' Now setup the distance table
        For i = 0 To 255: DT(i) = PatLen: Next i
    
        If Compare = vbBinaryCompare Then
            ' Binary compare, easy enough
            For i = 0 To PatLen - 1: DT(LUT(SafeArray2(i))) = PatLen - i - 1: Next i
        Else
            ' Text compare, we need to calculate offsets for
            ' case matches as well
            For i = 0 To PatLen2
                tmp = LUT(SafeArray2(i))
                
                DT(tmp) = PatLen2 - i
                Select Case tmp
                    Case 97& To 122&: alt = tmp - 32
                    Case 65& To 90&: alt = tmp + 32
                    
                    Case 138&: alt = 154
                    Case 140&: alt = 156
                    Case 142&: alt = 158
                    Case 154&: alt = 138
                    Case 156&: alt = 140
                    Case 158&: alt = 142
                    Case 159&: alt = 255
                    Case 192& To 214&: alt = tmp + 32
                    Case 216& To 222&: alt = tmp + 32
                    Case 224& To 246&: alt = tmp - 32
                    Case 248& To 254&: alt = tmp - 32
                    Case 255&: alt = 159

                    Case Else: GoTo NoAlt
                End Select
                DT(alt) = PatLen2 - i
NoAlt:      Next i
        End If
    End If
    
    ' Now we start making comparisions
    i = PatLen + Start - 2
    If i > TextLen - 1 Then GoTo Bail
    
    Do
Retry:  If DT(LUT(SafeArray1(i))) = 0 Then Exit Do Else i = i + DT(LUT(SafeArray1(i)))
        If i > TextLen - 1 Then GoTo Bail
    Loop
    
    For j = 1 To PatLen2
        ' ensure that we still match
        If Not (SafeArray1(i - j) And &HDF&) = (SafeArray2(PatLen2 - j) And &HDF&) Then
            ' matching failed
            i = i + DT(LUT(SafeArray1(i))) + 1
            ' make sure we aren't beyond the end of our string
            If i > TextLen - 1 Then GoTo Bail Else GoTo Retry
        End If
    Next j
    
    ' We can only get here if we've matched the entire string
    InString = i - PatLen + 2
    Exit Function
Bail:
    InString = 0
End Function

