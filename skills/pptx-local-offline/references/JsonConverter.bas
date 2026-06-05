Attribute VB_Name = "JsonConverter"
Option Explicit

Private jsonText As String
Private jsonIndex As Long

Public Function ParseJson(ByVal text As String) As Object
    jsonText = text
    jsonIndex = 1
    SkipWhitespace
    Set ParseJson = ParseObject()
End Function

Public Function ConvertToJson(ByVal value As Variant) As String
    ConvertToJson = SerializeValue(value)
End Function

Private Function ParseValue() As Variant
    SkipWhitespace
    Select Case CurrentChar()
        Case "{": Set ParseValue = ParseObject()
        Case "[": Set ParseValue = ParseArray()
        Case """": ParseValue = ParseString()
        Case "t": ConsumeLiteral "true": ParseValue = True
        Case "f": ConsumeLiteral "false": ParseValue = False
        Case "n": ConsumeLiteral "null": ParseValue = Null
        Case Else: ParseValue = ParseNumber()
    End Select
End Function

Private Function ParseObject() As Object
    Dim result As Object
    Dim key As String
    Dim value As Variant

    Set result = CreateObject("Scripting.Dictionary")
    ExpectChar "{"
    SkipWhitespace
    If CurrentChar() = "}" Then
        jsonIndex = jsonIndex + 1
        Set ParseObject = result
        Exit Function
    End If

    Do
        key = ParseString()
        SkipWhitespace
        ExpectChar ":"
        SkipWhitespace
        If CurrentChar() = "{" Or CurrentChar() = "[" Then
            Set value = ParseValue()
        Else
            value = ParseValue()
        End If
        AssignDictionaryValue result, key, value
        SkipWhitespace
        If CurrentChar() = "}" Then Exit Do
        ExpectChar ","
    Loop

    ExpectChar "}"
    Set ParseObject = result
End Function

Private Function ParseArray() As Collection
    Dim result As New Collection
    Dim value As Variant

    ExpectChar "["
    SkipWhitespace
    If CurrentChar() = "]" Then
        jsonIndex = jsonIndex + 1
        Set ParseArray = result
        Exit Function
    End If

    Do
        SkipWhitespace
        If CurrentChar() = "{" Or CurrentChar() = "[" Then
            Set value = ParseValue()
            result.Add value
        Else
            value = ParseValue()
            result.Add value
        End If
        SkipWhitespace
        If CurrentChar() = "]" Then Exit Do
        ExpectChar ","
    Loop

    ExpectChar "]"
    Set ParseArray = result
End Function

Private Function ParseString() As String
    Dim ch As String
    Dim result As String

    ExpectChar """"
    Do While jsonIndex <= Len(jsonText)
        ch = Mid$(jsonText, jsonIndex, 1)
        jsonIndex = jsonIndex + 1
        If ch = """" Then Exit Do
        If ch = "\" Then
            ch = Mid$(jsonText, jsonIndex, 1)
            jsonIndex = jsonIndex + 1
            Select Case ch
                Case """", "\", "/": result = result & ch
                Case "b": result = result & vbBack
                Case "f": result = result & vbFormFeed
                Case "n": result = result & vbLf
                Case "r": result = result & vbCr
                Case "t": result = result & vbTab
                Case "u"
                    result = result & ChrW$(CLng("&H" & Mid$(jsonText, jsonIndex, 4)))
                    jsonIndex = jsonIndex + 4
                Case Else: Err.Raise vbObjectError + 3001, "JsonConverter", "Unsupported escape sequence: \" & ch
            End Select
        Else
            result = result & ch
        End If
    Loop
    ParseString = result
End Function

Private Function ParseNumber() As Variant
    Dim startPos As Long
    Dim token As String

    startPos = jsonIndex
    Do While jsonIndex <= Len(jsonText)
        If InStr("-+0123456789.eE", Mid$(jsonText, jsonIndex, 1)) = 0 Then Exit Do
        jsonIndex = jsonIndex + 1
    Loop
    token = Mid$(jsonText, startPos, jsonIndex - startPos)
    If InStr(token, ".") > 0 Or InStr(token, "e") > 0 Or InStr(token, "E") > 0 Then
        ParseNumber = CDbl(token)
    Else
        ParseNumber = CLng(token)
    End If
End Function

Private Function SerializeValue(ByVal value As Variant) As String
    If IsObject(value) Then
        Select Case TypeName(value)
            Case "Dictionary"
                SerializeValue = SerializeDictionary(value)
            Case "Collection"
                SerializeValue = SerializeCollection(value)
            Case Else
                Err.Raise vbObjectError + 3002, "JsonConverter", "Unsupported object type: " & TypeName(value)
        End Select
        Exit Function
    End If

    If IsNull(value) Then
        SerializeValue = "null"
    ElseIf VarType(value) = vbString Then
        SerializeValue = """" & EscapeJson(CStr(value)) & """"
    ElseIf VarType(value) = vbBoolean Then
        SerializeValue = LCase$(CStr(CBool(value)))
    ElseIf IsNumeric(value) Then
        SerializeValue = Replace(CStr(value), ",", ".")
    Else
        SerializeValue = """" & EscapeJson(CStr(value)) & """"
    End If
End Function

Private Function SerializeDictionary(ByVal dict As Object) As String
    Dim key As Variant
    Dim parts As Collection
    Dim itemText As String

    Set parts = New Collection
    For Each key In dict.Keys
        itemText = """" & EscapeJson(CStr(key)) & """:" & SerializeValue(dict(key))
        parts.Add itemText
    Next key
    SerializeDictionary = "{" & JoinCollection(parts, ",") & "}"
End Function

Private Function SerializeCollection(ByVal col As Collection) As String
    Dim item As Variant
    Dim parts As Collection

    Set parts = New Collection
    For Each item In col
        If IsObject(item) Then
            parts.Add SerializeValue(item)
        Else
            parts.Add SerializeValue(item)
        End If
    Next item
    SerializeCollection = "[" & JoinCollection(parts, ",") & "]"
End Function

Private Function EscapeJson(ByVal text As String) As String
    Dim i As Long, ch As String, code As Long, result As String
    result = ""
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        code = AscW(ch)
        Select Case code
        Case 92:  result = result & "\\"
        Case 34:  result = result & "\"""
        Case 8:   result = result & "\b"
        Case 9:   result = result & "\t"
        Case 10:  result = result & "\n"
        Case 12:  result = result & "\f"
        Case 13:  result = result & "\r"
        Case Else
            If code >= 32 Then
                result = result & ch
            Else
                result = result & "\u" & Right$("0000" & Hex$(code), 4)
            End If
        End Select
    Next i
    EscapeJson = result
End Function

Private Sub AssignDictionaryValue(ByVal dict As Object, ByVal key As String, ByVal value As Variant)
    If IsObject(value) Then
        Set dict(key) = value
    Else
        dict(key) = value
    End If
End Sub

Private Function JoinCollection(ByVal values As Collection, ByVal separator As String) As String
    Dim i As Long
    For i = 1 To values.Count
        If i > 1 Then JoinCollection = JoinCollection & separator
        JoinCollection = JoinCollection & CStr(values(i))
    Next i
End Function

Private Sub SkipWhitespace()
    Do While jsonIndex <= Len(jsonText)
        Select Case Mid$(jsonText, jsonIndex, 1)
            Case " ", vbTab, vbCr, vbLf
                jsonIndex = jsonIndex + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

Private Function CurrentChar() As String
    If jsonIndex > Len(jsonText) Then
        CurrentChar = ""
    Else
        CurrentChar = Mid$(jsonText, jsonIndex, 1)
    End If
End Function

Private Sub ExpectChar(ByVal expected As String)
    SkipWhitespace
    If CurrentChar() <> expected Then
        Err.Raise vbObjectError + 3003, "JsonConverter", "Expected '" & expected & "' at position " & jsonIndex
    End If
    jsonIndex = jsonIndex + 1
End Sub

Private Sub ConsumeLiteral(ByVal literal As String)
    If Mid$(jsonText, jsonIndex, Len(literal)) <> literal Then
        Err.Raise vbObjectError + 3004, "JsonConverter", "Expected literal '" & literal & "' at position " & jsonIndex
    End If
    jsonIndex = jsonIndex + Len(literal)
End Sub