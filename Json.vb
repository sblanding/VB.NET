Option Explicit On 
Option Strict On

Imports System.Collections
Imports System.Globalization
Imports System.Text

''' <summary>
''' This class decodes JSON strings.
''' Specification details at http://www.json.org/
'''
''' JSON uses Objects and Arrays. These are represented by HashTables and ArrayLists.
''' </summary>
Public Class Json

    Private Sub New()

    End Sub

    Private Enum Token
        None = 0
        CurlyOpen = 1
        CurlyClose = 2
        SquareOpen = 3
        SquareClose = 4
        Colon = 5
        Comma = 6
        [String] = 7
        Number = 8
        [Boolean] = 9
        Null = 10
    End Enum

    ''' <summary>
    ''' Parses the JSON.
    ''' </summary>
    ''' <param name="json">A JSON string.</param>
    ''' <returns>HashTable -> ArrayList -> Object</returns>
    Public Shared Function Parse(ByVal json As String) As Object
        If (json Is Nothing OrElse json.Trim() = String.Empty) Then
            Return Nothing
        Else
            Dim charArray As Char() = json.ToCharArray()

            Return ParseValue(charArray, 0)
        End If
    End Function

    Private Shared Function ParseValue(ByVal c As Char(), ByRef index As Integer) As Object
        Dim val As Integer = GetToken(c, index)

        Select Case (val)
            Case Token.CurlyOpen
                index += 1

                Return ParseObject(c, index)
            Case Token.SquareOpen
                index += 1

                Return ParseArray(c, index)
            Case Token.String
                Return ParseString(c, index)
            Case Token.Number
                Return ParseNumber(c, index)
            Case Token.Boolean
                Return ParseBoolean(c, index)
            Case Token.Null
                Return Nothing
            Case Else
                Exit Select
        End Select

        Return Nothing
    End Function

    ' {
    Private Shared Function ParseObject(ByVal c As Char(), ByRef index As Integer) As Hashtable
        Dim table As Hashtable = New Hashtable

        While (index < c.Length)
            Dim val As Integer = GetToken(c, index)

            Select Case (val)
                Case Token.None
                    Return Nothing
                Case Token.Comma
                    index += 1
                Case Token.CurlyClose
                    index += 1

                    Return table
                Case Else
                    Dim name As String = String.Empty
                    Dim value As Object = Nothing

                    ' name
                    name = ParseString(c, index)

                    ' :
                    GetToken(c, index)

                    If (c(index) = ":") Then
                        index += 1
                    Else
                        Return Nothing
                    End If

                    ' value
                    value = ParseValue(c, index)

                    ' Add the name value pair to the hash table
                    table.Add(name, value)
            End Select
        End While

        Return table
    End Function

    ' [
    Private Shared Function ParseArray(ByVal c As Char(), ByRef index As Integer) As ArrayList
        Dim arr As ArrayList = New ArrayList

        While (index < c.Length)
            Dim val As Integer = GetToken(c, index)

            Select Case (val)
                Case Token.None
                    Return Nothing
                Case Token.Comma
                    index += 1
                Case Token.SquareClose
                    index += 1

                    Return arr
                Case Else
                    Dim value As Object = ParseValue(c, index)

                    arr.Add(value)
            End Select
        End While

        Return arr
    End Function

    ' "
    Private Shared Function ParseString(ByVal c As Char(), ByRef index As Integer) As String
        Dim s As StringBuilder = New StringBuilder

        ' Begin quote
        index += 1

        While (c(index) <> """")
            If (c(index) = "\") Then
                Dim str As String = String.Empty

                index += 1
                str = ParseEscapeChar(c, index)
                s.Append(str)
            Else
                s.Append(c(index))
                index += 1
            End If
        End While

        ' End quote
        index += 1

        Return s.ToString()
    End Function

    ' \", \\, \/, \b, \f, \n, \r, \t, \u
    Private Shared Function ParseEscapeChar(ByVal c As Char(), ByRef index As Integer) As String
        Dim str As String = c(index).ToString()

        Select Case (str)
            Case """"   ' Quotation mark
                index += 1
                Return """"
            Case "\"   ' Reverse solidus
                index += 1
                Return "\"
            Case "/"   ' Solidus
                index += 1
                Return "/"
            Case "b"   ' Backspace
                index += 1
                Return vbBack
            Case "f"   ' Formfeed
                index += 1
                Return vbFormFeed
            Case "n"   ' Newline
                index += 1
                Return vbNewLine
            Case "r"   ' Carriage return
                index += 1
                Return vbCr
            Case "t"   ' Horizontal Tab
                index += 1
                Return vbTab
            Case "u"   ' 4 hexadecimal digits
                Dim hex As String = String.Empty

                hex += c(index + 1) + c(index + 2) + c(index + 3) + c(index + 4)
                index += 4

                Return hex
            Case Else   ' Unknown
                Return String.Empty
        End Select
    End Function

    ' 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, -, .
    Private Shared Function ParseNumber(ByVal c As Char(), ByRef index As Integer) As Double
        Dim result As Double = Nothing
        Dim s As StringBuilder = New StringBuilder
        Dim chars() As Char = ("01234567890-.").ToCharArray()

        TrimWhiteSpace(c, index)

        While (c(index).ToString().IndexOfAny(chars) >= 0)
            s.Append(c(index))
            index += 1
        End While

        Double.TryParse(s.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, result)

        Return result
    End Function

    ' true, false
    Private Shared Function ParseBoolean(ByVal c As Char(), ByRef index As Integer) As Boolean
        If (c(index) = "t") Then
            index += 4

            Return True
        ElseIf (c(index) = "f") Then
            index += 5

            Return False
        Else
            Return Nothing
        End If
    End Function

    Private Shared Function GetToken(ByVal c As Char(), ByRef index As Integer) As Integer
        TrimWhiteSpace(c, index)

        If (index <= c.Length) Then
            Dim val As String = c(index).ToString()

            Select Case (val)
                Case "{"
                    Return Token.CurlyOpen
                Case "}"
                    Return Token.CurlyClose
                Case "["
                    Return Token.SquareOpen
                Case "]"
                    Return Token.SquareClose
                Case ","
                    Return Token.Comma
                Case ":"
                    Return Token.Colon
                Case """"
                    Return Token.String
                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "-"
                    Return Token.Number
                Case Else
                    If (val = "t") Then
                        ' true
                        If (c(index + 1) = "r" AndAlso c(index + 2) = "u" AndAlso c(index + 3) = "e") Then
                            Return Token.Boolean
                        End If
                    ElseIf (val = "f") Then
                        ' false
                        If (c(index + 1) = "a" AndAlso c(index + 2) = "l" AndAlso c(index + 3) = "s" AndAlso c(index + 4) = "e") Then
                            Return Token.Boolean
                        End If
                    ElseIf (val = "n") Then
                        ' null
                        If (c(index + 1) = "u" AndAlso c(index + 2) = "l" AndAlso c(index + 3) = "l") Then
                            Return Token.Null
                        End If
                    End If
            End Select
        End If

        Return Token.None
    End Function

    Private Shared Sub TrimWhiteSpace(ByVal c As Char(), ByRef index As Integer)
        While (index < c.Length)
            If (c(index) = " " OrElse c(index) = vbCr OrElse c(index) = vbLf OrElse c(index) = vbTab) Then
                index += 1
            Else
                Exit While
            End If
        End While
    End Sub

End Class