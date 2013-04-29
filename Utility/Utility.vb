Imports System.Text.RegularExpressions
Public Class Utility
    Public Shared Function CamelPascalToWords(ByVal text As String) As String
        'converts camel-case text to seperate words.
        Return Join(SplitUpperCase(text), " ")
    End Function
    Public Shared Function InitCap(ByVal text As String) As String
        If text.Length <= 2 Then Return text.ToLower
        Return text.Substring(0, 1).ToUpper & text.Substring(1).ToLower
    End Function

    Public Shared Function WordsToCamelPascal(ByVal text As String) As String
        Return text.Replace(" ", "")

    End Function

    Public Shared Function SplitUpperCase(ByVal source As String) As String()
        If source Is Nothing Then
            Return New String() {}

        End If
        'Return empty array.
        If source.Length = 0 Then
            Return New String() {""}

        End If
        Dim words As New Specialized.StringCollection()
        Dim wordStartIndex As Integer = 0
        Dim letters As Char() = source.ToCharArray()
        For i As Integer = 1 To letters.Length - 1
            ' Skip the first letter. we don't care what case it is.
            If Char.IsUpper(letters(i)) Then
                'Grab everything before the current index.
                words.Add(New String(letters, wordStartIndex, i - wordStartIndex))
                wordStartIndex = i
            End If
        Next

        'We need to have the last word.
        words.Add(New String(letters, wordStartIndex, letters.Length - wordStartIndex))
        'Copy to a string array.
        Dim wordArray As String() = New String(words.Count - 1) {}
        words.CopyTo(wordArray, 0)
        Return wordArray
    End Function

End Class
