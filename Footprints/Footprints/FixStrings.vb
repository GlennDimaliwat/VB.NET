Public Class FixStrings
    Function fixString(ByVal text As String) As String
        Dim newText As String
        newText = text
        newText = newText.Replace("null", "")
        newText = newText.Replace("__b", " ")
        newText = newText.Replace("__7", "&")
        newText = newText.Replace("__u", "-")
        newText = newText.Replace("__P", "(")
        newText = newText.Replace("__p", ")")
        newText = newText.Replace("__M", ",")
        newText = newText.Replace("__f", "/")
        newText = newText.Replace("__A", "+")
        newText = newText.Replace("__a", "'")
        newText = newText.Replace("__d", ".")

        '-- HTML Only
        newText = newText.Replace("&#160;", " ")
        newText = newText.Replace("&#58;", ":")
        newText = newText.Replace("&#39;", "'")
        newText = newText.Replace("&#34;", """")
        newText = newText.Replace("&#96;", "`")
        newText = newText.Replace("&#60;", "<")
        newText = newText.Replace("&#62;", ">")
        newText = newText.Replace("&#38;", "&")
        newText = newText.Replace("&#162;", "¢")
        newText = newText.Replace("&#163;", "£")
        newText = newText.Replace("&#165;", "¥")
        newText = newText.Replace("&#8364;", "€")
        newText = newText.Replace("&#169;", "©")
        newText = newText.Replace("&#174;", "®")

        '-- Unusual Stuff
        newText = newText.Replace("â€™", "'")
        newText = newText.Replace(ChrW(&H201C), "&#8220;")
        newText = newText.Replace("â€&#8220;", "-")
        newText = newText.Replace("â€œ", """")
        newText = newText.Replace("â€", """")

        Return newText
    End Function

    Function unfixString(ByVal text As String) As String
        Dim newText As String
        newText = text
        newText = newText.Replace(" ", "__b")
        newText = newText.Replace("&", "__7")
        newText = newText.Replace("-", "__u")
        newText = newText.Replace("(", "__P")
        newText = newText.Replace(")", "__p")
        newText = newText.Replace(",", "__M")
        newText = newText.Replace("/", "__f")
        newText = newText.Replace("+", "__A")
        newText = newText.Replace("'", "__a")
        newText = newText.Replace(".", "__d")
        Return newText
    End Function
End Class
