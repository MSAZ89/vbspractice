'Run in console: cscript //nologo stringComparison.vbs'

Dim a, b
a = "apple"
b = "banana"

If a = b Then
    WScript.Echo "Equal"
Else
    WScript.Echo "Not equal"
End If

If a < b Then
    WScript.Echo a & " comes before " & b
End If

If b > a Then
    WScript.Echo b & " comes after " & a
End If
