'Run in console: cscript //nologo collectionTypes.vbs'

' Fixed-size
Dim arr(2)   ' 3 elements: 0, 1, 2
arr(0) = "apple"
arr(1) = "banana"
arr(2) = "cherry"

' Dynamic
Dim dyn()
ReDim dyn(0)
dyn(0) = "first"
ReDim Preserve dyn(1)
dyn(1) = "second"

' Loop
Dim i
For i = 0 To UBound(arr)
    WScript.Echo arr(i)
Next
