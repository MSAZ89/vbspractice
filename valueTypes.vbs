'Run in console: cscript //nologo valueTypes.vbs'

Dim x
x = 10          ' x is a number right now
WScript.Echo x & " is of type " & VarType(x)

x = "hello"     ' now x is a string
WScript.Echo x & " is of type " & VarType(x)

x = Now         ' now x is a Date
WScript.Echo x & " is of type " & VarType(x)

'Value	Type'
'0	Empty'
'1	Null'
'2	Integer'
'3	Long'
'4	Single'
'5	Double'
'6	Currency'
'7	Date'
'8	String'
'9	Object'
'11	Boolean'
'12	Variant/Array'