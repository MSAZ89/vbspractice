'Run in console: cscript //nologo test.vbs'

Dim age

age = 30

WScript.Echo "Age: " & age
WScript.Echo "Next year, age will be: " & (age + 1)

if age >= 18 then
    WScript.Echo "You are an adult."
else
    WScript.Echo "You are not an adult."
end if