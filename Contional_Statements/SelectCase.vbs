'colour =inputbox(" Enter The Colour")

'Select case colour
'case "Red"
'msgbox("Red Colour Is Selected")
'case "Blue"
'msgbox("Blue Colour Is Selected")
'case "Black"
'msgbox("Black Colour Is Selected")
'case "White"
'msgbox("White Colour Is Selected")
'case "Pink"
'msgbox("Pink Colour Is Selected")
'end select




dim var
'var=Null
'var=10
'var="test"
'var=True
var=10.20

select case VarType(var)
case vbEmpty
msgbox(" Value is empty ")
case vbNull
msgbox(" Value is Null ")
case vbInteger
msgbox(" Value is Integer")
case vbString
msgbox(" Value is String")
case vbBoolean
msgbox(" Value is Boolean")
case vbDouble
msgbox(" Value is Double")
case Else
msgbox(" Not Proper Data Type ")
end Select