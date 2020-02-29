Option Explicit

dim password, user_input, shell, x, y

set shell = createobject("wscript.shell")

Randomize
x = round(Int(15000)*Rnd()+1)
Randomize
y = round(Int(8500)*Rnd()+1)

password = "watashi-wa-firipinjin-desu"
user_input = inputbox("Password: ", "", "", x, y)

function duplicate()
    shell.run """C:\Users\LuizHackz\Documents\VBS\Duplicate Box\duplicate.vbs"""
    shell.run """C:\Users\LuizHackz\Documents\VBS\Duplicate Box\duplicate.vbs"""
    shell.run """C:\Users\LuizHackz\Documents\VBS\Duplicate Box\duplicate.vbs"""
end function

if not user_input = password then
    duplicate()
end if
