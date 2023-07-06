set wShell = createObject("wscript.shell")


Dim x
x=1
Do Until x=1000
set webbrowser = createobject("internetexplorer.application")
webbrowser.visible = true
webbrowser.fullscreen = true
wscript.sleep(500)
webbrowser.navigate ("https://10minutemail.net")
wscript.sleep(3000)
With CreateObject("WScript.Shell")
    .Run "nircmd setcursor 650 725", 0, True
    .Run "nircmd sendmouse left click", 0, True
End With
wscript.sleep(3000)
webbrowser.visible = false
wscript.sleep (100)

set konkurs = createobject("internetexplorer.application")
konkurs.visible = true
konkurs.fullscreen = true
konkurs.navigate("https://competition.ifootagegear.com/video?id=5d3e17f181feca5c37de18ac")


wscript.sleep (4000)
Dim email
email = webbrowser.Document.getElementById("fe_text").value
With CreateObject("WScript.Shell")
    .Run "nircmd setcursor 525 275", 0, True
    .Run "nircmd sendmouse left click", 0, True
End With
wscript.sleep (100)

With CreateObject("WScript.Shell")
    .Run "nircmd setcursor 650 350", 0, True
    .Run "nircmd sendmouse left click", 0, True
End With
konkurs.document.all.item("email").value = email
With CreateObject("WScript.Shell")
    .Run "nircmd setcursor 650 350", 0, True
    .Run "nircmd sendmouse left click", 0, True

wShell.sendKeys "{BACKSPACE}"
wShell.sendKeys "m"
End With
wscript.sleep (200)
With CreateObject("WScript.Shell")
    .Run "nircmd setcursor 675 420", 0, True
    .Run "nircmd sendmouse left click", 200, True

End With
wscript.sleep (5000)
wShell.SendKeys "%{F4}"
wscript.sleep (2000)
webbrowser.visible = true
wscript.sleep (8200)

With CreateObject("WScript.Shell")
    .Run "nircmd setcursor 675 900", 0, True
    .Run "nircmd sendmouse left click", 200, True
End With

wscript.sleep (4000)
With CreateObject("WScript.Shell")
    .Run "nircmd setcursor 675 900", 0, True
    .Run "nircmd sendmouse wheel -600", 0, True
    .Run "nircmd setcursor 960 720", 0, True
    .Run "nircmd sendmouse left click", 0, True
    .Run "nircmd setcursor 960 610", 0, True
    .Run "nircmd sendmouse left click", 0, True
End With

wscript.sleep (5000)
With CreateObject("WScript.Shell")
    .Run "nircmd sendmouse left click", 200, True
End With
wscript.sleep (100)
wShell.SendKeys "%{F4}"
x=x+1
wscript.sleep (1000)

Loop





