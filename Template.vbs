set a = createobject("wscript.shell")
a.run "https://www.yourwebsite.com/"
wscript.sleep (5000)
a.sendkeys ("Username here")
a.sendkeys chr(9)
wscript.sleep (2000)
a.sendkeys ("Password here")
a.sendkeys"{Enter}"
' call msgbox("Finished (Remove the semi-colon if you want this))")
wscript.quit