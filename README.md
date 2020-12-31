# obsPPTXwithPython

Inspired by Scott Hanselman's [PowerPointToOBSSceneSwitcher][1], I was looking for a way to do this on macOS. Many hours consumed trying to re-adjust to AppleScript while dealing with MSFT's PowerPoint dictionary, but got there.

This is a MVP, but it is working. You can advance or retreat in a pptx presentation and OBS will change scenes as expected. Tested against PowerPoint 16.44.

You need to add "OBS:<SCENENAME>" to the notes section, just like in Scott's implmentation. You should have the presentation in Slide Show view for this to work. 

I'm using pygame for keyboard capture.   
* Looked at the keyboard package, but needs root and I wanted to avoid that like the plague.
* Looked at curses, which might be better  
* pygame was quick and easy, so it get's the nood for v1  


## Other References
* [PowerPoint Applescript Reference][2]
* [Running AppleScript from python][3]
* [Getting the notes via AppleScript][4]
    * I want to send this guy money, I'd never of figured this out
* [OBS Websocket access with python][5]

[1]:https://github.com/shanselman/PowerPointToOBSSceneSwitcher
[2]:http://www.codemunki.com/PPT2004AppleScriptRef.pdf
[3]:https://stackoverflow.com/questions/2940916/how-do-i-embed-an-applescript-in-a-python-script#2941735
[4]:https://stackoverflow.com/questions/12817639/how-can-i-access-the-text-of-notes-in-a-ppt-file-via-applescript
[5]:https://github.com/Elektordi/obs-websocket-py
