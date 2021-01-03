# obsPPTXwithPython

Inspired by Scott Hanselman's [PowerPointToOBSSceneSwitcher][1], I was looking for a way to do this on macOS. This uses Applescript to both control the slides (easy) and extract the notes (harder than it needed be). The Applescript is run from the python, so no need to have multiple things running at a time. I spent a lot more time trying to get Applescript to work than getting anything else running.

This is a MVP, so don't expect perfection. You can advance or retreat in a pptx presentation and OBS will change scenes as expected. Tested against PowerPoint 16.44, but more testing would be good.

You need to add "OBS:<SCENENAME>" to the notes section, just like in Scott's implmentation. You should have the presentation in Slide Show view for this to work out of the gate. Just like Scott's, you'll have a terminal open to run. 

I'm using pygame for keyboard capture.   
* pygame was quick and easy, so it get's the nod for v1  
* Looked at the keyboard package, but needs root and I wanted to avoid that
* Looked at curses, which might be better  

## One Trick Pony Scripts

I have split obsPPTX.py into three scripts, obsPPTXStart, obsPPTXAdvance, and obsPPTXRetreat, which only do one thing each. This should allow the script to be run from a hotkey launcher such as Keysmith, Keyboard Maestro, Alfred, etc. The benefit of this is that pygame is no longer needed. The websocket connection is opened and closed with every run, but this seems fairly quick.

## Requirements
These scripts were developed with python 3.7.3. The only other requirement to install is obs-websocket-py, via `pip install obs-websocket-py`

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
