from subprocess import Popen, PIPE
import re
import obswebsocket
import obswebsocket.requests

## References
# PowerPoint Applescript Reference
#   http://www.codemunki.com/PPT2004AppleScriptRef.pdf
# Running AppleScript from python
#   https://stackoverflow.com/questions/2940916/how-do-i-embed-an-applescript-in-a-python-script#2941735
# Getting the notes via AppleScript
#   I want to send this guy money, I'd never of figured this out
#   https://stackoverflow.com/questions/12817639/how-can-i-access-the-text-of-notes-in-a-ppt-file-via-applescript
# Scott Hanselman's PowerPointToOBSSceneSwitcher
#   https://github.com/shanselman/PowerPointToOBSSceneSwitcher
# OBS Websocket access with python
#   https://github.com/Elektordi/obs-websocket-py

obsSearch = re.compile("OBS:(.*)")
sceneDefault = "DefaultPPTX"
obsWSPasswd = "2020Sucks"

get_Notes = """
    on run {}
        tell application "Microsoft PowerPoint"                        
            set slideNum to current show position of slide show view of slide show window 1
            set slideNote to content of text range of text frame of place holder 2 of notes page of slide slideNum of active presentation
            return slideNote
        end tell
    end run
"""

first_slide = """
    on run {}
        tell application "Microsoft PowerPoint"
            go to first slide of slide show view of slide show window 1
        end tell
    end run
"""


def run_applescript(script):
    p = Popen(["osascript", "-"], stdin=PIPE, stdout=PIPE, stderr=PIPE)
    stdout, stderr = p.communicate(script.encode("utf-8"))
    return (p.returncode, stdout.decode("utf-8"), stderr.decode("utf-8"))


def get_pptxScene():
    asReturn, asOut, asErr = run_applescript(get_Notes)
    # print(f"Return code = {asReturn}\nstdout = {asOut.strip()}\nstderr ={asErr.strip()}")
    try:
        scene = obsSearch.match(asOut).group(1)
    except (AttributeError):
        scene = sceneDefault
    return scene


def main():
    client = obswebsocket.obsws("localhost", 4444, obsWSPasswd)
    client.connect()
    asReturn, asOut, asErr = run_applescript(first_slide)
    scene = get_pptxScene()  # Set first scene
    response = client.call(obswebsocket.requests.SetCurrentScene(scene))
    client.disconnect()
    return 0


if __name__ == "__main__":
    main()
