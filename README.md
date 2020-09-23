<div align="center">

## Use Microsofts Agents in your apps\!


</div>

### Description

Taken from Windows Magazine August 1999

Agent is Microsoft's name for the technology that puts those "cute" cartoon characters on the screen while running any Microsoft Office application. This technoloty will be a standard feature of Windows 2000, but you can also add it to Windows 98 or Windows 95 viea a free download at "HTTP://msdn.microsoft.com/workshop/imedia/agent/agentdl.asp" Once you install this update, you can use these characters in your applications, including macros!
 
### More Info
 
you can download the free text-to-speech converter and let it talk to you


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ryan Couch](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ryan-couch.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ryan-couch-use-microsofts-agents-in-your-apps__1-2770/archive/master.zip)





### Source Code

```
Dim objAgent
Dim objChar
Dim objRequest
Dim txtSpeak
Dim strName
set objAgent = CreateObject("Agent.Control.1")
objAgent.Connected = True
strName = "Peedy" 'you can use genie, or merlin, or robby or whatever
objAgent.Characters.Load strName, strName & ".acs"
Set objChar = objAgent.Characters(strName)
'objChar.LanguageID = &h409
objChar.Show
objChar.Speak("Hello! I'm " & strName)
objChar.Play "Wave"
txtSpeak = "What should I say next?"
while txtSpeak > ""
  objChar.Speak txtSpeak
  'objChar.Play "Hearing_1"
  txtSpeak = InputBox("What should I say next?", "Peedy App")
wend
objChar.Speak "Goodbye!"
objChar.Hide
msgbox "Goodbye!", vbokonly, "Peedy App"
Set objChar = Nothing
objAgent.Characters.Unload strName
```

