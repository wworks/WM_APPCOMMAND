# WM_APPCOMMAND
A VB.NET Library for using the appcommands in windows messages.

In this library you'll find everything you need to conveniently work with the WM_APPCOMMAND( https://msdn.microsoft.com/en-us/library/windows/desktop/ms646275%28v=vs.85%29.aspx).  
It contains subs to work with all of these commands, nicely organized.(Also contains enums for all the commands). You can use them to automate stuff or make remotes/keyboard mods.

Sample usage:  
WmAppCommands.Browser.Forward(me.handle)  
  
All possible tasks:  
Browser  
* Backward  
* Forward  
* Favorites  
* Home  
* Refresh  
* Search  
* Stop  

Media  
* ChannelDown  
* ChannelUp  
* FastForward  
* NextTrack  
* PreviousTrack  
* Play
* Pause
* PlayPause
* Record
* Rewind
* Stop

Microphone
* OnOffToggle
* VolumeDown
* VolumeUp
* VolumeMute

Sound
* Volume
  * Up
  * Down
  * Mute
* Bass
  * Up
  * Down
  * Boost
* Treble
  * Up
  * Down
  
Launch
  * Mail
  * MediaSelect
  * App1
  * App2
  
Actions(the stuff you find in programs like word, etc)
  * Clipboard
    * Copy
    * Cut
    * Paste
  * Mail
    * ReplyToMail
    * ForwardMail
    * SendMail
  
  * Find
  * Help
  * New
  * Open
  * Close
  * Save
  * Print
  * Undo
  * Redo
  * SpellCheck
  * DictateOrCommandControlToggle
  * Correctionlist
  
Further explanation can be found on the microsoft website.

I haven't tested all of them, but not all of them seem to work, e.g. Browser.Back

William Wernsen
  
  

