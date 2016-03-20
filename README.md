# StealthBot
Battle.net (pre-2.0) chat bot written in Visual Basic 6

The non-development-focused wiki can be found [here](http://www.stealthbot.net/wiki/index.php?title=Main_Page). Our forums can be found [here](http://www.stealthbot.net/forum/index.php?/index).

If you'd like to explore or contribute to the project, I'd recommend reading the [StealthBot Developer Guide](https://github.com/stealthbot/StealthBot/blob/master/docs/StealthBot%20Developer%20Guide.pdf). It's a little bit out of date, but it'll help you find your way around.

## Debugging 
You'll need Microsoft Visual Basic 6.0. 

1. Go into the project properties and set the compilation argument "COMPILE_DEBUG = 1". If you don't do this, closing the program or encountering an error while debugging will cause the entire IDE to crash.

2. Locate your visual studio installation folder.
  * C:\Program Files (x86)\Microsot Visual Studio\VB98, or
  * Start the debugger and when the main window appears, go to Settings -> View Files -> Open Bot Folder.

3. Copy the following files from the repo into this folder:
  * BNCSUtil.dll
  * CheckRevision.ini
  * commands.xml
  * Warden.dll
  * Warden.ini
  * zlib1.dll
