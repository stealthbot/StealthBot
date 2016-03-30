# StealthBot
Battle.net (pre-2.0) chat bot written in Visual Basic 6

The non-development-focused wiki can be found [here](http://www.stealthbot.net/wiki/index.php?title=Main_Page). Our forums can be found [here](http://www.stealthbot.net/forum/index.php?/index).

If you'd like to explore or contribute to the project, I'd recommend reading the [StealthBot Developer Guide](https://github.com/stealthbot/StealthBot/blob/master/docs/StealthBot%20Developer%20Guide.pdf). It's a little bit out of date, but it'll help you find your way around.


## Debugging 
You'll need Microsoft Visual Basic 6.0. 

1. Go into the project properties and set the compilation argument `COMPILE_DEBUG = 1`. If you don't do this, closing the program or encountering an error while debugging will cause the entire IDE to crash.

2. While still in the project properties, set the command line to `-ppath "C:\Path\To\Profile\" -addpath "C:\Path\To\Install\"` with the quotes and ending slashes. This sets your profile and install location to some paths of your choosing.

 * Set the `-ppath` to a profile just for development (such as the path to  `%APPDATA%\StealthBot\DevProfile\`, using the full path to your AppData StealthBot profiles location). If you don't do this, the development bot will expect files like commands.xml to be in your VB98 folder (the install directory for VB 6.0) instead of somewhere sane.

 * Set the `-addpath` to your install directory (such as `C:\Program Files (x86)\StealthBot v2.7\`). If you don't do this, the development bot will expect `BNCSutil.dll`, `Warden.dll`, and `zlib1.dll` in your VB98 folder also.

##### Visual Basic 6 Scroll Wheel Fix
The Visual Basic 6 IDE does not have built-in support for scrolling by using the mouse wheel.
To fix this, you need to download and follow instructions from [this KB article](https://support.microsoft.com/en-us/kb/837910).

If you are on 64-bit Windows, you will need to use the 32-bit version of regsvr32, located in `%systemroot%\SysWoW64\` (see [this KB article](https://support.microsoft.com/en-us/kb/249873)).
