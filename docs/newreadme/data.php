<?php
    /* 
    *   NewReadme for StealthBot
    *       December 2005 by Andy T (Stealth)
    *       Online readme and FAQ system
    */
    
    
    /**
            
    */

    defined("__READMEFILE") or die("Invalid access");
    
    /**
    *   COMMANDS
    */
    $COMMANDS = array(
            // 20
            array(
                "find", "about", "server", "add", "whoami", "cq", "scq",
            "designated", "findattr", "flip", "bmail", "roll",
                ),
            // 40
            array(
                "time", "trigger", "pingme",
                ),
            // 50
            array(
            "say", "shout", "ignore", "unignore", "addquote", "quote", "away", "back",
            "ping", "uptime", "mp3", "owner", "vote", "voteban", "votekick",
            "tally", "info", "expand", "math", "where", "safecheck", "cancel"
                ),
            // 60
            array(
            "kick", "ban", "unban", "lastwhisper", "define", "newdef", "fadd",
            "frem", "bancount", "allseen", "levelbans", "d2levelbans", "tagcheck",
                ),
            // 70
            array(
            "shitlist", "shitdel", "safeadd", "safedel", "safelist", "tagbans", "tagadd",
            "tagdel", "protect", "shitadd", "mimic", "nomimic", "check", "online",
            "monitor", "unmonitor", "cmdadd", "cmddel", "cmdlist",
            "setpmsg", "mmail", "addphrase", "phrases", "delphrase",
            "pon", "poff", "pstatus", "ipban", "banned", "notify", "denotify",
                ),
            // 80
            array(
            "reconnect", "designate", "rejoin", "settrigger", "igpriv", "unigpriv",
            "rem", "sethome", "idle", "next", "play", "stop", "setvol", "fos", 
            "pause", "shuffle", "repeat", "idletime", "idletype", "block", "filter", 
            "whispercmds", "profile", "greet", "levelban", "d2levelban", "allowmp3",
            "clientbans", "cadd", "cdel", "koy", "plugban", "useitunes", "setidle",
                ),
            // 90
            array(
            "join", "home", "resign", "setname", "setpass", "setserver", "quiettime", "giveup",
            "readfile", "chpw", "clan", "sweepban", "sweepignore", "idlebans",
            "setkey", "setexpkey", "exile", "unexile", "clearbanlist",
                ),
            // 100
            array(
            "quit", "locktext", "efp", "loadwinamp", "setmotd", "invite", "peonban"
                ),
        );
        
    $ACCESSLEVELS = array(20, 40, 50, 60, 70, 80, 90, 100);
        
    $ALIASES = array(
                "findattr" => array("findflag"),
                "about" => array("ver", "version"),
                "bmail" => array("mail"),
                "find" => array("whois"),
                "add" => array("set"),
                "ignore" => array("ign"),
                "define" => array("def"),
                "math" => array("eval"),
                "safeadd" => array("safelist <username>"),
                "shitadd" => array("pban", "shitlist <username>"),
                "cmdadd" => array("addcmd"),
                "cmddel" => array("delcmd"),
                "addphrase" => array("padd"),
                "phrases" => array("plist"),
                "delphrase" => array("pdel"),
                "pon" => array("phrasebans on"),
                "poff" => array("phrasebans off"),
                "pstatus" => array("phrasebans"),
                "designate" => array("des"),
                "rejoin" => array("rj"),
                "rem" => array("del"),
                "clientbans" => array("cbans", "clist"),
                "idlebans" => array("ib"),
                "sweepban" => array("cb"),
                "sweepignore" => array("cs"),
                "clearbanlist" => array("cbl"),
                "clan" => array("c"),
                "efp" => array("floodmode"),
                "pingme" => array("getping"),
                "giveup" => array("op"),
                );
                
    $stealth = "&lt;Stealth&gt";
    $stealthbot = "&lt;StealthBot&gt";
    $importantNote = "<b>Important note:</b> The say and shout commands are intelligently scaled based on how much access you have. Users with access of 69 or below will have their messages prefixed with &quot;Username says: &quot; to prevent abuse. Users with access 70 to 89 will have any slashes (\"/\") at the beginning of their message removed. Only users with 90 access or above can execute unrestricted say commands.";
                
    $CMD_DESC = array(   // array(ACCESS, ARGUMENTS, WILDCARDABLE?, DESCRIPTION, EXAMPLE) 
            // 20
            "find" => array(20, ".find (username)<br>The username to be searched.", 1, "Find searches the bot's database for a username and tells you whether or not they exist, and if they do, their access and/or flags.", "/find Stealth<br>Stealth has access 100 and flags ADMOST."),
            "about" => array(20, "", 0, "Displays the program version.", "/about<br>StealthBot v2.7 by Stealth"),
            "server" => array(20, "", 0, "Displays the server the bot is currently connected to.", "/server<br>I am currently connected to [ useast.battle.net ]"),
            "add" => array(20, ".add (username) [access] [flags]<br>The username to be added, followed by the desired access and flags. You can also use the + and - operators on flags to add or remove individual flags from a person's access level. See the second example.",
                            0, "Adds or modifies a user's access to the bot.", "/add Stealth 100<br>Set Stealth's access to 100.<br><br>/add Stealth ASDF<br>Set Stealth's flags to ASDF.<br><br>/add Stealth -S<br>Set Stealth's flags to ADF.<br><br>/add Stealth +Y<br>Set Stealth's flags to ADFY."),
            "whoami" => array(20, "", 0, "Displays the access of the user who tries to use the command.", "/whoami<br>You are the bot console."),
            "cq" => array(20, "", 0, "Clears the bot's queue noisily. The queue contains all the messages that the bot intends to send as soon as it can.", "/cq<br>Queue cleared."),
            "scq" => array(20, "", 0, "Clears the bot's queue silently. The queue contains all the messages that the bot intends to send as soon as it can.", "/scq<br>(the queue is cleared. no response is displayed.)"),
            "designated" => array(20, "", 0, "Displays the user that the bot has designated, if there is one.", "/designated<br>I have designated \"StealthBot2\"."),
            "findattr" => array(20, ".find (flag letter(s))<br>The flag or flag letters to be searched.", 0, "Searches the database for users with the specified flag or flag combination.", "/findattr p<br> Users found: personb, personf, personwithverylongname, personwithsmallerlongname [more]<br>Users found: otherpeople"),
            "flip" => array(20, "", 0, "Flips a coin and tells you the result.", "/flip<br>Tails."),
            "bmail" => array(20, ".bmail (recipient) (message)", 0, "Sends botmail to the recipient.", "/mail deadly7 Hey! What's up?<br>Added mail for deadly7."),
            "roll" => array(20, ".roll [top boundary]<br>Top boundary is optional.", 0, "Rolls a random value from 0 to 100. If a top boundary is specified, it will roll from 0 to [top boundary].", "/roll<br>94"),
            // 40
            "time" => array(40, "", 0, "Displays the computer's current time.", "/time<br>The current time on this computer is 11:10:54 PM, 12-26-2005."),
            "trigger" => array(40, "", 0, "Displays the bot's current trigger. (Useful only in the bot console)", "/trigger<br>Trigger is currently \"-\"."),
            "pingme" => array(40, "", 0, "Tells the person executing the command what their ping is.", "$stealth .pingme<br>$stealthbot Your ping at login was 31ms."),
            // 50
            "say" => array(50, ".say (message to be spoken)", 0, "The bot will repeat whatever follows. $importantNote", "$stealth say Taco Bell is the best!<br>$stealthbot Taco Bell is the best!"),
            "shout" => array(50, ".shout (message to be shouted)", 0, "The bot will repeat whatever follows, in all caps to appear as if shouting. $importantNote", "$stealth shout I love taco bell!<br>$stealthbot I LOVE TACO BELL!"),
            "ignore" => array(50, ".ignore (user to be ignored)", 0, "The bot will instruct Battle.net to ignore further messages from the target user. If IPBans are enabled, this will also IPBan them from the channel.", "$stealth .ignore anAnnoyingPerson<br>$stealthbot Ignoring messages from \"anAnnoyingPerson\"."),
            "unignore" => array(50, ".unignore (user to hear)", 0, "The bot will instruct Battle.net to once again allow messages from the target user. If IPBans are enabled, this will also un-IPBan the user.", "$stealth .unignore aNormalPerson<br>$stealthbot Receiving messages from \"NormalPerson\"."),
            "addquote" => array(50, ".addquote (quote to be added)", 0, "Adds the given quote to the bot's quotes database.", "$stealth .addquote What the devil do you think you're doing?<br>$stealthbot Quote added!"),
            "quote" => array(50, "", 0, "Draws a random quote from the bot's quotes.txt file and displays it.", "/quote<br>What the devil do you think you're doing?"),
            "away" => array(50, ".away [away message]", 0, "Sets the bot's away status to your desired message.", "$stealth .away I'm not here<br>&lt;StealthBot is away (I'm not here)&gt;"),
            "back" => array(50, "", 0, "Clears any away messages present.", "/back<br>You are no longer marked as away."),
            "ping" => array(50, ".ping (username to be pinged)", 0, "Ping retrieves the target's ping at logon if they're in the channel.", "$stealth .ping deadly7<br>$stealthbot Deadly7's ping at logon was 72ms."),
            "uptime" => array(50, "", 0, "Displays the system uptime and the bot's current connection uptime.", "/uptime<br>System uptime 3 days, 2 hours, 33 minutes and 4 seconds, connection uptime 0 days, 3 hours, 4 minutes and 12 seconds."),
            "mp3" => array(50, "", 0, "Displays Winamp's current track title. Also works in iTunes with the <a href='index.php?act=viewCmd&c=useitunes'>useitunes</a> toggle.", "$stealth .mp3<br>$stealthbot Current MP3: 51. Avenged Sevenfold - Bat Country"),
            "owner" => array(50, "", 0, "Displays the current Bot Owner.", "$stealth .owner<br>$stealthbot This bot's owner is Stealth."),
            "vote" => array(50, ".vote (number of seconds for the vote to last)", 0, "Instructs the bot to take Yes or No votes. The <a href='index.php?act=viewCmd&c=tally'>tally</a> command can be used to see current results, and the <a href='index.php?act=viewCmd&c=cancel'>cancel</a> command will stop the vote.", "$stealth .vote 45<br>$stealthbot Vote initiated. Type YES or NO to vote; your vote will be counted only once."),
            "voteban" => array(50, ".voteban (name of person to be banned)", 0, "The bot will start a Yes or No vote as to whether or not to ban the target user. If the vote succeeds, the user will be banned.", "$stealth .voteban aBadPerson<br>$stealthbot 30-second VoteBan vote started. Type YES to ban abadperson, NO to acquit him/her."),
            "votekick" => array(50, ".votekick (name of person to be kicked)", 0, "Starts a Yes or No vote as to whether or not to kick the target user. If the vote succeeds, the user will be kicked.", "$stealth .votekick aBadPerson <br>$stealthbot 30-second VoteKick vote started. Type YES to kick  abadperson, NO to acquit him/her."),
            "cancel" => array(50, "", 0, "This command stops any <a href='index.php?act=viewCmd&c=vote'>vote</a>, <a href='index.php?act=viewCmd&c=votekick'>votekick</a>, or <a href='index.php?act=viewCmd&c=voteban'>voteban</a> currently in progress.", "$stealth .cancel<br>$stealthbot Vote cancelled. Final results: [4] YES, [1] NO. "), 
            "tally" => array(50, "", 0, "Displays the current Yes and No vote totals for an ongoing vote.", "<Stealth> .tally<br><StealthBot> Current results: [0] YES, [0] NO; 48 seconds remain. The vote is a draw."),
            "info" => array(50, ".info (user to describe)", 0, "Displays some basic information about the specified user.", "$stealth .info Stealth<br>$stealthbot User Stealth is logged on using Starcraft Original with a ping time of 93ms.<br>$stealthbot He/she has been present in the channel for 0 days, 0 hours, 9 minutes and 8 seconds."),
            "expand" => array(50, ".expand (phrase to be expanded)", 0, "Repeats the phrase you give it with spaces between the letters.", "$stealth .expand Hello, world!<br>$stealthbot H e l l o ,   w o r l d !"),
            "math" => array(50, ".math (VBScript expression)", 0, "Math and eval allow you to execute a simple VBScript expression and display its results. One practical use is to give it a math expression and it will return the solution.", "$stealth .math (251 + 667) + (215 * 4)<br>$stealthbot 1778"),
            "where" => array(50, "", 0, "Displays the bot's current Battle.net location.", "$stealth .where<br>$stealthbot I am currently in channel Clan SBS (9 users present)"),
            "safecheck" => array(50, ".safecheck (username or tag to check)", 0, "This command tells you whether or not the username or tag you specify is currently safelisted.", "$stealth .safecheck Naivete<br>$stealthbot That user is safelisted."),
            // 60
            "kick" => array(60, ".kick (username to be kicked) [message]", 1, "Kicks the desired user. If you specify a message, it will be shown by Battle.net in the kick notice. $opmsg", "$stealth .kick anAnnoyingPerson You're really annoying!<br>anAnnoyingPerson was kicked by StealthBot (You're really annoying!)"),
            "ban" => array(60, ".ban (username to be banned) [message]", 1, "Bans the desired user. If you specify a message, it will be shown by Battle.net in the ban notice. $opmsg", "$stealth .kick anAnnoyingPerson You're REALLY annoying!<br>anAnnoyingPerson was banned by StealthBot (You're really annoying!)"),
            "unban" => array(60, ".unban (username to be unbanned)", 1, "Unbans the desired user.", "example"),
            "lastwhisper" => array(60, "", 0, "Displays the name of the last person who whispered the bot.", "example"),
            "define" => array(60, ".define (word to be defined)", 0, "Calls up a definition of the word you provided. If there is no definition, it will tell you.", "example"),
            "newdef" => array(60, ".newdef (term to define)|(definition)", 0, "Adds a new definition for the word you provide. Note that the pipe ( | ) must be present as used in the example.", "example"),
            "fadd" => array(60, ".fadd (friend to add)", 0, "Adds a user to your Battle.net friends' list.", "example"),
            "frem" => array(60, ".frem (friend to remove)", 0, "Removes a friend from your Battle.net friends' list.", "example"),
            "bancount" => array(60, "", 0, "Displays the current ban count as observed by the bot.", "example"),
            "allseen" => array(60, "", 0, "Displays the last 15 users seen by the bot.", "$stealth .allseen<br>$stealthbot Last 15 users seen: Distant.Echo@Azeroth, HdxBmx27@Azeroth, Naivete@Azeroth, Shox, Distant.Tech [more]<br>$stealthbot DeadHelp@Azeroth, Stealthbot-Tech, QuikHelp@Azeroth, Mega-Tech@Azeroth, God_Of_SlaYerS@Azeroth [more]<br>$stealthbot RoMi[x86], zeth.tech@Azeroth, StealthBot, Stealth"),
            "levelbans" => array(60, "", 0, "Displays the level under which your bot will ban Warcraft III users.", "example"),
            "d2levelbans" => array(60, "", 0, "Displays the level under which your bot will ban Diablo II users.", "example"),
            "tagcheck" => array(60, ".tagcheck (username or tag to check)", 0, "Checks to see whether or not a user is tagbanned. If they're tagbanned, it will tell you which banned wildcard tags they match.", "$stealth .tagcheck [xa]xefor<br>$stealthbot That user matches the following tagban(s): [xa]*"),
            // 70                                                
            "shitlist" => array(70, ".shitlist [username to add [message]]", 0, "If no arguments are provided, as in the first example, your whole shitlist is displayed. If you provide a username, that person will be added to the bot's shitlist.", "example"),
            "shitdel" => array(70, ".shitdel (username to remove)", 0, "Removes the target user from your bot's shitlist.", "example"),
            "safeadd" => array(70, ".safeadd (username to add)", 1, "Adds a user or wildcard tag to your bot's safelist.", "example"),
            "safedel" => array(70, ".safedel (username to remove)", 0, "Removes a user or wildcard tag from your bot's safelist.", "example"),
            "safelist" => array(70, ".safelist [username to add]", 1, "If no arguments are provided, as in the first example, your safelist is read to you. If you provide a username, that user will be added to the safelist.", "example"),
            "tagbans" => array(70, "", 0, "Lists the bot's current banned wildcard tags.", "example"),
            "tagadd" => array(70, ".tagadd (tag to add)", 1, "Adds a new wildcard tag to your bot's list of banned tags.", "example"),
            "tagdel" => array(70, ".tagdel (tag to remove)", 0, "Removes a wildcard tag from your bot's list of banned tags.", "example"),
            "protect" => array(70, ".protect (on/off/status)", 0, "<b>on</b> enables channel protection, <b>off</b> disables channel protection, and <b>status</b> will tell you whether or not it is currently enabled.", "example"),
            "shitadd" => array(70, ".shitadd (username to add) [message]", 0, "Adds a user to your bot's shitlist. If you provide a message, it will be used to ban them every time they join.", "example"),
            "mimic" => array(70, ".mimic (username)", 0, "Repeats everything that the target user says verbatim.", "example"),
            "nomimic" => array(70, "", 0, "Turns mimicing off.", "example"),
            "check" => array(70, ".check (username)", 0, "Checks to see whether the specified user is online according to the bot's User Monitor.", "example"),
            "online" => array(70, "", 0, "Displays the list of users that the bot's User Monitor sees online.", "example"),
            "monitor" => array(70, ".monitor (username)", 0, "Adds a username to the bot's User Monitor.", "example"),
            "unmonitor" => array(70, ".unmonitor (username)", 0, "Removes a username from the bot's User Monitor.", "example"),
            "cmdadd" => array(70, ".cmdadd (access required) (command word) (command actions)", 0, "Adds a custom command to the bot. See the Custom Command <a href='index.php?act=viewFeature&c=ccs'>feature page</a> for more detail.", "example"),
            "cmddel" => array(70, ".cmddel (command name)", 0, "Deletes the specified custom command.", "example"),
            "cmdlist" => array(70, "", 0, "Displays a list of all the Custom Commands you have created.", "example"),
            "setpmsg" => array(70, ".setpmsg (new message)", 0, "Sets a new message to use when banning people for Channel Protection.", "example"),
            "mmail" => array(70, ".mmail (access/flags to mail) (message)", 0, "Sends a bot mail to everyone with the specified access or flag(s).", "example"),
            "addphrase" => array(70, ".addphrase (phraseban to add)", 0, "Adds a new phraseban to the bot. Users who say that phrase and are not safelisted will be banned.", "example"),
            "phrases" => array(70, "", 0, "Lists your current phrasebans.", "example"),
            "delphrase" => array(70, ".delphrase (phraseban to remove)", 0, "Removes a phrase from your bot's phraseban list.", "example"),
            "pon" => array(70, "", 0, "Turns phrase banning on.", "example"),
            "poff" => array(70, "", 0, "Disables phrase banning.", "example"),
            "pstatus" => array(70, "", 0, "Displays the current status of phrase banning.", "example"),
            "ipban" => array(70, ".ipban (username)", 0, "Uses Battle.net's squelch feature to ban anyone on the same IP address as your target. Battle.net does a very effective job of hiding IP addresses, so you will never see the target's IP address, but they will be prevented from reentering your channel by changing usernames or cdkeys.", "example"),
            "banned" => array(70, "", 0, "Lists the users that your bot has seen banned since it joined the channel.", "example"),
            "notify" => array(70, ".notify (username in the monitor)", 0, "Sends you a message when the specified user is spotted as online by the bot's User Monitor.", "example"),
            "denotify" => array(70, ".denotify (username)", 0, "Removes online/offline notification from the specified user.", "example"),
            // 80
            "reconnect" => array(80, "", 0, "Instructs the bot to reconnect.", "example"),
            "designate" => array(80, ".designate (username)", 0, "Tells Battle.net that the specified user should get channel moderator status after you leave the channel.", "example"),
            "allowmp3" => array(80, "", 0, "Toggles whether or not the bot will accept MP3-related commands like <a href='index.php?act=viewCmd&c=play'>play</a> or <a href='index.php?act=viewCmd&c=stop'>stop</a>.", "example"),
            "rejoin" => array(80, "", 0, "Instructs the bot to leave and rejoin the channel.", "example"),
            "settrigger" => array(80, ".settrigger (new trigger)", 0, "Changes the bot's trigger character. By default, this character is a single period. Triggers cannot be more than one character long.", "example"),
            "igpriv" => array(80, "", 0, "Use wisely -- this command tells Battle.net you want to ignore ALL incoming messages from people who aren't on your friends' list in private channels. Private channels include Clan and Op channels, so be careful!", "example"),
            "unigpriv" => array(80, "", 0, "Reverses the effects of the <a href='index.php?act=viewCmd&c=igpriv'>igpriv</a> command.", "example"),
            "rem" => array(80, ".rem (username)", 0, "Removes a user from the bot's database.", "example"),
            "sethome" => array(80, ".sethome (new channel)", 0, "Sets a new home channel on the bot. It will automatically join this channel after connecting.", "example"),
            "idle" => array(80, ".idle (on/off/kick)", 0, "<b>on</b> and <b>off</b> control whether or not your bot displays an anti-idle message. <b>kick</b> affects whether or not idle bans will actually ban the user or simply kick them instead.", "example"),
            "next" => array(80, "", 0, "Plays the next track in Winamp or iTunes.", "example"),
            "play" => array(80, ".play [track number / part of a song name]", 0, "If it's by itself, this command will tell Winamp to start playing music. If you provide a track number, the bot will instruct Winamp to play that track. Otherwise, if you provide part of a song title, Winamp will attempt to match your song selection and play that song.", "example"),
            "stop" => array(80, "", 0, "Tells Winamp or iTunes to stop playing.", "example"),
            "setvol" => array(80, ".setvol (new volume percentage)", 0, "Sets Winamp's volume to a new value (0-100).", "example"),
            "fos" => array(80, "", 0, "Tells Winamp to do a fade-out stop.", "example"),
            "pause" => array(80, "", 0, "Pauses or unpauses Winamp or iTunes.", "example"),
            "shuffle" => array(80, "", 0, "Toggles Winamp's playlist shuffling feature.", "example"),
            "repeat" => array(80, "", 0, "Toggles Winamp's playlist repeat feature.", "example"),
            "idletime" => array(80, ".idletime (minutes)", 0, "Sets the delay between anti-idle messages.", "example"),
            "idletype" => array(80, ".idletype (quote / msg / mp3 / uptime)", 0, "Sets the type of anti-idle displayed. <b>quote</b> displays a random quote, <b>msg</b> displays a message of your choice, <b>mp3</b> displays Winamp or iTunes' current song title, and <b>uptime</b> displays your current computer and connection uptime.", "example"),
            "block" => array(80, ".block (username)", 1, "Blocks all messages from a username. Filtering must be turned on.", "example"),
            "filter" => array(80, ".filter (part of a message)", 0, "Filters any messages or emotes containing the desired phrase.", "example"), 
            "whispercmds" => array(80, "", 0, "Toggles whether or not command responses are whispered back.", "example"),
            "profile" => array(80, ".profile (username)", 0, "Reads aloud the target user's Battle.net profile.", "example"),
            "greet" => array(80, ".greet (on / off / whisper / your custom message)", 0, "<b>on</b> enables greet messages, while <b>off</b> disables them. <b>whisper</b> toggles whether or not they're whispered to the user who joins. If you don't specify on, off or whisper, the greet message will be set to whatever you provide.", "example"),
            "levelban" => array(80, ".levelban (new level to ban under)", 0, "Sets the new level under which your bot will ban Warcraft III users.", "example"),
            "d2levelban" => array(80, ".d2levelban (new level to ban under)", 0, "Sets the new level under which your bot will ban Diablo II users.", "example"),
            "clientbans" => array(80, "", 0, "Lists the bot's current ClientBans.", "<Stealth> .clientbans<br><StealthBot> Clientbans: D2DV, WAR3"),
            "cadd" => array(80, ".cadd (4-letter client code)", 0, "Adds a new ClientBan to your bot. Only accepts <a href='index.php?act=viewPage&page=codes'>4-letter game codes</a>.", "example"),
            "cdel" => array(80, ".cdel (4-letter client code)", 0, "Removes the specified ClientBan from your bot.", "example"),
            "koy" => array(80, ".koy (on/off/status)", 0, "Toggles <b>on</b> or </b>off</b> the bot's Kick On Yell moderation feature. Status will display whether or not the feature is enabled.", "example"),
            "plugban" => array(80, ".plugban (on/off/status)", 0, "Toggles and displays status of the bot's moderation Plugban feature. PlugBan bans users who have a UDP Plug, which appears instead of latency bars next to their name in the channel list. This usually means they're on a bot.", "example"),
            "useitunes" => array(80, "", 0, "Toggles whether or not music player commands like <a href='index.php?act=viewCmd&c=play'>play</a> or <a href='index.php?act=viewCmd&c=stop'>stop</a> are sent to Winamp or iTunes.", "example"),
            "setidle" => array(80, ".setidle (new idle message)", 0, "Sets the anti-idle message to your desired message.", "example"),  
            // 90                                                            
            "join" => array(90, ".join (new channel)", 0, "Moves the bot to a new channel.", "example"),
            "home" => array(90, "", 0, "Moves the bot to its home channel.", "example"),
            "resign" => array(90, "", 0, "Forces the bot to resign channel moderator status if it has it.", "example"),
            "setname" => array(90, ".setname (new username)", 0, "Changes the bot's configuration username. The next time it logs on, it will use that username to connect.", "example"),
            "setpass" => array(90, ".setpass (new password)", 0, "Changes the bot's configuration password. The next time it logs on, it will use that password to connect.", "example"),
            "setserver" => array(90, ".setserver (new server)", 0, "Changes the bot's configuration server. The next time it logs on, it will use that server to connect.", "example"),
            "quiettime" => array(90, ".quiettime (on/off/status)", 0, "Toggles and displays status of the bot's Quiet Time feature, which bans users who talk if they're not safelisted.", "example"),
            "giveup" => array(90, ".giveup (username)", 0, "Designates the target username, then resigns moderator status.", "example"),
            "readfile" => array(90, ".readfile (filename)", 0, "Reads the contents of a text file in the bot's folder.", "example"),
            "chpw" => array(90, ".chpw (on (desired password) / time (time limit to provide a password) / off / status)", 0, "Turns <b>on</b> or <b>off</b> Channel Passwording. Alternately, this command can be used to set the <b>time</b> before a user is banned if they fail to provide a password, or to set a new Channel Password of your choice.", "example"),
            "clan" => array(90, ".clan (private / public / arbitrary value)", 0, "Toggles clan channel private or public status. If pub/priv/public/private are not specified, the bot will simply send '/clan <arguments>' to Battle.net.", "example"),
            "sweepban" => array(90, ".sweepban (channel)", 0, "Attempts to ban every user in the target channel.", "example"),
            "sweepignore" => array(90, ".sweepignore (channel)", 0, "Attempts to use Battle.net to ignore every user in the target channel.", "example"),
            "idlebans" => array(90, ".idlebans (on/off/status/delay/kick)", 0, "", "example"),
            "setkey" => array(90, ".setkey (new CDKey)", 0, "Changes the bot's configuration CDKey. The next time it logs on, it will use that CDKey.", "example"),
            "setexpkey" => array(90, ".setexpkey (new CDKey)", 0, "Changes the bot's configuration Expansion CDKey. The next time it logs on, it will use that Expansion CDKey.", "example"),
            "exile" => array(90, ".exile (username) [message]", 0, "Adds the target user to the bot's shitlist, banning them. Then, IPBans the target user using Battle.net's ignore feature.", "example"),
            "unexile" => array(90, ".unexile (username)", 0, "Removes the target from the bot's shitlist and un-IPBans them.", "example"),
            "clearbanlist" => array(90, "", 0, "Clears the bot's list of banned users. This command is useful if Ban Evasion gets out of hand.", "example"),
            // 100
            "quit" => array(100, "", 0, "This command shuts down your StealthBot.", "example"),
            "locktext" => array(100, "", 0, "This command will lock the chat window of your StealthBot, preventing text from being written to it. Your bot will continue to function normally. The command is designed to reduce lag during critical times..", "example"),
            "efp" => array(100, ".efp (on/off)", 0, "This command toggles <b>on</b> or <b>off</b> your bot's Emergency Floodbot Protection [EFP] feature. For more information, see the <a href='index.php?act=viewFeature&c=efp'>EFP feature page</a>.", "example"),
            "loadwinamp" => array(100, "", 0, "This command attempts to load Winamp from its default installation directory. You can also point it at a different installation directory using a <a href='index.php?act=viewPage&c=ch'>config.ini hack</a>.", "example"),
            "setmotd" => array(100, ".setmotd (new motd)", 0, "Changes the bot's Warcraft III clan's Message of the Day. To work, the bot must be a Shaman or Chieftain in a Warcraft III clan.", "example"),
            "invite" => array(100, ".invite (username)", 0, "Invites a user to join your bot's Warcraft III clan. To work, the bot must be a Shaman or Chieftain in a Warcraft III clan.", "example"),
            "peonban" => array(100, ".peonban (on/off/status)", 0, "Toggles and displays status for your bot's Peon Ban moderation feature. If enabled, the bot will ban any Warcraft III user that joins and has a Peon icon.", "example"),
            
            );
?>