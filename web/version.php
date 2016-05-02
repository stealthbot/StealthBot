<?php
	// StealthBot Version Output - Quick & Dirty
	// 2009-1006-andy
	
    // set this to false to disable logging
    $enableLogging = true;
    
    $cv = 0;
    $lv = 0;
    
    if (isset($_GET["cv"])) $cv = $_GET["cv"];
    if (isset($_GET["lv"])) $lv = $_GET["lv"];
    
	$beta_build = 489;
	$regular_build = 1;
	$launcher_build = 0x01020000 + 1;
	$regular_news = str_replace("\n","\\n", trim("
Welcome to the long-awaited StealthBot 2.7. If you have problems with your bot,
visit us at http://www.stealthbot.net and post in the forums. You can also check
out our Wiki at http://www.stealthbot.net/wiki/. 

ÿcbOctober 18th 2009ÿcb
ÿcbStealthBot 2.7.1ÿcb has been released, this includes a lot of bug fixes. 
Visit http://www.stealthbot.net for more information.
	"));

	$beta_news = str_replace("\n","\\n", trim("
It's 2016 and a new build of StealthBot has been released!
You can get the latest version of StealthBot at http://stealthbot.net/sb/releases/
	"));

	echo "$beta_build|$regular_build|$launcher_build|$regular_news|$beta_news";
    
    
    // request logging, for statistical purposes - added 2016-03-22 by Pyro
    // records the following
    //   - date+time (in server time, currently America\Los_Angeles)
    //   - sha1 of remote IP
    //   - CVERSION
    //   - launcher version
    // CVERSION and launcher version are as reported by the bot.
    if ($enableLogging && isset($_GET["cv"])) {
        error_reporting(0);
        
        $fileName = "version.log";

        if (($file = @fopen($fileName, "a")) != false) {
            @fwrite($file, implode(" ", array(date("Y-m-d H:i:s"), sha1($_SERVER['REMOTE_ADDR']), $cv, $lv, "\r\n")));
            @fclose($file);
        }
    }
?>

