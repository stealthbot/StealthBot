<?php
	// StealthBot Version Output - Quick & Dirty
	// 2009-1006-andy
	
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
?>

