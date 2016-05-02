<?php
// (c) 2004 Stealth Networks all rights reserved
// Returns the current version code and update download URL.

function getVersion () 
{ 
	$rnd = rand(0, 1);
  	$ver = "261";
  	
  	$news = "---------------------\\nÿcbBattle.net Changes\\n---------------------\\nBlizzard has made an addition version check system change.\\n\\nThere is no way for you to patch this one at the time.\\nI am working on continuing to prepare the 2.7 release.\\nI will do my best. Thanks for your patience, and stay tuned to\\nhttp://www.stealthbot.net for updates.";

    $betanews = "This is the old beta news. Pick up a newer copy!";

  	switch($rnd)
  	{
 /*       case 0:
            $news .= "ÿcbShockware.net Gaming is the #1 source for your Blizzard games!\\nShockware Gaming is looking for experienced gaming moderators, check us out and apply!\\nShockware has templates, bot files, and more. We also support all Blizzard games!\\nCheck us out at http://www.shockware.net";

            break;



  		case 1:
  			$news .= "ÿcbForge Hosting is the new home of StealthBot.net\\n" .
   					 "Forge Hosting's lightning-fast servers can handle all of your website needs.\\n" .
   					 "Plans start at just \$4.99 monthly for 1GB space and 20GB transfer with 24 hour activation guaranteed. Check us out at http://www.forgehosting.com!";
   			break;
	*/
	}

  print("22194|$ver|$news|$betanews");
}

 getVersion();

?>