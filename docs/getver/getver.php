<?php
// (c) 2003 Stealth Networks all rights reserved
// Returns the current version code and update download URL.

function getVersion () 
{ 
	$rand = rand(0, 1);
  	$ver = "241";
  	$verurl = "x";
  	
  	$com = "If you need help or have questions feel free to join our forums at http://www.stealthbot.net. They're back up!\\n\\n";
  	
  	switch($rand)
  	{
		/***
	  	case 0:
  			$com .=
  			"Want to trade Diablo II items? Check out the new http://www.d2central.com ! D2Central holds tournaments raffles and more - all realms are welcome!";
  			break;
  			
  		case 1:
  			$com .= "ÿcbForge Hosting is the new home of StealthBot.net\\n" .
   					"Forge Hosting's lightning-fast servers can handle all of your website needs.\\n" .
   					"Plans start at just $4.99 monthly for 1GB space and 20GB transfer with 24 hour activation guaranteed. Check us out at http://www.forgehosting.com!";
   			***/		
	}
  	
		
  //$com = "The StealthBot website is currently experiencing technical difficulties. It will be back online hopefully within the next two days. Sorry for any inconvenience.\\n- Stealth";
  
  print("584171,$ver,$verurl,$com");
}

 getVersion();

 
 
 /*Forge Hosting is the new home of StealthBot.net
   Forge Hosting's lightning-fast servers can handle all of your website needs.
   Plans start at just $4.99 monthly for 1GB space and 20GB transfer with 24 hour activation guaranteed. Check us out at http://www.forgehosting.com!
 */
 
?>