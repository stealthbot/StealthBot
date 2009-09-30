<?php

/**********************************
 *Blah?
 **********************************/
 
 switch($_GET["p"]){
	case "lupdate": 
		echo "1.2\xFF00000000";
		break;
		
	case "linfo":
		echo "Input: {$_GET["crc"]}";
		break;
		
	case "latest_url":
		echo "http://fiftytoo.no-ip.com:45000/zipped/Stealthbot-2.7_Build-421.zip";
		break;
		
	case "lnews":
		echo "September 30th 2009 - Hdx\r\n";
		echo "Welcome to the new StealthBot Launcher\r\n";
		echo "This allows you to run multiple instances of SB from the same executable.\r\n";
		echo "This also makes all files used by the bot UAC compliant.\r\n";
		echo "Your Bot's information will now be located in {PROFILEPATH}\r\n";
		break;
 }

?>