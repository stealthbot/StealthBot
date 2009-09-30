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
		echo "http://fiftytoo.no-ip.com:45000/zipped/Stealthbot-2.7_Build-412.zip";
		break;
 }

?>