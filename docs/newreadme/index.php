<?php
    /* 
    *   NewReadme for StealthBot
    *       December 2005 by Andy T (Stealth)
    *       Online readme and FAQ system
    *       Main file
    */

    define("__READMEFILE", true);
    require_once("functions.php");
    require_once("globals.php");
    require_once("Database.php");
    
    $act = $_GET['act'];
    if(strlen($act) == 0)
        $act = "index";
        
    $subact = htmlentities(mysql_escape_string($_GET['c']));       
    
    $db = new DBHandler();
    
    switch($act)
    {
        case "index": // no input
            getPage("index");
            // TODO: Insert paypal donate link
            break;
                  
        case "viewPage": // general pages
            getPage($subact);
            break;
            
        case "featureList": // feature list
            printFeatureList();
            break;
            
        case "viewFeature": // detail on a feature
            printFeatureDetail($subact);
            break;
            
        case "cmdList": // table - list of commands and basic info
            printCommandList();
            break;
            
        case "viewCmd":
            printCommandDetail($subact);
            // print command detail
            break;
        
        case "team":
            printTeam();
            break;
            
        default:
            print404();
    }
    
    require_once("footer.php");
?>