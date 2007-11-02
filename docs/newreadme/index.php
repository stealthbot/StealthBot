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
    
    $act = $_GET['act'];
    if(strlen($act) == 0)
        $act = "index";
        
    $subact = $_GET['c'];
    
    $title = $TITLES[$act];
    if( strlen($title) == 0 )
        $title = "Welcome";
    else if( $title == "vp" )
    {
        require_once("pages.php");
        $title = $PAGETITLES[$subact];
    } 
       
    
    require_once("header.php"); 
    
    switch($act)
    {
        case "index": // no input
            echoln("<h2>Welcome to the StealthBot Readme!</h2>");
            echoln("<p>This readme is designed to help you learn what your StealthBot can do for you.");
            echoln("Where would you like to start?</p>");
            echoln("<ul>");
                echoln("<li>Learn how to <a href='index.php?act=viewPage&c=setup'>set up your StealthBot</a></li>");
                echoln("<li>Browse a list of the <a href='index.php?act=featureList'>bot's features</a></li>");
                echoln("<li>Check out the bot's <a href='index.php?act=cmdList'>numerous commands</a></li>");
                echoln("<li>Learn how to use the bot's <a href='index.php?act=viewPage&c=ccs'>Custom Command system</a></li>");
                echoln("<li>Learn the more <a href='index.php?act=viewPage&c=adv'>advanced tweaks</a> you can use with StealthBot</li>");
                echoln("<li>Have a look at the power of the bot's built-in <a href='index.php?act=viewPage&c=script'>VBScripting system</a></li>");
                echoln("<li>Learn how you can <a href='index.php?act=viewPage&c=donate'>donate to the project</a>, and what you get for it! ;)</li>");
            echoln("</ul>");
            echoln("<p>I'm always open to your comments and suggestions. You can <a href='mailto:stealth@stealthbot.net?subject=StealthBot'>e-mail them to me</a> or post them in <a href='http://www.stealthbot.net'>the forums</a>. ");
            echoln("Also, this readme system is for sale. E-mail me if you're interested!</p>");
            echoln("<p>Enjoy!<br />- Andy T (Stealth@USEast)</p>");
            // TODO: Insert paypal donate link
            break;
        
        case "viewPage": // general pages
            require_once("pages.php");
            
            echo $PAGES[$subact];
            break;
            
        case "featureList": // feature list
            require_once("features.php");
            echoln("<h2>Program Features</h2>");
            printFeatureList();
            break;
            
        case "viewFeature": // detail on a feature
            require_once("features.php");
            echoln("<h2>$subact</h2>");
            printFeatureDetail($subact);
            echoln("<p><a href='index.php?act=featureList'>Back to the list of features?</a></p>");
            break;
            
        case "cmdList": // table - list of commands and basic info
            require_once("data.php");
            printCommandList();
            break;
            
        case "viewCmd":
            require_once("data.php");
            printCommandDetail($subact);
            echoln("<p><a href='index.php?act=cmdList'>Back to the list of commands?</a></p>");
            // print command detail
            break;
            
        default:
            echoln("<h2>Invalid action</h2><br />");
            echoln("Please press the Back button in your browser and try your link again.");
    }
    
    require_once("footer.php");
?>