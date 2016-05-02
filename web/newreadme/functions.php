<?php
    // echoln
    // Echoes a single line with line break for pretty source code
    // Future: skinnable
    // Converted: N/A
    function echoln($str)
    {
        echo $str . "\n";   
    }
    
    // openTable
    // Opens a UI table
    // Future: skinnable
    // Converted: Yes
    function openTable($align = "left")
    {
        echoln('<table align="center" width="750" cellpadding="15" cellspacing="0" bgcolor="#FFFFFF" style="border:1px solid #000">');
        echoln("<tr><td align='$align'>");
    }
    
    // closeTable
    // Closes a UI table
    // Future: skinnable
    // Converted: Yes
    function closeTable()
    {
        echoln("</td></tr>");
        echoln("</table><br />");
    }
    
    // print404
    // Displays an error to the user with a dead link reporting method
    // Future: skinnable
    // Converted: Yes
    function print404()
    {
        startPage("Error");
        
        echoln("<h2>Invalid action</h2><br />");
        echoln("Please press the Back button in your browser and try your link again.");
        echoln("");
        echoln("If the problem persists, please <a href='mailto:stealth@stealthbot.net?subject=Dead readme link&body=" . $_SERVER['REQUEST_URI'] . "'>let me know</a>.");
    }
    
    // getPage
    // Grabs a page from the database and spits it out
    // Future: skinnable
    // Converted: Yes
    function getPage($page)
    {
        global $db;
        
        $page = mysql_escape_string($page);
        $r = $db->doQuery("SELECT pageName, pageTitle, pageContents FROM pages WHERE pageName='$page';");
        if( mysql_num_rows($r) > 0 )
        {
            $row = mysql_fetch_row($r);
            
            startPage($row[1]);
            
            echoln("<h2>" . $row[1] . "</h2>");
            echoln(stripslashes($row[2]));  
        }
        else   
            print404();
    }
    
    // getCommandCount
    // Retrieves the current count of commands in the system
    // Converted: Yes
    function getCommandCount()
    {
        global $db;
        
        $ret = $db->doQuery("SELECT COUNT(cid) FROM commands;");
        $r = mysql_fetch_row($ret);
        return $r[0];
    }
    
    // printCommandList
    // Presents the user with a list of commands in the system linking to their detail
    // Future: skinnable
    // Converted: Yes
    function printCommandList()
    {
        global $COMMANDS;
        global $CMD_DESC;
        global $ALIASES;
        global $ACCESSLEVELS;
        global $db;
        
        startPage("Command List");
        
        $color[0] = "#CCCCCC";
        $color[1] = "#FFFFFF";
        $i = 0;
        
        // intro
        echoln("<h2>StealthBot Commands</h2>");
        echoln("<p>Here is the full list of all " . getCommandCount() . " StealthBot commands. You can use them inside the bot by typing a slash, then the command:</p>");
        echoln("<p><b>/ban anAnnoyingPerson</b></p>");    
        echoln("<p>You can use them inside the bot, but have their results displayed publicly to the channel you're in, by typing two slashes before the command:</p>");
        echoln("<p><b>//mp3<br />&lt;StealthBot&gt; Current MP3: 1418: Lamb of God - Again We Rise</b></p>");
        echoln("<p>You can also use them from outside the bot, if you have enough access:</p>");
        echoln("<p><b>&lt;Stealth&gt; .say Hello, world!<br />&lt;StealthBot&gt; Hello, world!</b></p>");
        echoln("<p>If you want to chain together multiple commands, you can do that also:</p>");
        echoln("<p><b>&lt;Stealth&gt; .whoami; say Hello, world!<br />&lt;StealthBot&gt; You have access 999 and flags ADMOST.<br>&lt;StealthBot&gt; Hello, world!</b></p>");
        echoln("<p>Lastly, if more than one of your bots share the same trigger, you can order them around by name:</p>");
        echoln("<p><b>&lt;Stealth&gt; .StealthBot say Hello, world!<br />&lt;StealthBot&gt; Hello, world!</b></p>");
        echoln("<p>Click on a command's name to see a detailed description of it, and an example of its use.</p>");
        
        // begin table output        
        echoln("<table width='600' cellpadding='1' cellspacing='15'>");
        echoln("<tr valign='top'>");
        $j = 0;
        $n = 0;
        $thisAccess = 0;
        
        $ret = $db->doQuery("SELECT access, cword, cid FROM commands ORDER BY access ASC;");
        
        // loop through commands, listing and classifying them appropriately
        while( $r = mysql_fetch_row($ret) )
        {
            $thisCID = $r[2];
            
            // alternate row colors
            if($i == 0)
                $i = 1;
            else
                $i = 0;
                
            if($r[0] != $thisAccess)
            {
                $thisAccess = $r[0]; 
                $i = 0;
                
                if( $j > 0 )
                    closeCLMiniTable();
                    
                $j++;
            
                if($j == 3)
                {
                    echoln("</tr><tr valign='top'>");
                    $j = 1;
                }
                    
                openCLMiniTable($thisAccess);
            }
            
            // print this command with a link to detail
            echoln("<tr bgcolor='" . $color[$i] . "'>");
            echoln("<td><a href='index.php?act=viewCmd&c=" . $r[1] . "'>" . $r[1] . "</a></td>");
            echoln("<td>" . $thisAccess . "</td>");
            echoln("<td>");
            
            // print any aliases the command has 
            $aliasret = $db->doQuery("SELECT alias FROM aliases WHERE cid='$thisCID';");
            
            if(mysql_num_rows($aliasret) > 0)         
            {
                $aOut = "";
                
                while($aliases = mysql_fetch_row($aliasret))
                    $aOut .=  (strlen($aOut)>0 ? ", " : "") . $aliases[0];
                        
                echoln($aOut);
            }
            else
                echoln("-");
                
            echoln("</td>");           
            
            $n++;
        } 
        
        closeCLMiniTable();
        
        echoln("</tr></table>");
    }
    
    // openCLMiniTable
    // Opens up a new command list table
    // Future: skinnable
    // Converted: Yes
    function openCLMiniTable($access)
    {
        echoln("<td>");    
        echoln("<table cellpadding='4' cellspacing='2' style='border: 1px solid #000000'>");
        echoln("<caption>Access level " . $access . "</caption>");
        echoln("<tr class='headerrow'><td>Command Word</td><td>Access</td><td>Alias(es)</td></b></tr>");
    }
    
    // closeCLMiniTable
    // Closes a command list table
    // Future: skinnable
    // Converted: Yes
    function closeCLMiniTable()
    {
        echoln("</table>");
        echoln("</td>");
    }
    
    // printFeatureList
    // Prints a list of features from the feature db
    // Future: skinnable
    // Converted: No
    function printFeatureList()
    {
        global $db;
        
        startPage("Feature List");
        
        echoln("<ul>");
        
        foreach($FEATURES as $f)
        {
            echoln("<li><a href='index.php?act=viewFeature&c=$f'>" . $FEATURE_DESC[$f][0] . "</a></li>");
        }
        
        echoln("</ul>");
    }
    
    // printCommandDetail
    // Shows detailed information on the selected command
    // Future: skinnable
    // Converted: Partial
    function printCommandDetail($cmd)
    {
        global $db;
        
        $cmd = mysql_escape_string($cmd);
        startPage("Command: " . $cmd);
        
        $cmdret = $db->doQuery("SELECT * FROM commands WHERE cword='$cmd';");
        
        $c = mysql_fetch_array($cmdret, MYSQL_ASSOC);
        $cmdID = $c['cid'];
        
        if($cmdID > 0)
        {
            $aliasret = $db->doQuery("SELECT alias FROM aliases WHERE cid='" . $cmdID . "';");
            
            echoln("<h2>$cmd (" . $c['access'] . ")</h2>");
            echoln("<p><b>Alias(es)</b><br />");
            
            if(mysql_num_rows($aliasret)>0)
            {
                while($aliases = mysql_fetch_row($aliasret))
                    $aOut .=  (strlen($aOut)>0 ? ", " : "") . $aliases[0];
                        
                echoln($aOut);
            }
            else
                echoln("None");
                 
            echoln("</p>");
            
            echoln("<p><b>Description</b><br />" . stripslashes($c['descr']) . "</p>");
            echoln("<p><b>Arguments</b><br />");
            //echoln("<i>optional arguments are in [square brackets], required arguments are in (parentheses)</i><br><br>");
            
            if( intval($c['wildcardable']) == 1)
                echoln("(This command supports <a href='index.php?act=viewPage&c=wildcards'>wildcards</a>)");  
            
            // TODO: GET ARGUMENTS FROM DB
            $argret = 
              $db->doQuery("SELECT argument,explanation,optional FROM arguments WHERE cid='$cmdID';");
            
            if( mysql_num_rows($argret) > 0 )
            {
                $aOut = "<ul>";
                
                while($arguments = mysql_fetch_row($argret))
                {
                    $optional = intval($arguments[2]);
                    $explain = stripslashes($arguments[1]);
                    $argument = stripslashes($arguments[0]);
                                              
                    $aOut .= "<li><b>$argument</b>";
                    
                    if( $optional == 1 )
                        $aOut .= " (optional)";
                    
                    $aOut .= "<br />$explain";
                }
                
                echoln($aOut . "</ul>");
            }
            else
                echoln("<ul><li>No arguments</ul>");                       
                                                 
            echoln("</p>");
            echoln("<p><b>Usage Example</b><br />");
            
            if($c['cusage'] == "example")
                echoln("(no example for this command)");
            else
                echoln(stripslashes($c['cusage'])."</p>");
        }
        else
            echoln("<b>That's not a valid command!</b>");
        
        echoln("<p><a href='index.php?act=cmdList'>Back to the list of commands?</a></p>"); 
    }
    
    // printFeatureDetail
    // Pulls and prints details on the selected feature
    // Future: skinnable
    // Converted: Yes
    function printFeatureDetail($feat)
    {
        startPage("Under Construction");
        echo "This page isn't ready yet. Sorry! $feat";
        
        echoln("<p><a href='index.php?act=featureList'>Back to the list of features?</a></p>");
    }
    
    // startPage
    // Allows individual pages to set the page title
    // Converted: Yes
    function startPage($title)
    {
        static $started;
        
        // prevent multiple page starts
        if($started == true)
            return;
        else
            $started == true;
            
        // Start drawing page
        require_once("header.php");    
    }
    
    // printTeam
    // Displays the dev team bios page
    function printTeam()
    {
        global $db;
        $align = 0;
        
        startPage("Meet the StealthBot Team");
        
        echoln("<h2>Meet the Team</h2>");
        echoln("<p>As of September 24th, 2007, there are five StealthBot developers. Their bios are listed below. <b>This page is under construction and far from finished :)</b></p><br />");
        
        $ret = $db->doQuery("SELECT * FROM dev;");
        
        // Field format for this database:
        //  0       1       2       3       4
        // devid    name    loc     imgurl  bio
        while($row = mysql_fetch_row($ret))
        {
            echoln("<h3>" . $row[1] . "</h3>");
            echoln("<p>
                    <a href='http://www.stealthbot.net/board/index.php?showuser=" . $row[4] . "'>
                    <img src='" . $row[3] . "' align='" . ($align == 0 ? "right" : "left") . "'>
                    </a>");
            echoln("\t" . stripslashes($row[5]));
            echoln("</p>");
            
            if($align==0)
                $align = 1;
            else
                $align = 0;
            
            echoln("<br />");
        }
    }
    
?>              