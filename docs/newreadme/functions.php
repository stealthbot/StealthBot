<?php
    function echoln($str)
    {
        echo $str . "\n";   
    }
    
    function openTable($align = "left")
    {
        echoln('<table align="center" width="750" cellpadding="15" cellspacing="0" bgcolor="#FFFFFF" style="border:1px solid #000">');
        echoln("<tr><td align='$align'>");
    }
    
    function closeTable()
    {
        echoln("</td></tr>");
        echoln("</table><br />");
    }
    
    function printCommandList()
    {
        global $COMMANDS;
        global $CMD_DESC;
        global $ALIASES;
        global $ACCESSLEVELS;
        
        $color[0] = "#CCCCCC";
        $color[1] = "#FFFFFF";
        $i = 0;
        
        // intro
        echoln("<h2>StealthBot Commands</h2>");
        echoln("<p>Here is the full list of all " . count($COMMANDS, COUNT_RECURSIVE) . " StealthBot commands. You can use them inside the bot by typing a slash, then the command:</p>");
        echoln("<p><b>/ban anAnnoyingPerson</b></p>");
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
        
        foreach($COMMANDS as $thisLevel)
        {
            $j++;
            
            if($j == 3)
            {
                echoln("</tr><tr valign='top'>");
                $j = 1;
            }
            
            echoln("<td>");    
            echoln("<table cellpadding='4' cellspacing='2' style='border: 1px solid #000000'>");
            echoln("<caption>Access level " . $ACCESSLEVELS[$n] . "</caption>");
            echoln("<tr class='headerrow'><td>Command Word</td><td>Access</td><td>Alias(es)</td></b></tr>");
            $i = 0;
            
            foreach($thisLevel as $c)
            {
                echoln("<tr bgcolor='" . $color[$i] . "'>");
                echoln("<td><a href='index.php?act=viewCmd&c=$c'>" . $c . "</a></td>");
                echoln("<td>" . $ACCESSLEVELS[$n] . "</td>");
                echoln("<td>");
                
                if(count($ALIASES[$c]) > 0)
                {
                    $aOut = "";
                    
                    foreach($ALIASES[$c] as $a)
                        $aOut .= $a . ", ";
                    
                    $aOut = substr($aOut, 0, strlen($aOut) - 2);
                    echoln($aOut);
                }
                else
                    echoln("-");
                    
                echoln("</td>");           
                
                if($i == 0)
                    $i = 1;
                else
                    $i = 0;
            }
            
            echoln("</table>");
            echoln("</td>");
            
            $n++;
        } 
        
        echoln("</tr></table>");
    }
    
    function printFeatureList()
    {
        global $FEATURES;
        global $FEATURE_DESC;
        
        echoln("<ul>");
        
        foreach($FEATURES as $f)
        {
            echoln("<li><a href='index.php?act=viewFeature&c=$f'>" . $FEATURE_DESC[$f][0] . "</a></li>");
        }
        
        echoln("</ul>");
    }
    
    function printCommandDetail($cmd)
    {
        global $CMD_DESC;
        global $ALIASES;
        
        $c = $CMD_DESC[$cmd];
        
        echoln("<h2>$cmd (" . $c[0] . ")</h2>");
        echoln("<p><b>Alias(es)</b><br />");
        
        if(isset($ALIASES[$cmd]))
        {
            $aOut = "";
                foreach($ALIASES[$cmd] as $a)
                    $aOut .= $a . ", ";
            
            echoln(substr($aOut, 0, strlen($aOut)-2));
        }
        else
            echoln("None");
             
        echoln("</p>");
        
        echoln("<p><b>Description</b><br />" . $c[3] . "</p>");
        echoln("<p><b>Arguments</b><br />");
        echoln("<i>optional arguments are in [square brackets], required arguments are in (parentheses)</i><br><br>");
        if( strlen($c[1]) == 0 )
            echoln("None");
        else
            echoln($c[1]);
        
        if( $c[2] == 1)
            echoln(" (supports wildcards)");
        
        echoln("</p>");
        echoln("<p><b>Usage Example</b><br />" . $c[4]."</p>");
    }
    
    function printFeatureDetail($feat)
    {
        echo "This page isn't ready yet. Sorry! $feat";
    }
?>              