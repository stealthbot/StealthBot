<?php
    // (c) 2006 Stealth (andy@stealthbot.net)
    $vc = $_GET['vc'];

    // current vercode | regular news | beta news | cvstring | beta cvstring
    echo "264|";

    switch($vc)
    {
        case "2699":
            echo "This version is a Release Candidate. Test it thoroughly, in particular recent topics moved to the Beta Development Archive.";
            echo "|";
            echo "-";
            break;
            
        case "2696":
            echo "This version is old, get a new one!";
            echo "|";
            echo "-";
            break;
            
        case "269":
            echo "This version is a Release Candidate. Test it thoroughly -- you know the drill.";
            echo "|";
            echo "-";
            break;
            
        case "268":
            echo "This version is Release Candidate 2. Or, because I forgot to update the vercode in 2.692, it might be Release Candidate 3. Either way, make sure you're up-to-date ;)";
            echo "|";
            echo "-";
            break;
            
        case "267":
            echo "This version is Release Candidate 1.";
            echo "|";
            echo "Update your beta!";
            break;

        case "264":
        default:
            echo "Welcome to the new StealthBot News! Blah blah blah.";
            echo "|";
            echo "Update your betas, bitches.";

        break;

    }
    
    echo "|2.6R3|2.6991";  
?>