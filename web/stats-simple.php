<?php

    class logEntry {
        var $date;
        var $time;
        var $ip;
        var $cv;
        var $lv;
    }
    
    class versionStat {
        var $version;
        var $total;
        
        var $unique;
    }
    
    function createLogEntry($rawLine) {
        $a = explode(" ", trim($rawLine));
        if (count($a) != 5) return null;
        
        $entry = new logEntry();
        $entry->date = $a[0];
        $entry->time = $a[1];
        $entry->ip   = $a[2];
        $entry->cv   = $a[3];
        $entry->lv   = $a[4];
        
        return $entry;
    }
    
    $entries = array();
    $unique = array();
    $versions = array();
    
    $firstDate = null;
    
    $log = file("version.log");
    if (isset($log)) {
        foreach ($log as $item) {
            array_push($entries, createLogEntry($item));
        }
    } else {
        die("Log file missing.");
    }
    
    foreach ($entries as $e) {
        if (!($e == null)) {
            if (!isset($firstDate)) {
                $firstDate = $e->date;
            }
        
            if (!in_array($e->ip, $unique)) {
                array_push($unique, $e->ip);
            }
        
            if (!array_key_exists($e->cv, $versions)) {
                $vs = new versionStat();
                $vs->version = $e->cv;
                $vs->total = 0;
                $vs->unique = array();
            
                $versions[$e->cv] = $vs;
            }
        
            $versions[$e->cv]->total++;
        
            if (!in_array($e->ip, $versions[$e->cv]->unique)) {
                array_push($versions[$e->cv]->unique, $e->ip);
            }
        }
    }
    
?>
<html>
    <head>
        <title>StealthBot Usage Statistics</title>
        
        <style>
            td {
                text-align: center;
                padding: 10px;
            }
        </style>
    </head>
    
    <body>
        Start date: <?php echo $firstDate . " [" . date_diff(new DateTime($firstDate), new DateTime())->d . " days ago]"?> <br />
        <table border="1" cellspacing="0" width="25%">
            <tr><th>Version</th><th>Unique Users</th><th>Total Uses</th></tr>
<?php
    
    echo "\t\t\t<tr><td>ALL</td><td>" . count($unique) . "</td><td>" . count($entries) . "</td></tr>\r\n";
    foreach ($versions as $v) {
        echo "\t\t\t<tr><td>" . $v->version . "</td><td>" . count($v->unique) . "</td><td>" . $v->total . "</td></tr>\r\n";
    }
    
?>
        </table>
    </body>
</body>