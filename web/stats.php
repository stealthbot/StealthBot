<?php

// Written by Ribose in March 2016

ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);
date_default_timezone_set('EST');

const ROW_TOTAL = 'Total';
const ROW_LUSED = 'Launcher';
const ROW_LNO   = 'Standalone';

interface iGetVersion {
    public function get_version();
}

final class LogEntry implements iGetVersion {
    private $date; // [ XXXX-XX-XX ]
    private $time; // [ XX:XX:XX ]
    private $ip; // IP [ SHA1 ]
    private $cv; // current version [ 2.7.# | Beta 2.7 Build # ]
    private $lv; // launcher version [ 0 | hex value 0x1020XX | hex value 0x103000 ]
    
    // in: 5 values from log
    public function __construct($date, $time, $ip, $cv, $lv) {
        $this->date = $date;
        $this->time = $time;
        $this->ip = $ip;
        $this->cv = $cv;
        $this->lv = $lv;
    }
    
    // out: "date time" integer (UNIX timestamp)
    public function get_datetime() {
        return strtotime("$this->date $this->time");
    }
    
    public function get_date() {
        return strtotime("$this->date 00:00:00");
    }
    
    // out: "ip" SHA1 string
    public function get_ip() {
        return $this->ip;
    }
    
    // out: "version" string
    public function get_version() {
        $v = intval($this->cv);
        if ($v > 100) {
            return "Build $v";
        } else {
            return "2.7.$v";
        }
    }
    
    // out: "launcherversion" string
    public function get_launcher_version() {
        $v = intval($this->lv);
        if ($v > 0) {
            $ma = ($v >> 24) & 0x00ff;
            $mi = ($v >> 16) & 0x00ff;
            $rv = ($v      ) & 0xffff;
            return "$ma.$mi.$rv";
        } else {
            return ROW_LNO;
        }
    }
    
    // out: whether launcher was used, boolean
    public function was_launcher_used() {
        return (intval($this->lv) != 0);
    }
}

final class VersionTableRow implements iGetVersion {
    private $table_only_row; // boolean, whether to only show this in "table" outputs
    
    private $version; // version string
    private $total; // total using this version
    
    private $uniques; // array of unique user "ip" strings using this version
    
    // in: version for this row
    // sets other private values to 0/empty
    public function __construct($version, $table_only_row) {
        $this->table_only_row = $table_only_row;
        $this->version = $version;
        $this->total = 0;
        $this->uniques = array();
    }
    
    // out: "version" string
    public function get_version() {
        return $this->version;
    }
    
    // out: total, integer
    public function get_total() {
        return intval($this->total);
    }
    
    // out: uniques, integer
    public function get_uniques() {
        return count($this->uniques);
    }
    
    public function is_table_only_row() {
        return $this->table_only_row;
    }
    
    // in: ip string to check/increment unique counter for this row
    // always increments total counter for this row
    public function increment_with($ip_val) {
        $this->total++;
        
        if (isset($this->uniques[$ip_val])) {
            $this->uniques[$ip_val]++;
        } else {
            $this->uniques[$ip_val] = 1;
        }
    }
}

// comparison function for iVersionObj
function iGetVersion_cmp($a, $b) {
    if ($a instanceof iGetVersion && $b instanceof iGetVersion) {
        return strcmp($a->get_version(), $b->get_version());
    } else {
        return 0;
    }
}

$entries_v = file("version.log");
if (isset($entries_v)) {
    $entries_c = count($entries_v);
    for ($i = 0; $i < $entries_c; $i++) {
        $ex = explode(' ', trim($entries_v[$i]));
        if (count($ex) == 5) {
            list($date,$time,$ip,$cv,$lv) = $ex;
            $entries_v[$i] = new LogEntry($date,$time,$ip,$cv,$lv);
        } else {
            unset($entries_v[$i]);
        }
    }

    $entries_v = array_values($entries_v); // array of actual entries (LogEntry)
} else {
    $error = "Log file missing.";
    $entries_v = array();
}

$vers_v = array(ROW_TOTAL => new VersionTableRow(ROW_TOTAL, true)); // array of versions ('ver'=>VersionTableRow)
$lvers_v = array(ROW_LUSED => new VersionTableRow(ROW_LUSED, true)); // array of launcher versions ('ver'=>VersionTableRow)
$vers_v_by_date = array(); // array of "dated" columns (date => ['ver'=>VersionTableRow, 'ver2'=>VersionTableRow])
foreach ($entries_v as $entry) {
    //echo '<pre>'.var_dump($entry).'</pre>';
    
    // use IP value for uniques
    $ip_val = $entry->get_ip();
    
    // populate version table with this row
    $ver_val = $entry->get_version();
    if (!isset($vers_v[$ver_val])) {
        $vers_v[$ver_val] = new VersionTableRow($ver_val, false);
    }
    $vers_v[$ver_val]->increment_with($ip_val);
    // special row: TOTAL
    $vers_v[ROW_TOTAL]->increment_with($ip_val);
    
    // populate version-by-date table
    $date_val = $entry->get_date();
    if (!isset($vers_v_by_date[$date_val])) {
        $vers_v_by_date[$date_val] = array();
    }
    if (!isset($vers_v_by_date[$date_val][$ver_val])) {
        $vers_v_by_date[$date_val][$ver_val] = new VersionTableRow($ver_val, false);
    }
    $vers_v_by_date[$date_val][$ver_val]->increment_with($ip_val);
    
    // populate launcher table with this row
    $lver_val = $entry->get_launcher_version();
    if (!isset($lvers_v[$lver_val])) {
        $lvers_v[$lver_val] = new VersionTableRow($lver_val, false);
    }
    $lvers_v[$lver_val]->increment_with($ip_val);
    // special row: LAUNCHER USED
    if ($entry->was_launcher_used()) {
        $lvers_v[ROW_LUSED]->increment_with($ip_val);
    }
}

ksort($vers_v);
ksort($lvers_v);

function get_version_row_data($table, $row_name) {
    if (isset($table[$row_name])) {
        $ver = addslashes($row_name);
        $cnt = $table[$row_name]->get_total();
        $uqs = $table[$row_name]->get_uniques();
    } else {
        $ver = addslashes($row_name);
        $cnt = 0;
        $uqs = 0;
    }
    echo "'$ver', $cnt, $uqs";
}

function get_version_table_data($table) {
    $co = '';
    foreach ($table as $version_table_row) {
        $ver = addslashes($version_table_row->get_version());
        $cnt = $version_table_row->get_total();
        $uqs = $version_table_row->get_uniques();
        if ($cnt > 0 && !$version_table_row->is_table_only_row()) {
            echo $co;
            echo "['$ver',$cnt,$uqs]";
            $co = ',';
        }
    }
}

function get_dated_version_column_data($table) {
    $co = ',';
    $columns = array();
    foreach ($table as $version_table_row) {
        $ver = addslashes($version_table_row->get_version());
        $cnt = $version_table_row->get_total();
        $uqs = $version_table_row->get_uniques();
        if ($cnt > 0 && !$version_table_row->is_table_only_row()) {
            echo "$co{'type': 'number', 'caption': '$ver'}";
            $columns[] = $ver;
            $co = ',';
        }
    }
    return $columns;
}

function get_dated_version_table_data($table, $columns) {
    $co = '';
    foreach ($table as $date_key => $date_table) {
        echo $co;
        $date_val = $date_key*1000;
        echo "[new Date($date_val)";
        $co2 = ',';
        foreach ($columns as $version) {
            echo $co2;
            if (isset($table[$date_key][$version])) {
                echo $table[$date_key][$version]->get_total();
            } else {
                echo '0';
            }
        }
        echo "]";
        $co = ',';
    }
}

//echo '<html><body><pre>';
//var_dump($vers_v);
//echo '</pre></body></html>';
//exit;

?>
<html>
    <head>
        <title>StealthBot Usage Statistics</title>

<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
<script type="text/javascript">
var serverData_Versions = {
    "columns": [
        {'type': 'string', 'caption': 'StealthBot Version'},
        {'type': 'number', 'caption': 'Total'},
        {'type': 'number', 'caption': 'Unique'}
    ],
    "table_only_rows": [
        [<?php get_version_row_data($vers_v, ROW_TOTAL); ?>]
    ],
    "data": [<?php get_version_table_data($vers_v); ?>]
};

var serverData_TVersionsPerDay = {
    "columns": [
        {'type': 'date', 'caption': 'Date'}<?php $col_v = get_dated_version_column_data($vers_v); ?>
    ],
    "table_only_rows": [],
    "data": [<?php get_dated_version_table_data($vers_v_by_date, $col_v); ?>]
};

var serverData_LVersions = {
    "columns": [
        {'type': 'string', 'caption': 'Launcher Version'},
        {'type': 'number', 'caption': 'Total'},
        {'type': 'number', 'caption': 'Unique'}
    ],
    "table_only_rows": [
        [<?php echo get_version_row_data($lvers_v, ROW_LUSED); ?>]
    ],
    "data": [<?php echo get_version_table_data($lvers_v); ?>]
};

// packages
google.charts.load('current', {packages: ['corechart', 'table']});

function mkDataTableFromServerData(isTable, columnsChosen, serverData) {
    var data = new google.visualization.DataTable();
    
    for (var i = 0; i < serverData.columns.length; i++) {
        if (columnsChosen == null || columnsChosen.indexOf(i) >= 0) data.addColumn(serverData.columns[i].type, serverData.columns[i].caption);
    }
    
    var filterRows = function (row) {
        var newRow = [];
        for (var i = 0; i < row.length; i++) {
            if (columnsChosen.indexOf(i) >= 0) {
                newRow.push(row[i]);
            }
        }
        return newRow;
    };
    
    if (columnsChosen == null) {
        filterRows = function (row) { return row; }
    }
    
    data.addRows(serverData.data.map(filterRows));
    
    if (isTable) {
        data.addRows(serverData.table_only_rows.map(filterRows));
    }
    
    return data;
}


function loadChart(chartServerData, chartType, chartColumns, dom_element, options) {
    google.charts.setOnLoadCallback(function () {
        var data = mkDataTableFromServerData(chartType == "Table", chartColumns, chartServerData);
        var chart = new google.visualization[chartType](document.getElementById(dom_element));
        chart.draw(data, options);
    });
}

// client
loadChart(serverData_Versions, 'ColumnChart', [0,1,2], 'sbnewslog_v',
          {
              title: 'StealthBot Version Usage',
              //vAxis: {scaleType: 'log'},
              legend: {position: 'top'}
          });
loadChart(serverData_Versions, 'PieChart', [0,1], 'sbnewslog_vto',
          {
              title: 'StealthBot Version Usage (Total)',
              is3D: true,
              sliceVisibilityThreshold: 0,
              legend: {position: 'right'}
          });
loadChart(serverData_Versions, 'Table', null, 'sbnewslog_vtbl', {});
loadChart(serverData_TVersionsPerDay, 'LineChart', null, 'sbnewslog_vpd',
          {
              title: "StealthBot Version Usage (Per Day)",
              //vAxis: {scaleType: 'log'},
              pointSize: 5,
              legend: {position: 'right'}
          });

// launcher
loadChart(serverData_LVersions, 'ColumnChart', [0,1,2], 'sbnewslog_lv',
          {
              title: 'Launcher Version Usage',
              //vAxis: {scaleType: 'log'},
              legend: {position: 'top'}
          });
loadChart(serverData_LVersions, 'PieChart', [0,1], 'sbnewslog_lvto',
          {
              title: 'Launcher Version Usage (Total)',
              is3D: true,
              sliceVisibilityThreshold: 0,
              legend: {position: 'right'}
          });
loadChart(serverData_LVersions, 'Table', null, 'sbnewslog_lvtbl', {});

</script>

<style>
td {
    text-align: center;
    padding: 5px;
}
.error {
    color: red;
}
</style>
    </head>
    
    <body>
        <?php if (isset($error)) echo "<h2 class=\"error\">Error: $error</h2>"; ?>
        <h1>StealthBot 2.7 Versions</h1>
        <div id="sbnewslog_vtbl"  style="width: 250px; height: 250px; display: inline-block;"></div>
        <div id="sbnewslog_v" style="width: 960px; height: 250px; display: inline-block"></div>
        <div id="sbnewslog_vto"  style="width: 480px; height: 250px; display: inline-block;"></div>
        <div id="sbnewslog_vpd"  style="width: 960px; height: 320px; display: block;"></div>
        <h1>Launcher Versions</h1>
        <div id="sbnewslog_lvtbl"  style="width: 250px; height: 250px; display: inline-block"></div>
        <div id="sbnewslog_lv" style="width: 960px; height: 250px; display: inline-block"></div>
        <div id="sbnewslog_lvto"  style="width: 480px; height: 250px; display: inline-block;"></div>
    </body>
</body>