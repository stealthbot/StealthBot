<?php
    /** 
    *   Andy Trevino
    *   Project:     
    *   File:       
    */
    
    define("__READMEFILE", true);
    
    require("data.php");
    
    $query = "INSERT INTO commands(name, access, arguments, wildcardable, desc) VALUES('$name','$access','$args','$wc','$desc');";
    
    $query = "INSERT INTO examples(
    
?>