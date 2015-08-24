<?php
$server = 'localhost';
$login='root';
$pass='';
$db='netpoint_online';

$conn=mysql_connect($server,$login,$pass) or die ('failed connect to database');
//mysql_select_db($db,$conn);
//mysql_query("Set names 'utf8'");
?>