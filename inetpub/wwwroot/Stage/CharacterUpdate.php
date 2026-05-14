<?php
$uniqueTime = strtotime('now');
//file_put_contents('charlog.txt', '============================'.$uniqueTime.'============================'.PHP_EOL, FILE_APPEND | LOCK_EX);
	
	$serverName = "SD-140113\LUNIA_CLONE";   
	$uid = "sa";     
	$pwd = 'c5=F&35zSx5@#6fc';    
	$databaseName = "web";   
		   
	$connectionInfo = array( "UID"=>$uid, "PWD"=>$pwd,"Database"=>$databaseName);   

	$connsql = sqlsrv_connect( $serverName, $connectionInfo);    
	if( $connsql === false ) {die( print_r( sqlsrv_errors(), true));}
	
	$theData = urldecode($_POST['theData']);
	$array = preg_split("/\r\n|\n|\r/", $theData);
	
	foreach($array as $query){
			$runQuery = sqlsrv_query($connsql, 'USE [v3_character]; declare @nError int, @strSP varchar(50); '.$query);
			
			//file_put_contents('charlog.txt', $uniqueTime.$query.PHP_EOL, FILE_APPEND | LOCK_EX);
			if($runQuery === false){
				if( ($errors = sqlsrv_errors() ) != null) {
					$theError = '';
					foreach( $errors as $error ) {
						$theError.="	SQLSTATE: ".$error[ 'SQLSTATE']."
	code: ".$error[ 'code']."
	message: ".$error[ 'message'];
					}
					//file_put_contents('charlog.txt', $uniqueTime.$theError.PHP_EOL, FILE_APPEND | LOCK_EX);
				}
			}
	}
	
	echo '0';
?>