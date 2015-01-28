#!/usr/bin/php -q
<?php
//Include class
/* Created By Airkiss
*  Report PDF Information
*/
error_reporting(E_ALL);
ini_set("display_errors",true);
ini_set("html_errors",false);
date_default_timezone_set("Asia/Taipei");
require_once("../PHPExcel/Classes/PHPExcel.php");

function GetRawDataFromDB($DB,&$objPHPExcel)
{
	$dbh = new PDO($DB['DSN'],$DB['DB_USER'], $DB['DB_PWD'],
		array( PDO::ATTR_PERSISTENT => false));
	$dbh->setAttribute(PDO::ATTR_ERRMODE,PDO::ERRMODE_EXCEPTION);
	try {
			$p = $dbh->prepare("select id,bricklink,name,item_type,sub_name from item_info 
				where item_type in ('Sets','Gears') and packagingType is null");
			$p->execute();
			$resData = $p->fetchAll(PDO::FETCH_OBJ);
			var_dump($resData);
	} catch (PDOException $e) {
		print "Error: " . $e->getMessage() . "<br/>";
		return;
	}
	
	$rows = 2;
	foreach($resData as $item)
	{
		$objPHPExcel->getActiveSheet()->setCellValueExplicit("A$rows",$item->id,PHPExcel_Cell_DataType::TYPE_STRING);
		$objPHPExcel->getActiveSheet()->setCellValueExplicit("B$rows",$item->bricklink,PHPExcel_Cell_DataType::TYPE_STRING);
		$objPHPExcel->getActiveSheet()->setCellValueExplicit("C$rows",$item->name,PHPExcel_Cell_DataType::TYPE_STRING);
		$objPHPExcel->getActiveSheet()->setCellValueExplicit("D$rows",$item->item_type,PHPExcel_Cell_DataType::TYPE_STRING);
		$objPHPExcel->getActiveSheet()->setCellValueExplicit("E$rows",$item->sub_name,PHPExcel_Cell_DataType::TYPE_STRING);
		$rows++;
	}
}

function GenerateExcel($DB,$filename,$useTemplate=false)
{
	try {
		// Load Files
		if($useTemplate)
		{
			$objPHPExcel = PHPExcel_IOFactory::load("./Template3.xlsx");
		}
		else
			$objPHPExcel = PHPExcel_IOFactory::load($filename);

		$objPHPExcel->setActiveSheetIndex(0);
		
		GetRawDataFromDB($DB,$objPHPExcel);
		// Save File	
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel2007');
		$objWriter->save($filename);
	}catch (Exception $e) {
		echo "PHPExcel Error : ".$e->getMessage()."<BR>";
		return;
	}
	return ;
}


$ini_array = parse_ini_file("db.ini",true);
$DB = $ini_array['DB'];
$filename = "bricksetNoMapping.xlsx";
GenerateExcel($DB,$filename,true);
?>
