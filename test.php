<?php


include_once 'ExportExcel.php';
include_once 'lib/PHPExcel/PHPExcel.php';

$json = file_get_contents('data_list.json');
$arr = json_decode($json,1);
$e = new ExportExcel();
try{
	$e->createExcel($arr,'template.xlsx','demo.xlsx');
}catch(Exception $e){
	echo $e->getMessage();
}
