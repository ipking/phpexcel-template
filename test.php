<?php


include_once 'ExportExcel.php';
include_once 'ExportExcelParameter.php';
include_once 'lib/PHPExcel/PHPExcel.php';

$json = file_get_contents('data_list.json');
$data_list = json_decode($json,1);

$data['_title'] = '这是标题';
$e = new ExportExcel();
try{
	$param = new ExportExcelParameter();
	$param->sheet_name = 'Sheet1';
	$param->sheet_data_list = array_values($data_list);
	$param->sheet_main_data = $data;
	$param->sheet_main_data_ceil = ['F',12];
	$e->addSheet($param);
	$e->createExcel('template.xlsx','demo.xlsx',false);
}catch(Exception $e){
	echo $e->getMessage();
}

$k = 1;