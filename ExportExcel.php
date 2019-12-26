<?php
/**
 **********************
 * 根据模板 导出excel文件
 **********************
*/
class ExportExcel
{
	/**
	 * 创建excel文件
	 * @param array $resume_list 数据信息
	 * @param string $excel_template 模板路径
	 * @param string $file_path 输出文件路径
	 * @return bool
	 * @throws \PHPExcel_Reader_Exception
	 * @throws \PHPExcel_Writer_Exception
	 * @throws \PHPExcel_Exception
	 */
    public function createExcel($resume_list,$excel_template,$file_path){
        $phpExcel = \PHPExcel_IOFactory::load($excel_template);
        $phpExcel = $this->writeSheet($resume_list, $phpExcel);
        $objWriter = \PHPExcel_IOFactory::createWriter($phpExcel, 'Excel2007');
        $objWriter->save($file_path);
        return true;
    }
	
	
	/**
	 * 写入表数据
	 * @param array $resume_list 数据信息
	 * @param PHPExcel $phpExcel
	 * @param int $index 第几个表 从0开始
	 * @return PHPExcel
	 * @throws \PHPExcel_Exception
	 */
    private function writeSheet($resume_list, $phpExcel,$index=0){
	    $resume = $resume_list[0]; //获取数据列表中第一条数据
        
        //找到表格
        $currentSheet = $phpExcel->getSheet($index);
        //设置默认坐标位置 为左上角
        $coordinate = ['A',1];
	    //设置列数
	    $columnNum = 'Z';
	    //设置行数
	    $rowNum = 20;
	    //设置在A-Z 1-20的区域内 找到list 数据所处的坐标
	    //开始遍历列
	    for($colIndex='A'; $colIndex<=$columnNum; $colIndex++){
		    //遍历行
		    for($rowIndex = 1; $rowIndex<=$rowNum; $rowIndex++){
			    $cell = $currentSheet->getCell($colIndex.$rowIndex)->getValue();
			    if($cell instanceof PHPExcel_RichText){
				    $cell = $cell->__toString();
			    }
			    if(isset($resume[trim($cell,'$')])){
			    	//如果数据中有这个key 说明这个就是数据的坐标位置
				    $coordinate = [$colIndex,$rowIndex];
				    break 2;
			    }
		    }
	    }
	
	    $colIndex=$coordinate[0];
	    //写入数据
	    //如果有n条数据 就插入n-1条空行 吧数据行往下移动 保证数据内容下面的内容不会被数据覆盖
	    if(count($resume_list)-1 > 0){
		    $currentSheet->insertNewRowBefore($coordinate[1]+1,count($resume_list)-1);
	    }
        //开始遍历列
	    while(true){
		    //遍历行
		    $rowIndex = $coordinate[1];
		    $cell = $currentSheet->getCell($colIndex.$rowIndex)->getValue();
		    if($cell instanceof PHPExcel_RichText){
			    $cell = $cell->__toString();
		    }
		    //退出while 条件  不允许数据行存在空的单元格
		    if(!$cell){
		    	break;
		    }
		    foreach ($resume_list as $value) {
		    	//遍历数据列表,对应的值填充到该列的下一行
			    $currentSheet->setCellValue($colIndex.$rowIndex,$value[trim($cell,'$')]);
			    $rowIndex++;
		    }
		
		    //换到下一列开始
		    $colIndex++;
	    }
       
        return $phpExcel;
    }
}