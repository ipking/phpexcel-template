<?php



/**
 **********************
 * 根据模板 导出excel文件
 **********************
 */
class ExportExcel
{
	
	private $ExportExcelParameter;
	public function addSheet(ExportExcelParameter $ExportExcelParameter){
		$this->ExportExcelParameter[] = $ExportExcelParameter;
	}
	
	/**
	 * 创建excel文件
	 * @param string $excel_template 模板路径
	 * @param string $filename 输出文件名称
	 * @param bool $is_download 是否web下载
	 * @return bool
	 * @throws \PHPExcel_Reader_Exception
	 * @throws \PHPExcel_Writer_Exception
	 * @throws \PHPExcel_Exception
	 */
	public function createExcel($excel_template,$filename,$is_download){
		$phpExcel = \PHPExcel_IOFactory::load($excel_template);
		
		foreach($this->ExportExcelParameter?:[] as $ExportExcelParameter){
			/**
			 * @var ExportExcelParameter $ExportExcelParameter
			 */
			if($ExportExcelParameter->sheet_main_data){
				$phpExcel = $this->writeSheetData($ExportExcelParameter->sheet_main_data, $phpExcel,$ExportExcelParameter->sheet_name,$ExportExcelParameter->sheet_main_data_ceil);
			}
			
			if($ExportExcelParameter->sheet_data_list){
				$phpExcel = $this->writeSheetList($ExportExcelParameter->sheet_data_list, $phpExcel,$ExportExcelParameter->sheet_name);
			}
		}
		
		
		$objWriter = \PHPExcel_IOFactory::createWriter($phpExcel, 'Excel2007');
		if($is_download){
			header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
			header('Content-Disposition: attachment;filename="'.$filename);
			header('Cache-Control: max-age=0');
			$objWriter->save( 'php://output');
		}else{
			$objWriter->save( $filename);
		}
		
		return true;
	}
	
	
	/**
	 * 写入 数据
	 * @param array $data 数据信息
	 * @param PHPExcel $phpExcel
	 * @param string $sheet_name 表名
	 * @param array $ceil 模板数据坐在的最大单元格
	 * @return PHPExcel
	 * @throws \PHPExcel_Exception
	 */
	private function writeSheetData($data, $phpExcel,$sheet_name,$ceil){
		
		//找到表格
		$currentSheet = $phpExcel->getSheetByName($sheet_name);
		
		if(!$currentSheet){
			throw new \Exception('找不到表名:'.$sheet_name);
		}
		//设置列数
		$columnNum = $ceil[0];
		//设置行数
		$rowNum = $ceil[1];
		//设置在A-Z 1-100的区域内 找到data 数据所处的坐标
		//开始遍历列
		for($colIndex='A'; $colIndex<=$columnNum; $colIndex++){
			if(!$data){
				break;
			}
			//遍历行
			for($rowIndex = 1; $rowIndex<=$rowNum; $rowIndex++){
				if(!$data){
					break;
				}
				$cell = $currentSheet->getCell($colIndex.$rowIndex)->getValue();
				if($cell instanceof PHPExcel_RichText){
					$cell = $cell->__toString();
				}
				if(isset($data[trim($cell,'$')])){
					//如果数据中有这个key 说明这个就是数据的坐标位置
					$currentSheet->setCellValue($colIndex.$rowIndex,$data[trim($cell,'$')]);
					
					unset($data[trim($cell,'$')]);
				}
			}
		}
		
		return $phpExcel;
	}
	
	/**
	 * 写入 列表数据
	 * @param array $resume_list 数据列表
	 * @param PHPExcel $phpExcel
	 * @param string $sheet_name 表名
	 * @return PHPExcel
	 * @throws \PHPExcel_Exception
	 */
	private function writeSheetList($resume_list, $phpExcel,$sheet_name){
		$resume = $resume_list[0]; //获取数据列表中第一条数据
		
		//找到表格
		$currentSheet = $phpExcel->getSheetByName($sheet_name);
		
		if(!$currentSheet){
			throw new \Exception('找不到表名:'.$sheet_name);
		}
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
		$i = 0;
		//开始遍历列
		while(true){
			//遍历行
			$rowIndex = $coordinate[1];
			$cell = $currentSheet->getCell($colIndex.$rowIndex)->getValue();
			if($cell instanceof PHPExcel_RichText){
				$cell = $cell->__toString();
			}
			//退出while 条件  不允许数据行存在空的单元格
			if(!$cell or substr(trim($cell),0,1) != '$'){
				$i++;
				
				//允许3个连续的列单元格 没有变量 或者 空值
				if($i>3){
					break;
				}else{
					//换到下一列开始
					$colIndex++;
					continue;
				}
			}
			//重置计数器
			$i = 0;
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