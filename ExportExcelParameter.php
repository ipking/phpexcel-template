<?php



/**
 **********************
 * 导出表格参数
 **********************
*/


/**
 * @property string $sheet_name 需要操作的表名
 * @property array $sheet_data_list  需要填充的列表数据
 * @property array $sheet_main_data  需要填充的单独数据
 * @property array $sheet_main_data_ceil 需要填充的单独数据的 最大单元格的位置 ['A','100']
 */
class ExportExcelParameter
{
	public $sheet_name;
	public $sheet_data_list;
	public $sheet_main_data;
	public $sheet_main_data_ceil;
}