<?php 
include_once "db.php";
/**Convert PHP data type to SQL equivalent data type*/
function phpToSqlEquivalent($dataType=""){
	// echo "</br>".$dataType."</br>";
    switch ($dataType) {
        case 'string':
            return "varchar(100)";
            break;

        case 'double':
            return "double";
            break;
        
        default:
	        if($dataType=="" || $dataType=="NULL"){
	        	$dataType = "varchar(100)";
	        }
	        return $dataType;
	        break;
    }
}

function createPivotTable($columnNames){
	$columnNameList = array();
	$rowValuesList = array();
	global $conn;
	
	$createTableQuery = "CREATE TABLE IF NOT EXISTS summary_table(name varchar(100), total_count double);";
	if ($conn->query($createTableQuery) === TRUE) {
        // echo "New table created successfully";
    } else {
        // echo "Error: " . $sql . "<br>" . $conn->error;
    }


	$sql = "SELECT DISTINCT name FROM example";
	$result = $conn->query($sql);

	if ($result->num_rows > 0) {
	    // output data of each row
	    while($row = $result->fetch_assoc()) {
	        $name = $row["name"];
	        $sql1 = "SELECT name, SUM(Count) as total_count FROM example WHERE name = '".$name."';";
	        $result1 = $conn->query($sql1);
	        if ($result1->num_rows > 0) {
		    	// output data of each row
		    	while($row1 = $result1->fetch_assoc()) {
					//echo $sql1."</br>";
					//print_r($row1);
					foreach($row1 as $columnName=>$value){
						array_push($columnNameList, $columnName);
						array_push($rowValuesList, $value);
					}
					$sql1 = "INSERT INTO summary_table(".implode(",",$columnNameList).") VALUES ('".implode("','", $rowValuesList)."')";
					//echo "Second query is : ".$sql1."</br>";
		            if ($conn->query($sql1) === TRUE) {
		                // echo "New record created successfully";
		            } else {
		                // echo "Error: " . $sql . "<br>" . $conn->error;
		            }
		            $columnNameList = array();
		            $rowValuesList = array();

		    	}
		    }
		}
	} else {
	    // echo "0 results";
	}
	exportToExcel();
}

function exportToExcel(){
global $conn;	
$objPHPExcel = new PHPExcel();

 $sql2 = "SELECT * FROM summary_table";

$result2 = $conn->query($sql2);

 $objPHPExcel = new PHPExcel();

 $rowCount_exportToExcel = 1;

 while($row2 = $result2->fetch_assoc()){
	$objPHPExcel->getActiveSheet()->SetCellValue('A'.$rowCount_exportToExcel, $row2['name']);
    $objPHPExcel->getActiveSheet()->SetCellValue('B'.$rowCount_exportToExcel, $row2['total_count']);
	$rowCount_exportToExcel++;
	// pr($objPHPExcel);
 }

 header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
 header('Content-Disposition: attachment;filename="summary_data.xlsx"');
 header('Cache-Control: max-age=0');
ob_end_clean();
flush();
 $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
 $objWriter->save('php://output');
 $objWriter->save('summary_data.xlsx');
 // clearTabless();

 $sql3 = "DROP TABLE IF EXISTS example, summary_table;";
 $conn->query($sql3);


}

function db_to_excel_pole($usr_forest_block){
	// global $conn;	
	$objPHPExcel = new PHPExcel();
	$rowCount_exportToExcel = 2;
	$styleArray = array(
	    'font'  => array(
	        'bold'  => true,
	        'color' => array('rgb' => '000000'),
	        'size'  => 12,
	        'name'  => 'calibri'
	        ));
	$cellStyleArray = array(
	    'fill'  => array(
	    	'type' => PHPExcel_Style_Fill::FILL_SOLID,
	        'color' => array('rgb' => 'ccddcc')
	        ));

	$objPHPExcel->getActiveSheet()->getRowDimension('1')->setRowHeight(35);


	$objPHPExcel->getActiveSheet()->SetCellValue('A1', "Surveyer");
	$objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);


	$objPHPExcel->getActiveSheet()->SetCellValue('B1', "CF Code");
	$objPHPExcel->getActiveSheet()->getStyle('B1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('B1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);


	$objPHPExcel->getActiveSheet()->SetCellValue('C1', "Forest Block");
	$objPHPExcel->getActiveSheet()->getStyle('C1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('C1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);

	$objPHPExcel->getActiveSheet()->SetCellValue('D1', "Species");
	$objPHPExcel->getActiveSheet()->getStyle('D1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('D1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);

	$objPHPExcel->getActiveSheet()->SetCellValue('E1', "#pole/ha");
	$objPHPExcel->getActiveSheet()->getStyle('E1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('E1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);

	$objPHPExcel->getActiveSheet()->SetCellValue('F1', "Pole Volume(m3/ha)");
	$objPHPExcel->getActiveSheet()->getStyle('F1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('F1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);

	$objPHPExcel->getActiveSheet()->SetCellValue('G1', "Timber(m3/ha)");
	$objPHPExcel->getActiveSheet()->getStyle('G1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('G1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);

	$objPHPExcel->getActiveSheet()->SetCellValue('H1', "Firewood(m3/ha)");
	$objPHPExcel->getActiveSheet()->getStyle('H1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('H1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);

	$objPHPExcel->getActiveSheet()->SetCellValue('I1', "CO2 (MT/ha)");
	$objPHPExcel->getActiveSheet()->getStyle('I1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('I1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);

	foreach ($usr_forest_block as $name => $name_values) {
		$objPHPExcel->getActiveSheet()->SetCellValue('A'.$rowCount_exportToExcel, $name);

		foreach ($name_values as $forest => $forest_values) {
			$objPHPExcel->getActiveSheet()->SetCellValue('B'.$rowCount_exportToExcel, $forest);

			foreach ($forest_values as $block => $block_values) {
				$objPHPExcel->getActiveSheet()->SetCellValue('c'.$rowCount_exportToExcel, $block);
				$arr_tree_total = array('sum_counth'=>0,'sum_volumeh'=>0,'sum_timberh'=>0,'sum_firewoodh'=>0,'sum_co2h'=>0);
				foreach ($block_values as $tree => $tree_values) {
					$objPHPExcel->getActiveSheet()->SetCellValue('D'.$rowCount_exportToExcel, $tree);
					$colCount_exportToExcel = array('E','F','G','H','I');
					$arr_count = 0;
					foreach ($tree_values as $param => $value) {
						$objPHPExcel->getActiveSheet()->SetCellValue($colCount_exportToExcel[$arr_count].$rowCount_exportToExcel, $value);
						$arr_count++;
						$arr_tree_total[$param] =$arr_tree_total[$param] + $value;
					}
					$rowCount_exportToExcel++;
					// $colCount_exportToExcel++;
				}
				$objPHPExcel->getActiveSheet()->SetCellValue('C'.$rowCount_exportToExcel, $block.' Total');
				$objPHPExcel->getActiveSheet()->getStyle('C'.$rowCount_exportToExcel)->applyFromArray($styleArray);
				$objPHPExcel->getActiveSheet()->getRowDimension($rowCount_exportToExcel)->setRowHeight(25);


				$colCount_exportToExcel_total = array('E','F','G','H','I');
				$arr_count_total = 0;
				foreach ($arr_tree_total as $param => $value) {
					$objPHPExcel->getActiveSheet()->SetCellValue($colCount_exportToExcel_total[$arr_count_total].$rowCount_exportToExcel, $value);
					$objPHPExcel->getActiveSheet()->getStyle($colCount_exportToExcel_total[$arr_count_total].$rowCount_exportToExcel)->applyFromArray($styleArray);
					$arr_count_total++;
				} $rowCount_exportToExcel++;
			}
		}
		// pr($objPHPExcel);
	 }

	 // header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	 // header('Content-Disposition: attachment;filename="pole_summary_data.xlsx"');
	 // header('Cache-Control: max-age=0');
	ob_end_clean();
	flush();
	 $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	 $objWriter->save('php://output');
	 $objWriter->save('summary/pole_summary_data.xlsx');
	 // clearTabless();

}

function db_to_excel_sapling($usr_forest_block){
	// global $conn;	
	$objPHPExcel = new PHPExcel();
	$rowCount_exportToExcel = 2;
	$styleArray = array(
	    'font'  => array(
	        'bold'  => true,
	        'color' => array('rgb' => '000000'),
	        'size'  => 12,
	        'name'  => 'calibri'
	        ));
	$cellStyleArray = array(
	    'fill'  => array(
	    	'type' => PHPExcel_Style_Fill::FILL_SOLID,
	        'color' => array('rgb' => 'ccddcc')
	        ));

	$objPHPExcel->getActiveSheet()->getRowDimension('1')->setRowHeight(35);


	$objPHPExcel->getActiveSheet()->SetCellValue('A1', "Surveyer");
	$objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);


	$objPHPExcel->getActiveSheet()->SetCellValue('B1', "CF Code");
	$objPHPExcel->getActiveSheet()->getStyle('B1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('B1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);


	$objPHPExcel->getActiveSheet()->SetCellValue('C1', "Forest Block");
	$objPHPExcel->getActiveSheet()->getStyle('C1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('C1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);

	$objPHPExcel->getActiveSheet()->SetCellValue('D1', "Species");
	$objPHPExcel->getActiveSheet()->getStyle('D1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('D1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);

	$objPHPExcel->getActiveSheet()->SetCellValue('E1', "Number/hectare");
	$objPHPExcel->getActiveSheet()->getStyle('E1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('E1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);

	foreach ($usr_forest_block as $name => $name_values) {
		$objPHPExcel->getActiveSheet()->SetCellValue('A'.$rowCount_exportToExcel, $name);

		foreach ($name_values as $forest => $forest_values) {
			$objPHPExcel->getActiveSheet()->SetCellValue('B'.$rowCount_exportToExcel, $forest);

			foreach ($forest_values as $block => $block_values) {
				$objPHPExcel->getActiveSheet()->SetCellValue('c'.$rowCount_exportToExcel, $block);
				$arr_tree_total = array('sum_counth'=>0);
				foreach ($block_values as $tree => $tree_values) {
					$objPHPExcel->getActiveSheet()->SetCellValue('D'.$rowCount_exportToExcel, $tree);
					$colCount_exportToExcel = array('E');
					$arr_count = 0;
					foreach ($tree_values as $param => $value) {
						$objPHPExcel->getActiveSheet()->SetCellValue($colCount_exportToExcel[$arr_count].$rowCount_exportToExcel, $value);
						$arr_count++;
						$arr_tree_total[$param] =$arr_tree_total[$param] + $value;
					}
					$rowCount_exportToExcel++;
					// $colCount_exportToExcel++;
				}
				$objPHPExcel->getActiveSheet()->SetCellValue('C'.$rowCount_exportToExcel, $block.' Total');
				$objPHPExcel->getActiveSheet()->getStyle('C'.$rowCount_exportToExcel)->applyFromArray($styleArray);
				$objPHPExcel->getActiveSheet()->getRowDimension($rowCount_exportToExcel)->setRowHeight(25);


				$colCount_exportToExcel_total = array('E');
				$arr_count_total = 0;
				foreach ($arr_tree_total as $param => $value) {
					$objPHPExcel->getActiveSheet()->SetCellValue($colCount_exportToExcel_total[$arr_count_total].$rowCount_exportToExcel, $value);
					$objPHPExcel->getActiveSheet()->getStyle($colCount_exportToExcel_total[$arr_count_total].$rowCount_exportToExcel)->applyFromArray($styleArray);
					$arr_count_total++;
				} $rowCount_exportToExcel++;
			}
		}
		// pr($objPHPExcel);
	 }

	 // header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	 // header('Content-Disposition: attachment;filename="pole_summary_data.xlsx"');
	 // header('Cache-Control: max-age=0');
	ob_end_clean();
	flush();
	 $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	 $objWriter->save('php://output');
	 $objWriter->save('summary/sapling_summary_data.xlsx');
	 // clearTabless();

}

function db_to_excel_seedling($usr_forest_block){
	// global $conn;	
	$objPHPExcel = new PHPExcel();
	$rowCount_exportToExcel = 2;
	$styleArray = array(
	    'font'  => array(
	        'bold'  => true,
	        'color' => array('rgb' => '000000'),
	        'size'  => 12,
	        'name'  => 'calibri'
	        ));
	$cellStyleArray = array(
	    'fill'  => array(
	    	'type' => PHPExcel_Style_Fill::FILL_SOLID,
	        'color' => array('rgb' => 'ccddcc')
	        ));

	$objPHPExcel->getActiveSheet()->getRowDimension('1')->setRowHeight(35);


	$objPHPExcel->getActiveSheet()->SetCellValue('A1', "Surveyer");
	$objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);


	$objPHPExcel->getActiveSheet()->SetCellValue('B1', "CF Code");
	$objPHPExcel->getActiveSheet()->getStyle('B1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('B1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);


	$objPHPExcel->getActiveSheet()->SetCellValue('C1', "Forest Block");
	$objPHPExcel->getActiveSheet()->getStyle('C1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('C1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);

	$objPHPExcel->getActiveSheet()->SetCellValue('D1', "Species");
	$objPHPExcel->getActiveSheet()->getStyle('D1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('D1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);

	$objPHPExcel->getActiveSheet()->SetCellValue('E1', "Number/hectare");
	$objPHPExcel->getActiveSheet()->getStyle('E1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('E1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);

	foreach ($usr_forest_block as $name => $name_values) {
		$objPHPExcel->getActiveSheet()->SetCellValue('A'.$rowCount_exportToExcel, $name);

		foreach ($name_values as $forest => $forest_values) {
			$objPHPExcel->getActiveSheet()->SetCellValue('B'.$rowCount_exportToExcel, $forest);

			foreach ($forest_values as $block => $block_values) {
				$objPHPExcel->getActiveSheet()->SetCellValue('c'.$rowCount_exportToExcel, $block);
				$arr_tree_total = array('sum_counth'=>0);
				foreach ($block_values as $tree => $tree_values) {
					$objPHPExcel->getActiveSheet()->SetCellValue('D'.$rowCount_exportToExcel, $tree);
					$colCount_exportToExcel = array('E');
					$arr_count = 0;
					foreach ($tree_values as $param => $value) {
						$objPHPExcel->getActiveSheet()->SetCellValue($colCount_exportToExcel[$arr_count].$rowCount_exportToExcel, $value);
						$arr_count++;
						$arr_tree_total[$param] =$arr_tree_total[$param] + $value;
					}
					$rowCount_exportToExcel++;
					// $colCount_exportToExcel++;
				}
				$objPHPExcel->getActiveSheet()->SetCellValue('C'.$rowCount_exportToExcel, $block.' Total');
				$objPHPExcel->getActiveSheet()->getStyle('C'.$rowCount_exportToExcel)->applyFromArray($styleArray);
				$objPHPExcel->getActiveSheet()->getRowDimension($rowCount_exportToExcel)->setRowHeight(25);


				$colCount_exportToExcel_total = array('E');
				$arr_count_total = 0;
				foreach ($arr_tree_total as $param => $value) {
					$objPHPExcel->getActiveSheet()->SetCellValue($colCount_exportToExcel_total[$arr_count_total].$rowCount_exportToExcel, $value);
					$objPHPExcel->getActiveSheet()->getStyle($colCount_exportToExcel_total[$arr_count_total].$rowCount_exportToExcel)->applyFromArray($styleArray);
					$arr_count_total++;
				} $rowCount_exportToExcel++;
			}
		}
		// pr($objPHPExcel);
	 }

	 // header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	 // header('Content-Disposition: attachment;filename="pole_summary_data.xlsx"');
	 // header('Cache-Control: max-age=0');
	ob_end_clean();
	flush();
	 $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	 $objWriter->save('php://output');
	 $objWriter->save('summary/seedling_summary_data.xlsx');
	 // clearTabless();

}

function db_to_excel_tree($usr_forest_block){
	// global $conn;	
	$objPHPExcel = new PHPExcel();
	$rowCount_exportToExcel = 2;
	$styleArray = array(
	    'font'  => array(
	        'bold'  => true,
	        'color' => array('rgb' => '000000'),
	        'size'  => 12,
	        'name'  => 'calibri'
	        ));
	$cellStyleArray = array(
	    'fill'  => array(
	    	'type' => PHPExcel_Style_Fill::FILL_SOLID,
	        'color' => array('rgb' => 'ccddcc')
	        ));

	$objPHPExcel->getActiveSheet()->getRowDimension('1')->setRowHeight(35);


	$objPHPExcel->getActiveSheet()->SetCellValue('A1', "Surveyer");
	$objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);


	$objPHPExcel->getActiveSheet()->SetCellValue('B1', "CF Code");
	$objPHPExcel->getActiveSheet()->getStyle('B1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('B1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);


	$objPHPExcel->getActiveSheet()->SetCellValue('C1', "Forest Block");
	$objPHPExcel->getActiveSheet()->getStyle('C1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('C1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);

	$objPHPExcel->getActiveSheet()->SetCellValue('D1', "Species");
	$objPHPExcel->getActiveSheet()->getStyle('D1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('D1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);

	$objPHPExcel->getActiveSheet()->SetCellValue('E1', "#Tree/ha");
	$objPHPExcel->getActiveSheet()->getStyle('E1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('E1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);

	$objPHPExcel->getActiveSheet()->SetCellValue('F1', "Tree Volume(m3/ha)");
	$objPHPExcel->getActiveSheet()->getStyle('F1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('F1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);

	$objPHPExcel->getActiveSheet()->SetCellValue('G1', "Tree Timber(m3/ha)");
	$objPHPExcel->getActiveSheet()->getStyle('G1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('G1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);

	$objPHPExcel->getActiveSheet()->SetCellValue('H1', "Tree Firewood(m3/ha)");
	$objPHPExcel->getActiveSheet()->getStyle('H1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('H1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);

	$objPHPExcel->getActiveSheet()->SetCellValue('I1', "Tree CO2 (MT/ha)");
	$objPHPExcel->getActiveSheet()->getStyle('I1')->applyFromArray($cellStyleArray);
	$objPHPExcel->getActiveSheet()->getStyle('I1')->applyFromArray($styleArray);
	$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);

	foreach ($usr_forest_block as $name => $name_values) {
		$objPHPExcel->getActiveSheet()->SetCellValue('A'.$rowCount_exportToExcel, $name);

		foreach ($name_values as $forest => $forest_values) {
			$objPHPExcel->getActiveSheet()->SetCellValue('B'.$rowCount_exportToExcel, $forest);

			foreach ($forest_values as $block => $block_values) {
				$objPHPExcel->getActiveSheet()->SetCellValue('c'.$rowCount_exportToExcel, $block);
				$arr_tree_total = array('sum_counth'=>0,'sum_volumeh'=>0,'sum_timberh'=>0,'sum_firewoodh'=>0,'sum_co2h'=>0);
				foreach ($block_values as $tree => $tree_values) {
					$objPHPExcel->getActiveSheet()->SetCellValue('D'.$rowCount_exportToExcel, $tree);
					$colCount_exportToExcel = array('E','F','G','H','I');
					$arr_count = 0;
					foreach ($tree_values as $param => $value) {
						$objPHPExcel->getActiveSheet()->SetCellValue($colCount_exportToExcel[$arr_count].$rowCount_exportToExcel, $value);
						$arr_count++;
						$arr_tree_total[$param] =$arr_tree_total[$param] + $value;
					}
					$rowCount_exportToExcel++;
					// $colCount_exportToExcel++;
				}
				$objPHPExcel->getActiveSheet()->SetCellValue('C'.$rowCount_exportToExcel, $block.' Total');
				$objPHPExcel->getActiveSheet()->getStyle('C'.$rowCount_exportToExcel)->applyFromArray($styleArray);
				$objPHPExcel->getActiveSheet()->getRowDimension($rowCount_exportToExcel)->setRowHeight(25);


				$colCount_exportToExcel_total = array('E','F','G','H','I');
				$arr_count_total = 0;
				foreach ($arr_tree_total as $param => $value) {
					$objPHPExcel->getActiveSheet()->SetCellValue($colCount_exportToExcel_total[$arr_count_total].$rowCount_exportToExcel, $value);
					$objPHPExcel->getActiveSheet()->getStyle($colCount_exportToExcel_total[$arr_count_total].$rowCount_exportToExcel)->applyFromArray($styleArray);
					$arr_count_total++;
				} $rowCount_exportToExcel++;
			}
		}
		// pr($objPHPExcel);
	 }

	 // header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	 // header('Content-Disposition: attachment;filename="pole_summary_data.xlsx"');
	 // header('Cache-Control: max-age=0');
	ob_end_clean();
	flush();
	 $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	 $objWriter->save('php://output');
	 $objWriter->save('summary/tree_summary_data.xlsx');
	 // clearTabless();
}

function compress_all(){
	$the_folder = 'summary';
	$zip = 'summarized_data_'.(rand()+rand()*rand(5,28)/rand());
	$zip_name = $zip.".zip";
	$zip_file_name = 'compressed/'.$zip_name;

	class FlxZipArchive extends ZipArchive {
	        /** Add a Dir with Files and Subdirs to the archive;;;;; @param string $location Real Location;;;;  @param string $name Name in Archive;;; @author Nicolas Heimann;;;; @access private  **/
	    public function addDir($location, $name) {
	        $this->addEmptyDir($name);
	         $this->addDirDo($location, $name);
	     } // EO addDir;

	        /**  Add Files & Dirs to archive;;;; @param string $location Real Location;  @param string $name Name in Archive;;;;;; @author Nicolas Heimann * @access private   **/
	    private function addDirDo($location, $name) {
	        $name .= '/';         $location .= '/';
	      // Read all Files in Dir
	        $dir = opendir ($location);
	        while ($file = readdir($dir))    {
	            if ($file == '.' || $file == '..') continue;
	          // Rekursiv, If dir: FlxZipArchive::addDir(), else ::File();
	            $do = (filetype( $location . $file) == 'dir') ? 'addDir' : 'addFile';
	            $this->$do($location . $file, $name . $file);
	        }
	    } 
	}

	$za = new FlxZipArchive;
	$res = $za->open($zip_file_name, ZipArchive::CREATE);
	if($res === TRUE)    {
	    $za->addDir($the_folder, basename($the_folder)); 
	    $za->close();
	    echo "<script>location.href = 'download.php?fname=".$zip."'</script>";
	}
	else  { echo 'Could not create a zip archive';}
	}
?>