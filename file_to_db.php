<?php
include_once 'db.php';
include_once "utils.php";
include_once "pole.php";
include_once "sapling.php";
include_once "seedling.php";
include_once "tree.php";

//  Include PHPExcel_IOFactory
include_once 'PHPExcel-1.8/Classes/PHPExcel/IOFactory.php';

function file_to_db($inputFileName){
    global $conn;

    $absoluteInputFileName = "extracted/".$inputFileName.".csv";

    //  Read your Excel workbook
    try {
        $inputFileType = PHPExcel_IOFactory::identify($absoluteInputFileName);
        $objReader = PHPExcel_IOFactory::createReader($inputFileType);
        $objPHPExcel = $objReader->load($absoluteInputFileName);
    } catch(Exception $e) {
        die('Error loading file "'.pathinfo($absoluteInputFileName,PATHINFO_BASENAME).'": '.$e->getMessage());
    }

    //  Get worksheet dimensions
    $sheet = $objPHPExcel->getSheet(0); 
    $highestRow = $sheet->getHighestRow(); 
    $highestColumn = $sheet->getHighestColumn();
    $rowDataTypes = $sheet->rangeToArray('A2:' . $highestColumn .'2',NULL,TRUE,FALSE); //is used to determine data type of column later
    //  Loop through each row of the worksheet in turn
    $rowCount = 0; //used to determine count of row(row number)
    $columnNames = array();
    $query = "CREATE TABLE IF NOT EXISTS ".$inputFileName."(id INT NOT NULL AUTO_INCREMENT PRIMARY KEY";
    for ($row = 1; $row <= $highestRow; $row++){ 
        //  Read a row of data into an array
        $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row,
                                        NULL,
                                        TRUE,
                                        FALSE);
        $size = sizeof($rowData[0]);
        for($x = 0; $x < $size; $x++) {
            if($rowCount == 0){
                $query .= ",".$rowData[0][$x]." ".phpToSqlEquivalent(gettype($rowDataTypes[0][$x]));
                array_push($columnNames,$rowData[0][$x]);

            } 
            else{
                // echo "error</br>";
            }

            $flag = ($size-1);
            if ($x== $flag) {
                    $rowCount++;
            }
        }
        if ($rowCount == 1) {
                $query .= ");"; 
                // echo "</br>".$query."</br>";
                if ($conn->query($query) === TRUE) {
                    // echo "New table created successfully";
                } else {
                    // echo "Error: table creation error: <br>" . $conn->error;
                }
        } else{
            $sql = "INSERT INTO ".$inputFileName."(".implode(",",$columnNames).") VALUES ('".implode("','", $rowData[0])."')";
                if ($conn->query($sql) === TRUE) {
                    // echo "New record created successfully</br>";
                } else {
                    // echo "Error: " . $sql . "<br>" . $conn->error."<br>";
                }
        }
    }
    // print_r($columnNames);
    // $conn->close();
}

?>