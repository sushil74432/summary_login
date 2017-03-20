<?php
include_once 'db.php';
include_once "utils.php";
ini_set('xdebug.var_display_max_depth', 5);
ini_set('xdebug.var_display_max_children', 256);
ini_set('xdebug.var_display_max_data', 1024);

//  Include PHPExcel_IOFactory
include_once 'PHPExcel-1.8/Classes/PHPExcel/IOFactory.php';

function poles(){
    global $conn;
    $user_qry = "SELECT DISTINCT(forest_user_name) FROM pole";
    $user_list = $conn->query($user_qry);
    $user_list_array = array();
    while($row = $user_list->fetch_assoc()) {
        $name = $row['forest_user_name'];
        array_push($user_list_array, $name);
    }
    // print_r($user_list_array);

    $forest_list_array = array();
    foreach ($user_list_array as $users) {
        $cf_qry = "SELECT DISTINCT(forest_regime_cfcode) FROM pole WHERE forest_user_name = '".$users."'";
        // echo $cf_qry."\n";
        $forest_list = $conn->query($cf_qry);
        $forest_name = array();
        while($row = $forest_list->fetch_assoc()) {
            $forest_name[] = $row['forest_regime_cfcode'];
            // array_push($forest_list_array, $forest_name);
            $forest_list_array[$users] = $forest_name;
        }
    }
    // var_dump($forest_list_array);


    $usr_forest_block = array();
    foreach ($forest_list_array as $username => $forests) {
        $block_vs_forest = array();
        foreach ($forests as $forest) {
            $tree_vs_block = array();
            $block_list_array = array();
            $block_qry = "SELECT DISTINCT(forest_block_block_num) FROM pole WHERE forest_regime_cfcode = '".$forest."' AND forest_user_name = '".$username."'";
            $block_list = $conn->query($block_qry);
            $tree_list_array = array();
            while($row = $block_list->fetch_assoc()) {
                $block_name = $row['forest_block_block_num'];
                $block_list_array[] = $block_name;
                $tree_query = "SELECT DISTINCT(species_vernacular_name) FROM pole WHERE forest_block_block_num='".$block_name."' AND forest_regime_cfcode ='".$forest."' AND forest_user_name = '".$username."'";
                $tree_list = $conn->query($tree_query);
                $tree_summary_vs_tree = array();
                while($tree_row = $tree_list->fetch_assoc()) {
                    $tree_name = $tree_row['species_vernacular_name'];
                    // $tree_list_array[] =$tree_name;

                    $summary_report_query = "SELECT SUM(counth) as sum_counth, SUM(volumeh) as sum_volumeh, SUM(timberh) as sum_timberh, SUM(firewoodh) as sum_firewood, SUM(co2h) as sum_co2h FROM pole WHERE species_vernacular_name = '".$tree_name."' AND forest_block_block_num='".$block_name."' AND forest_regime_cfcode ='".$forest."' AND forest_user_name = '".$username."'";
                    // echo $summary_report_query."</br>";
                    $general_summary = $conn->query($summary_report_query);
                    // var_dump($general_summary);
                    $tree_summary = array();
                    while($summary = $general_summary->fetch_assoc()) {
                        $tree_summary['sum_counth'] = $summary['sum_counth'];
                        $tree_summary['sum_volumeh'] = $summary['sum_volumeh'];
                        $tree_summary['sum_timberh'] = $summary['sum_timberh'];                    
                        $tree_summary['sum_firewoodh'] = $summary['sum_firewood'];
                        $tree_summary['sum_co2h'] = $summary['sum_co2h'];
                        // var_dump($tree_summary);
                        $tree_summary_vs_tree[$tree_name] = $tree_summary;
                    }
                    // var_dump($tree_summary_vs_tree);
                    $tree_list_array[$block_name] = $tree_summary_vs_tree;
                }
                // var_dump($tree_list_array);
                // var_dump($tree_vs_block);
            }
            // var_dump($tree_list_array);
            $tree_vs_block = $tree_list_array;

            // var_dump($tree_vs_block);

            $block_vs_forest[$forest]=$tree_vs_block;
        }
        // var_dump($block_vs_forest);

        $usr_forest_block[$username]= $block_vs_forest;
    }

    // echo "</br></br></br></br></br>";
    // var_dump($usr_forest_block);

    db_to_excel_pole($usr_forest_block);
    $drop_pole_table = "DROP TABLE IF EXISTS pole;";
    $conn->query($drop_pole_table);
}
?>