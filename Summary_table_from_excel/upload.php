<?php
include_once 'file_to_db.php';

$target_dir = "uploads/";
$fileName = basename($_FILES["fileToUpload"]["name"]);
$target_file = $target_dir.$fileName;
$uploadOk = 1;
$imageFileType = pathinfo($target_file,PATHINFO_EXTENSION);

// Check if file already exists

// if (file_exists($target_file)) {
//     echo "Sorry, file already exists.";
//     $uploadOk = 0;
// }

// Check file size

// if ($_FILES["fileToUpload"]["size"] > 500000) {
//     echo "Sorry, your file is too large.";
//     $uploadOk = 0;
// }

// Allow certain file formats
if($imageFileType != "zip") {
    echo "Sorry, only ZIP files are allowed.";
    $uploadOk = 0;
}

// Check if $uploadOk is set to 0 by an error
if ($uploadOk == 0) {
    echo "Sorry, your file was not uploaded.";
// if everything is ok, try to upload file
} else {
    if (move_uploaded_file($_FILES["fileToUpload"]["tmp_name"], $target_file)) {
        // echo "The file ". basename( $_FILES["fileToUpload"]["name"]). " has been uploaded.";
        zip_extractor($fileName);
    } else {
        echo "Sorry, there was an error uploading your file.";
    }
}

function zip_extractor($fileName){
    // echo "</br>Extractor Called with filename : ".$fileName;
    $path = pathinfo(realpath("uploads/".$fileName), PATHINFO_DIRNAME);
    // echo "</br> File Path: ".$path."</br>";
    $zip = new ZipArchive;
    $res = $zip->open("uploads/".$fileName);
    if ($res === TRUE) {
      $zip->extractTo('extracted/');
      $zip->close();
      // echo 'File Extracted';
      export_filename();
    } else {
      echo 'Failed to extract file!';
    }
}

function export_filename(){
    $file_list = array("pole","sapling", "seedling", "tree");
    foreach ($file_list as $file) {
        file_to_db($file);
        run_summary_query($file);

    }
}

function run_summary_query($file){
    $function = $file.'s';
    $function();
}
?>