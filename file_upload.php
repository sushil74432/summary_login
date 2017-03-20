<?php require('includes/config.php'); 

//if not logged in redirect to login page
if(!$user->is_logged_in()){ header('Location: login.php'); } 

//define page title
$title = 'Members Page';

//include header template
require('layout/header.php'); 
?>
<div class="upload_form col-lg-12" id="upload_form">
	<form action="upload.php" method="post" enctype="multipart/form-data">
	    <div class="label_custom">
	    	<input type="readonly" class="label_readonly" id="label_readonly" value="Select field Inventory ZIP file to upload">
	    </div>
	    <input type="file" name="fileToUpload" id="fileToUpload" class="fileToUpload">
	    <input type="submit" value="Get Growing Stock Summary Report" name="submit" class="btn btn-success btn-large upload-btn">
	</form>
	<span class="info_text">
		<br>This is a machine. It helps our collected forest inventory data to analyize.<br>
		<br>The Data Must be taken from android mobile application and exported from it.<br>
		<br>Please visit for more info howto: <a href="https://sites.google.com/view/forest-inventory-calculater/home">Learn more using imis data collection and analysis</a></br>
		<a href="logout.php" class = "logout-btn"><button class="btn btn-success logout-btn">Logout</button></a>
	</span>
</div>

<script type="text/javascript">
	$("input:file").change(function (){
		var fileName = $('input[type=file]').val().split('\\').pop();
		$("#label_readonly").val(fileName);
		console.log(fileName);
	})
</script>

<?php 
//include header template
require('layout/footer.php'); 
?>
