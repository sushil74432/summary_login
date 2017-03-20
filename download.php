<!DOCTYPE html>
<html>
<body>
	<h1>Downloading...</h1>
	<?php 
		$fileName = $_GET['fname'];
		echo "<script type='text/javascript'> location.href = 'compressed/".$fileName.".zip'</script>";
	 ?>
</body>
</html>