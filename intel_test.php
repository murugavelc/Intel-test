/**
 * Created by Murugavel C of bahwancybertek.com
 * Intel-Offline Question
 *
 */

<!doctype>
<html>
<head>
<link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js" integrity="sha384-Tc5IQib027qvyjSMfHjOMaLkfuWVxZxUPnCJA7l2mCWNIpG9mGCD8wGNIcPD7Txa" crossorigin="anonymous"></script>

<script>
var _validFileExtensions = [".xls", ".xlsx", ".csv"];    
function ValidateSingleInput(oInput) {
    if (oInput.type == "file") {
        var sFileName = oInput.value;
         if (sFileName.length > 0) {
            var blnValid = false;
            for (var j = 0; j < _validFileExtensions.length; j++) {
                var sCurExtension = _validFileExtensions[j];
                if (sFileName.substr(sFileName.length - sCurExtension.length, sCurExtension.length).toLowerCase() == sCurExtension.toLowerCase()) {
                    blnValid = true;
                    break;
                }
            }
             
            if (!blnValid) {
                alert("Sorry, " + sFileName + " is invalid, allowed extensions are: " + _validFileExtensions.join(", "));
                oInput.value = "";
                return false;
            }
        }
    }
    return true;
}

</script>

<style>
table, th, td {
  border-collapse: collapse;
  border-radius: 3px;
}
th, td {

    border-spacing: 5px 10px;
    padding: 20px!important;
    line-height: 1.7em;

}
</style>

</head>
<body>
<div class="container">
	<h1>Intel - Offline Question</h1>
<p class="lead">**This code reads the first two columns of an Excel file.Just upload the excel which available in the source folder(data.xlsx)**</p>
<?php
if(isset($_FILES['excel']) && $_FILES['excel']['error']==0) {
		require_once "Classes/PHPExcel.php";
		$tmpfname = $_FILES['excel']['tmp_name'];
		$excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
		$excelObj = $excelReader->load($tmpfname);
		$worksheet = $excelObj->getSheet(0);
		$lastRow = $worksheet->getHighestRow();

						
		echo "<table class=\"table table-bordered\" style=\"width:300px;\">"; 

		for ($row = $lastRow; $row >=1; $row--) { // For desending order of value
			 echo "<tr>";

			 	if ($worksheet->getCell('A'.$row)->getValue()=='') { // Empty cell print 
			      echo "<td style=\"color:white;background-color:#ffffff;border-color:#ffffff\";>"; echo $worksheet->getCell('A'.$row)->getValue();
			    }

				if ($worksheet->getCell('A'.$row)->getValue()=='2015') {
			      echo "<td style=\"color:white;background-color:#333333\";>"; echo $worksheet->getCell('A'.$row)->getValue();
			    }
			 
			    if($worksheet->getCell('A'.$row)->getValue()=='A'){
			    	echo "<td style=\"color:black;background-color:#58a65a\">"; echo $worksheet->getCell('A'.$row)->getValue();
			     
			     }
			   
			    if($worksheet->getCell('A'.$row)->getValue()=='B' || $worksheet->getCell('A'.$row)->getValue()=='C' || $worksheet->getCell('A'.$row)->getValue()=='D'){
			    	echo "<td style=\"color:black;background-color:#cccccc\">"; echo $worksheet->getCell('A'.$row)->getValue();
			     
			     }

			     if($worksheet->getCell('B'.$row)->getValue()=='2016'){
			    	echo "<td style=\"color:white;background-color:#333333\">"; echo $worksheet->getCell('B'.$row)->getValue();
			     
			     }
			   
			    if($worksheet->getCell('B'.$row)->getValue()=='E'){
			    	echo "<td style=\"color:black;background-color:#f3f21b\">"; echo $worksheet->getCell('B'.$row)->getValue();
			     
			     }

			     if($worksheet->getCell('B'.$row)->getValue()=='F'|| $worksheet->getCell('B'.$row)->getValue()=='G' || $worksheet->getCell('B'.$row)->getValue()=='H' || $worksheet->getCell('B'.$row)->getValue()=='I' || $worksheet->getCell('B'.$row)->getValue()=='J' || $worksheet->getCell('B'.$row)->getValue()=='K' || $worksheet->getCell('B'.$row)->getValue()=='L' || $worksheet->getCell('B'.$row)->getValue()=='M' || $worksheet->getCell('B'.$row)->getValue()=='N' || $worksheet->getCell('B'.$row)->getValue()=='O'){
			     	echo "<td style=\"color:black;background-color:#58a65a\">"; echo $worksheet->getCell('B'.$row)->getValue();
			     
			     }
			 echo "</td><tr>";
			 
		}
		echo "</table>";	
}


?>
<div class="container">
<form action = "" method = "POST" enctype = "multipart/form-data">
         <input type = "file" name = "excel" onchange="ValidateSingleInput(this)" /></form><br>
         <input type = "submit" value="Upload" class="btn-success" />
</form>
</div>
</div>

</body>
</html>

