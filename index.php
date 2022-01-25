<?php
set_time_limit(0);
$con=mysqli_connect('localhost','root','','importDB');
if(isset($_POST['submit'])){
	$time_start = microtime(true);
	$file=$_FILES['doc']['tmp_name'];
	
	$ext=pathinfo($_FILES['doc']['name'],PATHINFO_EXTENSION);
	if($ext=='xls'){
		require('PHPExcel/PHPExcel.php');
		require('PHPExcel/PHPExcel/IOFactory.php');
		
		
		$obj=PHPExcel_IOFactory::load($file);
		foreach($obj->getWorksheetIterator() as $sheet){
			$getHighestRow=$sheet->getHighestRow();
			if($getHighestRow > 1){
				$query = "insert into community(report_name,report_number,version_number,hub,community,community_name,source,source_report_name,source_type,combiner,ITW,program_name,Dat,day_of_event,currency,total_sales,refund,net_sales,commission,net_pool,net_add_in,minus_break,breakage) values ";
			}
			// if header exists $i will initiliase with 2 otherwise 1
			for($i=2;$i<=$getHighestRow;$i++){
				$report_name=$sheet->getCellByColumnAndRow(0,$i)->getValue();
				$report_number=$sheet->getCellByColumnAndRow(1,$i)->getValue();
				$version_number=$sheet->getCellByColumnAndRow(2,$i)->getValue();
				$hub=$sheet->getCellByColumnAndRow(3,$i)->getValue();
				$community=$sheet->getCellByColumnAndRow(4,$i)->getValue();
				$community_name=$sheet->getCellByColumnAndRow(5,$i)->getValue();
				$source=$sheet->getCellByColumnAndRow(6,$i)->getValue();
				$source_report_name=$sheet->getCellByColumnAndRow(7,$i)->getValue();
				$source_type=$sheet->getCellByColumnAndRow(8,$i)->getValue();
				$combiner=$sheet->getCellByColumnAndRow(9,$i)->getValue();
				$ITW=$sheet->getCellByColumnAndRow(10,$i)->getValue();
				$program_name=$sheet->getCellByColumnAndRow(11,$i)->getValue();
				$Dat=$sheet->getCellByColumnAndRow(12,$i)->getValue();
				$day_of_event=$sheet->getCellByColumnAndRow(13,$i)->getValue();
				$currency=$sheet->getCellByColumnAndRow(14,$i)->getValue();
				$total_sales=$sheet->getCellByColumnAndRow(15,$i)->getValue();
				$refund=$sheet->getCellByColumnAndRow(16,$i)->getValue();
				$net_sales=$sheet->getCellByColumnAndRow(17,$i)->getValue();
				$commission=$sheet->getCellByColumnAndRow(18,$i)->getValue();
				$net_pool=$sheet->getCellByColumnAndRow(19,$i)->getValue();
				$net_add_in=$sheet->getCellByColumnAndRow(20,$i)->getValue();
				$minus_break=$sheet->getCellByColumnAndRow(21,$i)->getValue();
				$breakage=$sheet->getCellByColumnAndRow(22,$i)->getValue();
				// if($getHighestRow > 1){
				// 	$query = "insert into community(report_name,report_number,version_number,hub,community,community_name,source,source_report_name,source_type,combiner,ITW,program_name,Dat,day_of_event,currency,total_sales,refund,net_sales,commission,net_pool,net_add_in,minus_break,breakage) values ('$report_name','$report_number','$version_number','$hub','$community','$community_name','$source','$source_report_name','$source_type','$combiner','$ITW','$program_name','$Dat','$day_of_event','$currency','$total_sales','$refund','$net_sales','$commission','$net_pool','$net_add_in','$minus_break','$breakage');";
				// }
				// mysqli_query($con,$query);
				if ($i == $getHighestRow){
                    $query .= "('$report_name','$report_number','$version_number','$hub','$community','$community_name','$source','$source_report_name','$source_type','$combiner','$ITW','$program_name','$Dat','$day_of_event','$currency','$total_sales','$refund','$net_sales','$commission','$net_pool','$net_add_in','$minus_break','$breakage');";
                }else{
					$query .= "('$report_name','$report_number','$version_number','$hub','$community','$community_name','$source','$source_report_name','$source_type','$combiner','$ITW','$program_name','$Dat','$day_of_event','$currency','$total_sales','$refund','$net_sales','$commission','$net_pool','$net_add_in','$minus_break','$breakage'),";
                }
			}

			mysqli_query($con,$query);
		}
	}else{
		echo "Invalid file format";
	}
	$time_end = microtime(true);
	$execution_time = ($time_end - $time_start);
	echo '<b>Total Execution Time:</b> '.($execution_time*1000).'Milliseconds';
}
?>
<form method="post" enctype="multipart/form-data">
	<input type="file" name="doc"/>
	<input type="submit" name="submit"/>
</form>