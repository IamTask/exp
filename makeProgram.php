<?php
$mode = "";
if (isset($argv[1])) {
	$mode = $argv[1];
}

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use \PhpOffice\PhpSpreadsheet\IOFactory;

//getting sheet
$inputFileName = './esp.xls';
$spreadsheet = IOFactory::load($inputFileName);
$sheet = $spreadsheet->getActiveSheet();

//getting next Month text
$nextMonthTimestamp = time() + (30 * 24 * 60 * 60);
$oldLocale = setlocale(LC_TIME,'it_IT.utf8');
$nextMonth = utf8_encode(strftime("%B %Y",$nextMonthTimestamp));
$nextMonth = strtoupper($nextMonth{0}) . substr($nextMonth, 1);
echo $nextMonth;
setlocale(LC_TIME, $oldLocale);
//set next month text
$sheet->setCellValue('A11', $nextMonth);


//get latest date of program
$highestRow = $sheet->getHighestDataRow();
$latest = explode(" ", $sheet->getCellByColumnAndRow(1, $highestRow)->getValue())[1];
$settRow = 13;

function setDay($sheet,$dayText,$settRow,$settOffset,$latest,$latestOffset,$next,$diff) {
	$dayText = $dayText . ($latest+1);
	$sheet->setCellValue(("A" . ($settRow + $settOffset)), $dayText);
	if($next === $dayText) {
		return $latest;
	}

	$thisMonthTimestamp = time();
	$oldLocale = setlocale(LC_TIME,'it_IT.utf8');
	$thisMonth = utf8_encode(strftime("%B",$thisMonthTimestamp));
	setlocale(LC_TIME, $oldLocale);

	$dg = 30;
	if(
		$thisMonth === "novembre" ||
		$thisMonth === "aprile" ||
		$thisMonth === "giugno" ||
		$thisMonth === "settembre" 
	)$dg = 30;
	elseif($thisMonth === "febbraio") $dg = 28;
	else $dg=31;

	if(($latest+$diff) > $dg) {
		global $nextMonth;
		$nextMonthTimestamp = time() + (30 * 24 * 60 * 60)*2;
		$oldLocale = setlocale(LC_TIME,'it_IT.utf8');
		$nextMonth = utf8_encode(strftime("%B %Y",$nextMonthTimestamp));
		$nextMonth = strtoupper($nextMonth{0}) . substr($nextMonth, 1);

		if(
			$nextMonth === "Novembre" ||
			$nextMonth === "Aprile" ||
			$nextMonth === "Giugno" ||
			$nextMonth === "Settembre" 
		) {
			$newDate = $latest+$diff - 30 + 1;
		}
		elseif($nextMonth === "Febbraio") $newDate = $latest+$diff - 28 +1 ;
		else $newDate = $latest+$diff - 31 +1 ;
		return $newDate;
	}	 
	return $latest+$diff;
}
for ($i=0; $i < 5; $i++) { 
	$settValue = "Sett. Del " . ($latest + 1) . " " . explode(" ", $nextMonth)[0];
	$cell = ("A" . ($settRow));
	print_r ($cell);
	$sheet->setCellValue($cell, $settValue);

	$lun = "Lunedì ";
	$merc = "Mercoledì ";
	$sab = "Sabato ";
	$do = "Domenica ";

	$latest = setDay($sheet,$lun,$settRow,1,$latest,1,$merc,2);

	$latest = setDay($sheet,$merc,$settRow,3,$latest,3,$merc,0);
	$latest = setDay($sheet,$merc,$settRow,4,$latest,3,$merc,0);
	$latest = setDay($sheet,$merc,$settRow,6,$latest,3,$sab,3);	
	
	$latest = setDay($sheet,$sab,$settRow,8,$latest,6,$sab,0);
	$latest = setDay($sheet,$sab,$settRow,9,$latest,6,$do,1);
	
	$latest = setDay($sheet,$do,$settRow,11,$latest,7,$lun,1);
	
	$settRow+=13;
}



$writer = new Xlsx($spreadsheet);
$writer->save("esp.xls");
