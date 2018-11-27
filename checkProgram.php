<?php
$mode = "";
if (isset($argv[2])) {
	$mode = $argv[2];
}

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use \PhpOffice\PhpSpreadsheet\IOFactory;
$inputFileName = "./" . $argv[1];

$uniquePeople = [
"Filetti Gabriella",
"Rapisarda Giuseppe",
"Ninzola Manuel",
"Catania Vincenzo",
"Tascone Danilo",
"Sciuto Giulia",
"Catania Marisa",
"La China Laura",
"Castiglione Grazia",
"Franco Marco",
"Alì Giuseppe",
"Cuzzupè Angelo",
"Turiano Chiara",
"Franco Viviana",
"Torrisi Eugenio",
"Margaglio Francesco",
"Franceschini Paola",
"Torrisi Elisa",
"Capizzi Vincenzo",
"Caruso Nunzio",
"Pafumi Angela",
"Lombardo Micol",
"Motta Paola",
"Alonzo Beatrice",
"Di Martino Mario",
"Sciuto Alfio",
"Lena Claudio",
"Mirabella Salvo",
"La China Gabriele",
"Malnowicz Eva",
"Battiato Cettina",
"Pannuzzo Giovanni",
"Pannuzzo Pina",
"Lena Gabriella",
"Sciuto Maria Grazia",
"Porto Piero",
"Graci Marcello",
"Lombardo Rosalba",
"Bartilotti Tiziana",
"Lombardo Giosè",
"Motta Simone",
"Groppi Mariella",
"Chillemi Cettina",
"Filetti Gabriella",
"Alì Daniela",
"Capizzi Antonella",
"Ninzola Viviana",
"Malnowicz Gregorio",
"Battiato Beniamino",
"Compagnini Adriano",
"Alì Samuele",
"Caruso Giuseppe",
"Porto Denise",
"Trovato Luana",
"Margaglio Marinella",
"Nava Giovanna",
"Motta Elena",
"Giaquinta Marika"
];
/** Load $inputFileName to a Spreadsheet Object  **/
$spreadsheet = IOFactory::load($inputFileName);
$sheet = $spreadsheet->getActiveSheet();

//$sheet->setCellValue('A1', 'esp!');

$rows = [];
$highestRow = $sheet->getHighestDataRow();
$highestColumn = $sheet->getHighestDataColumn();
echo $highestColumn . "\n";
echo $highestRow . "\n";

$people = [];
$counter = 1;
$doubles = [];
//retrieving names from excel sheet
for ($row = 16; $row <= $highestRow; ++$row) {
    for ($col = 4; $col <= 7; ++$col) {
        $person = $sheet->getCellByColumnAndRow($col, $row)->getValue();
        			
        if($person != "") {

        	/*
	        echo $counter . ")|";
	        print_r($person);
	        echo "|\n";
	        */
	        $counter++;
	        if(!in_array($person, $people)) array_push($people, $person);
	        elseif (array_key_exists($person, $doubles)) $doubles[$person]++;
	        else $doubles[$person] = 2;
        }
    }

}

//print retrieved people list
if($mode === "pr") {
	print_r($people);
	echo " \n No. Of People: ";
	print_r(sizeof($people));
}


//print missing
if($mode === "mi") {
	$missing = [];
	foreach ($uniquePeople as $person) {
		//print_r($person);
		if( in_array($person, $people) === false) $missing[] = $person;
	}
	print_r($missing);
}

if($mode === "do") {
	print_r($doubles);
}

if($mode === "NoN") print_r("No. of Names: " . $counter);

/*
$writer = new Xlsx($spreadsheet);
$writer->save("es.xlsx");
*/

/*
//sample
foreach ($sheet->getRowIterator() AS $row) {
    $cellIterator = $row->getCellIterator();
    $cellIterator->setIterateOnlyExistingCells(FALSE); 
    // This loops through all cells,
    $cells = [];
    foreach ($cellIterator as $cell) {
        if($cell-> )$cells[] = $cell->getValue();
    }
    $rows[] = $cells;
}
*/
/*
foreach ($rows as $column) {
	foreach ($column as $cell) {
		if($cell != "") echo " " . $cell . " ";
	}
}
*/
?>


