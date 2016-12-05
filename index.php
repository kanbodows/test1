<?php
// либа для работы с excel
require_once  __DIR__.'/PHPExcel/PHPExcel/IOFactory.php';

// получаем и читаем файл
$inputFileName = __DIR__.'/test.xlsx';
try {
    $inputFileType = PHPExcel_IOFactory::identify($inputFileName);
    $objReader     = PHPExcel_IOFactory::createReader($inputFileType);
    $objPHPExcel   = $objReader->load($inputFileName);
} catch(Exception $e) {
    die('Error loading file "'.pathinfo($inputFileName,PATHINFO_BASENAME).'": '.$e->getMessage());
}

$sheet = $objPHPExcel->getSheet(0);
$highestRow = $sheet->getHighestRow();
$highestColumn = $sheet->getHighestColumn();

// коннект к базе
$con = mysql_connect('localhost','root','');
if (!$con)
  die('Could not connect: ' . mysql_error());
mysql_select_db('conradki_conrad2015', $con);

// получаем строку с колонками
$columns = $sheet->rangeToArray('A1' . ':' . $highestColumn . '1', NULL, TRUE, FALSE)[0];
$colIndexes = array();
// узаем индексы нужных колонок
foreach ($columns as $key => $col) {
	var_dump($col);
	if($col == 'IDP' || $col == 'ID проекта' || $col == 'идп')
		$colIndexes['id'] = $key;
	else if($col == 'Менеджер')
		$colIndexes['manager'] = $key;
	else if($col == 'Проект')
		$colIndexes['project'] = $key;
}
// вытаскаем записи из базы
$result = mysql_query('SELECT * from test_table');
while($rowDB = mysql_fetch_array($result)){
    // проходимся по каждой строке в файле (расчет что в файле не слишком много записей, иначе надо рефакторить)
    for ($row = 2; $row <= $highestRow; $row++){
        $data = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, NULL, TRUE, FALSE)[0];
        // если совпали id то записываем
    	if($data[$colIndexes['id']] == $rowDB['id']){
			$objPHPExcel->getActiveSheet()->SetCellValue(getNameFromNumber($colIndexes['manager']).$row, $rowDB['manager']);
			$objPHPExcel->getActiveSheet()->SetCellValue(getNameFromNumber($colIndexes['project']).$row, $rowDB['project']);
			continue;
    	}
	}
}
// сохраняем
$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
$objWriter->save('test.xlsx');

echo 'готово';
// функция получения буквы колонки от индекса
function getNameFromNumber($num) {
    $numeric = $num % 26;
    $letter = chr(65 + $numeric);
    $num2 = intval($num / 26);
    if ($num2 > 0)
        return getNameFromNumber($num2 - 1) . $letter;
    else
        return $letter;
}
?>