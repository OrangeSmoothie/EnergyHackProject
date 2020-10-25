<?php 
$mysqli = new mysqli("185.21.142.28:33061", "energy", "storage", "energy");
$mysqli->set_charset("utf8");
    $mysqli->query("SET NAMES 'utf-8'");
if ($mysqli->connect_errno) {
	?>
	<html>
 <head>
  <title><?php echo date(DATE_RFC822);?></title>
 </head>
 <body>
<?php 
echo "Не удалось подключиться к БД: (" . $mysqli->connect_errno . ") " . $mysqli->connect_error; ?>
 </body>
</html>
<?php 
exit;
}

// Redirect output to a client’s web browser (Excel2007)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="simple.xlsx"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0
		
require_once './vendor/autoload.php';
error_reporting(E_ALL);

/** PHPExcel */
//require_once '../Classes/PHPExcel.php';
// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

// Set properties
//echo date('H:i:s') . " Set properties\n";
$objPHPExcel->getProperties()->setCreator("Assembler-inserts")
->setLastModifiedBy("Assembler-inserts")
->setTitle("Office 2007 XLSX Test Document")
->setSubject("Office 2007 XLSX Test Document")
->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
->setKeywords("office 2007 openxml php")
->setCategory("EH3");
// Add some data








//echo date('H:i:s') . " Add some data\n";



$sheet = $objPHPExcel->setActiveSheetIndex(0);
$orgID_GET = !empty($_GET['orgID']) ? $_GET['orgID'] : 3;
$orgNAME_GET = "";

$num_curr_sheet = 0;


$query  = "SELECT id, Name FROM Organization WHERE id = ".$orgID_GET." LIMIT 1";
$query .= "SELECT id, Name FROM Branch WHERE id_Organization = ".$orgID_GET." ORDER BY id ASC";

$mysqli->real_query("SELECT * FROM Organization WHERE id = ".$orgID_GET." LIMIT 1");
$res = $mysqli->use_result();

if($row = $res->fetch_assoc()) {
    $orgNAME_GET=$row['Name'];
}
$mysqli->close();


$sheet = $objPHPExcel->setActiveSheetIndex(0);
	//HEAD LEP
$sheet
->setCellValue('A1', 'Наименование РЭС')
->setCellValue('B1', '№ п.п.')
->setCellValue('C1', 'Подстанции')
->setCellValue('C2', 'Наименование')
->setCellValue('D2', 'Класс напряжения')
->setCellValue('E2', 'Год ввода')
->setCellValue('F1', 'Трансформаторы')
->setCellValue('F2', 'Тип')
->setCellValue('G2', 'Номинальная мощность, МВА')
->setCellValue('H2', 'Номинальная мощность, МВА')
->setCellValue('I2', 'Загрузка (зимний максимум), %')
->setCellValue('J2', 'Год изготовления')
->setCellValue('K2', 'Год включения')
->setCellValue('L2', '% износа* ')
->setCellValue('M2', 'Состояние')
->setCellValue('N1', 'Срок службы на 01.01.2020')
->setCellValue('N2', 'с года ввода ПС')
->setCellValue('O2', 'с года изготов.');
$sheet
->mergeCells('A1:A5')
->mergeCells('B1:B5')
->mergeCells('C1:E1')
->mergeCells('C2:C5')
->mergeCells('D2:D5')
->mergeCells('E2:E5')
->mergeCells('F1:M1')
->mergeCells('F2:F5')
->mergeCells('G2:G5')
->mergeCells('H2:H5')
->mergeCells('I2:I5')
->mergeCells('J2:J5')
->mergeCells('K2:K5')
->mergeCells('L2:L5')
->mergeCells('M2:M5')
->mergeCells('N1:O1')
->mergeCells('N2:N5')
->mergeCells('O2:O5');
$borderStyle = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN)));
$sheet->getStyle('A1:O5')->applyFromArray($borderStyle);
$sheet->getStyle('A1:O5')->getAlignment()->setWrapText(true);
$sheet->getStyle('A1:B1')->getAlignment()->setTextRotation(90);
$sheet->getStyle('C2:O2')->getAlignment()->setTextRotation(90);

$sheet->getStyle('C1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$sheet->getStyle('F1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

/////////////////
$startLine =  7; // Начальная координата y
$sheet->setCellValueByColumnAndRow(0, 6, $orgNAME_GET );
$sheet->mergeCells('A6:O6');
$sheet->getStyle('A6')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$mysqli = new mysqli("185.21.142.28:33061", "energy", "storage", "energy");
$mysqli->set_charset("utf8");
$mysqli->query("SET NAMES 'utf-8'");
$mysqli->real_query("SELECT * FROM Branch WHERE id_Organization = ".$orgID_GET."");
$resBranch = $mysqli->use_result();
while ($rowBranch = $resBranch->fetch_assoc()) {
	
	$sheet->setCellValueByColumnAndRow(0, $startLine, $rowBranch['Name']);
$sheet->mergeCells("A".($startLine).":O".($startLine));
$sheet->getStyle("A".($startLine))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$columnPosition = 0; // Начальная координата x
$startLine++;
///////////////////////
$mysqli3 = new mysqli("185.21.142.28:33061", "energy", "storage", "energy");
$mysqli3->set_charset("utf8");
$mysqli3->query("SET NAMES 'utf-8'");
$mysqli3->real_query("SELECT * FROM Substations WHERE id_Branch =".$rowBranch['id']);
$resD = $mysqli3->use_result();
while ($rowD = $resD->fetch_assoc()) {
 $sheet->setCellValueByColumnAndRow(0, $startLine, $rowD['NumberPP']);
 $sheet->setCellValueByColumnAndRow(1, $startLine, $rowD['NameRES']);
 $sheet->setCellValueByColumnAndRow(2, $startLine, $rowD['NameSubstation']);
 $sheet->setCellValueByColumnAndRow(3, $startLine, $rowD['ClassVoltage']);
 $sheet->setCellValueByColumnAndRow(4, $startLine, $rowD['YearInput']);
 $sheet->setCellValueByColumnAndRow(5, $startLine, $rowD['TransNumbe']);
 $sheet->setCellValueByColumnAndRow(6, $startLine, $rowD['TransType']);
 $sheet->setCellValueByColumnAndRow(7, $startLine, $rowD['TransNomPower']);
 $sheet->setCellValueByColumnAndRow(8, $startLine, $rowD['TransLoadWinter']);
 $sheet->setCellValueByColumnAndRow(9, $startLine, $rowD['TransYearManufacture']);
 $sheet->setCellValueByColumnAndRow(10, $startLine, $rowD['TransYearOn']);
 $sheet->setCellValueByColumnAndRow(11, $startLine, $rowD['TransPercentWear']);
 $sheet->setCellValueByColumnAndRow(12, $startLine, $rowD['TransCondition']);
 $sheet->setCellValueByColumnAndRow(13, $startLine, 2020-$rowD['YearInput']);
 $sheet->setCellValueByColumnAndRow(14, $startLine, 2020-$rowD['TransYearManufacture']);
	$startLine++;
}


	$mysqli3->close();
///////////////////////




}
	//$sheet = $objPHPExcel->setActiveSheetIndex($num_curr_sheet); // Выбираем первый лист в документе
	//$objPHPExcel->getActiveSheet()->setTitle($rowBranch['Name']);
$sheet->getStyle("A6:O".($startLine-1))->applyFromArray($borderStyle);

$mysqli->close();
	

	//$this->xls->createSheet();
	//$num_curr_sheet++;









/* закрываем соединение */

// Rename sheet
//echo date('H:i:s') . " Rename sheet\n";

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);
// Save Excel 2007 file
//echo date('H:i:s') . " Write to Excel2007 format\n";
// Save Excel 2007 file
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');
?>
