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

$mysqli = new mysqli("185.21.142.28:33061", "energy", "storage", "energy");
$mysqli->set_charset("utf8");
$mysqli->query("SET NAMES 'utf-8'");
$mysqli->real_query("SELECT * FROM Branch WHERE id_Organization = ".$orgID_GET."");
$resBranch = $mysqli->use_result();

while ($rowBranch = $resBranch->fetch_assoc()) {
$sheet = $objPHPExcel->setActiveSheetIndex($num_curr_sheet);
	//HEAD LEP
$sheet
->setCellValue('A1', '№ п.п.')
->setCellValue('B1', 'Диспетчерский номер ЛЭП')
->setCellValue('C1', 'Наименование ЛЭП')
->setCellValue('D1', 'Напряжение, кВ')
->setCellValue('E1', 'Год ввода в эксплуатацию')
->setCellValue('F1', 'Провод, кабель')
->setCellValue('F2', 'Количество цепей')
->setCellValue('G2', 'Длина всего, км')
->setCellValue('G3', 'по трассе')
->setCellValue('H3', 'на 1 цепь')
->setCellValue('I2', 'Длина в т.ч. по участкам, км')
->setCellValue('I3', 'по трассе')
->setCellValue('J3', 'на 1 цепь')
->setCellValue('K2', 'Марка')
->setCellValue('L1', 'Техническое состояние')
->setCellValue('M1', 'Срок службы ЛЭП на 01.01.2020');
$sheet
->mergeCells('A1:A5')
->mergeCells('B1:B5')
->mergeCells('C1:C5')
->mergeCells('D1:D5')
->mergeCells('E1:E5')
->mergeCells('F1:K1')
->mergeCells('F2:F5')
->mergeCells('G2:H2')
->mergeCells('G3:G5')
->mergeCells('H3:H5')
->mergeCells('I2:J2')
->mergeCells('I3:I5')
->mergeCells('J3:J5')
->mergeCells('K2:K5')
->mergeCells('L1:L5')
->mergeCells('M1:M5');
$borderStyle = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN)));
$sheet->getStyle('A1:M5')->applyFromArray($borderStyle);
$sheet->getStyle('A1:M5')->getAlignment()->setWrapText(true);
$sheet->getStyle('A1:B1')->getAlignment()->setTextRotation(90);
$sheet->getStyle('D1:E1')->getAlignment()->setTextRotation(90);
$sheet->getStyle('F2')->getAlignment()->setTextRotation(90);
$sheet->getStyle('G3:J3')->getAlignment()->setTextRotation(90);
$sheet->getStyle('K2')->getAlignment()->setTextRotation(90);
$sheet->getStyle('L1:M1')->getAlignment()->setTextRotation(90);
$sheet->getStyle('F1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
/////////////////


	$sheet->setCellValueByColumnAndRow(0, 6, $orgNAME_GET );
$sheet->mergeCells('A6:M6');
	$sheet->setCellValueByColumnAndRow(0, 7, $rowBranch['Name']);
$sheet->mergeCells('A7:M7');
$sheet->getStyle('A6:A7')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$columnPosition = 0; // Начальная координата x


	
$mysqli3 = new mysqli("185.21.142.28:33061", "energy", "storage", "energy");
$mysqli3->set_charset("utf8");
$mysqli3->query("SET NAMES 'utf-8'");
$mysqli3->real_query("SELECT * FROM PowerLines, mm_PowerLines_Branch WHERE PowerLines.id = mm_PowerLines_Branch.id_PowerLines AND mm_PowerLines_Branch.id_Branch =".$rowBranch['id']);
$resD = $mysqli3->use_result();
$startLine =  8; // Начальная координата y
while ($rowD = $resD->fetch_assoc()) {
    



 $sheet->setCellValueByColumnAndRow(0, $startLine, $rowD['NumberPP']);
 $sheet->setCellValueByColumnAndRow(1, $startLine, $rowD['NumberPowerLines']);
 $sheet->setCellValueByColumnAndRow(2, $startLine, $rowD['NamePowerLines']);
 $sheet->setCellValueByColumnAndRow(3, $startLine, $rowD['Voltage']);
 $sheet->setCellValueByColumnAndRow(4, $startLine, $rowD['CommissioningYear']);
 $sheet->setCellValueByColumnAndRow(5, $startLine, $rowD['CountСircuit']);
 $sheet->setCellValueByColumnAndRow(6, $startLine, $rowD['trackLengthALL']);
 $sheet->setCellValueByColumnAndRow(7, $startLine, $rowD['LenghtPerСircuitALL']);
 $sheet->setCellValueByColumnAndRow(8, $startLine, $rowD['trackLength']);
 $sheet->setCellValueByColumnAndRow(9, $startLine, $rowD['LenghtPerСircuit']);
 $sheet->setCellValueByColumnAndRow(10, $startLine, $rowD['Model']);
 $sheet->setCellValueByColumnAndRow(11, $startLine, $rowD['TechnicalCondition']);
 $sheet->setCellValueByColumnAndRow(12, $startLine, 2020-$rowD['CommissioningYear']);
	
	$startLine++;
	}
	
	$sheet->getStyle("A6:M".($startLine-1))->applyFromArray($borderStyle);
	$objPHPExcel->getActiveSheet()->setTitle($rowBranch['Name']);
	$objPHPExcel->createSheet();
	$num_curr_sheet++;	
}
	//$sheet = $objPHPExcel->setActiveSheetIndex($num_curr_sheet); // Выбираем первый лист в документе
	//$objPHPExcel->getActiveSheet()->setTitle($rowBranch['Name']);


$mysqli->close();
$mysqli3->close();	

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
