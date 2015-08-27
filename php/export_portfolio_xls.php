<?php
/** Error reporting */
error_reporting(E_ALL);
ini_set('max_execution_time',6000);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

/** Include PHPExcel */
require_once '../tools/PHPExcel/Classes/PHPExcel.php';
require_once '../tools/PHPExcel/Classes/PHPExcel/IOFactory.php';

/** Include Define **/
require_once("inc_db.php");

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

// Set document properties
$objPHPExcel->getProperties()->setCreator("netpoint-media")
							 ->setLastModifiedBy("netpoint-media")
							 ->setTitle("PHPExcel Test Document")
							 ->setSubject("PHPExcel Test Document")
							 ->setDescription("Test document for PHPExcel, generated using PHP classes.")
							 ->setKeywords("office PHPExcel php")
							 ->setCategory("Test result file");

$objPHPExcel2 = PHPExcel_IOFactory::load("excel/deckblatt.xlsx");
foreach ($objPHPExcel2->getAllSheets() as $worksheet) {
		$objPHPExcel->AddExternalSheet($worksheet);
}

$objPHPExcel->setActiveSheetIndex(2);
//Clone worksheet
$sheet2 = $objPHPExcel->getActiveSheet()->copy();

//Add Clone worksheet
$clone = clone $sheet2;
$clone->setTitle('clone');
$objPHPExcel->addSheet($clone);

$status = true;
$subchannel = null;
$portid = null;
$portcount2 = null;
$portcount = null;

/* Style */
$linkStyle = array(
	'font' => array(
			'underline' => 'single',
			'color' => array ('rgb' => '0000FF'),
			'name' 	=> getSchriftart()
	)
);

$menuStyle = array(
	'font'  => array(
			'bold'  => false,
			'color' => array('rgb' => 'FFFFFF'),
			'name' 	=> getSchriftart()
	),
	'fill' => array(
					'type' => PHPExcel_Style_Fill::FILL_SOLID,
					'color' => array('rgb' => getMenuColor())
			),
	'borders' => array(
					'allborders' => array(
							'style' => PHPExcel_Style_Border::BORDER_THIN
							)
			)
);
$greyCellBackroundStyle = array(
	'fill' => array(
					'type' => PHPExcel_Style_Fill::FILL_SOLID,
					'color' => array('rgb' => 'E8E8E8')
			),
	'borders' => array(
					'allborders' => array(
							'style' => PHPExcel_Style_Border::BORDER_THIN,
							'color' => array('rgb' => 'FFFFFF')
							)
	),
	'font' => array(
		'name' 	=> getSchriftart()
	)
);
$center = array(
	'alignment' => array(
			'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
			'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
	)
);
$styleUnderLine = array(
			'font' => array(
				'underline' => PHPExcel_Style_Font::UNDERLINE_SINGLE,
				'name' 	=> getSchriftart()
			)
);
$addFont = array(
			'font' => array(
				'name' 	=> getSchriftart()
			)
);
function createxls($show_imp = false,$show_view = false,$show_unique = false,	$strPath)
{
		global $objPHPExcel;
		global $clone;
		global $status;
		global $sheetEx;
		global $linkStyle;
		global $menuStyle;
		global $greyCellBackroundStyle;
		global $center;
		global $styleUnderLine;
		global $addFont;

		if($status==false){
			//echo "status: false",EOL;
			$objPHPExcel = new PHPExcel();

			// Set document properties
			//echo date('H:i:s') , " Set document properties" , EOL;
			$objPHPExcel->getProperties()->setCreator("netpoint-media")
										 ->setLastModifiedBy("netpoint-media")
										 ->setTitle("PHPExcel Test Document")
										 ->setSubject("PHPExcel Test Document")
										 ->setDescription("Test document for PHPExcel, generated using PHP classes.")
										 ->setKeywords("office PHPExcel php")
										 ->setCategory("Test result file");

			$objPHPExcel2 = PHPExcel_IOFactory::load("excel/deckblatt.xlsx");
			foreach ($objPHPExcel2->getAllSheets() as $worksheet) {
					$objPHPExcel->AddExternalSheet($worksheet);
			}

			$objPHPExcel->setActiveSheetIndex(2);
			//Clone worksheet index 2
			$sheet2 = $objPHPExcel->getActiveSheet()->copy();

			//Add Clone worksheet
			$clone = clone $sheet2;
			$clone->setTitle('clone');
			$objPHPExcel->addSheet($clone);

		}

		//Arbeitsblatt technik
		$objPHPExcel->setActiveSheetIndex(2);
		{
		  $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(36);
		  $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(58);
		  $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(11);
		  $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(11);
		  $objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight(36);

		  $objPHPExcel->getActiveSheet()->mergeCells('A2:D2');
		  $objPHPExcel->getActiveSheet()->mergeCells('B3:D3');
		  $objPHPExcel->getActiveSheet()->mergeCells('B4:D4');
		  $objPHPExcel->getActiveSheet()->mergeCells('B5:D5');
		  $objPHPExcel->getActiveSheet()->mergeCells('A7:D7');
		  $objPHPExcel->getActiveSheet()->mergeCells('B8:D8');
		  $objPHPExcel->getActiveSheet()->mergeCells('A8:A9');

		  $objPHPExcel->getActiveSheet()->getCell('A2')->setValue("SPEZIFIKATION & ADSERVER");
		  $objPHPExcel->getActiveSheet()->getStyle('A2:D2')->applyFromArray($menuStyle);
			$objPHPExcel->getActiveSheet()->getStyle('A2:D2')->applyFromArray($center);

			$r = 2;
			$r++;
			$objPHPExcel->getActiveSheet()->getCell('A'.$r++)->setValue("URL Werbemittelspezifikationen");
			$objPHPExcel->getActiveSheet()->getCell('A'.$r++)->setValue("E-Mail Banner-Anlieferungsadresse");
			$objPHPExcel->getActiveSheet()->getCell('A'.$r++)->setValue("Adserver Hersteller / Typ / Version");

			$r++;
			$objPHPExcel->getActiveSheet()->getCell('A'.$r)->setValue("SPEZIFIKATIONEN DER WICHTIGSTEN STANDARD-FORMATE");
			$objPHPExcel->getActiveSheet()->getStyle('A7:D7')->applyFromArray($menuStyle);
			$objPHPExcel->getActiveSheet()->getStyle('A2:D2')->applyFromArray($center);

			$r++;
			$objPHPExcel->getActiveSheet()->getCell('A'.$r)->setValue("Format");

			$r=2;
			$r++;
			$objPHPExcel->getActiveSheet()->getCell('B'.$r)->setValue("http://www.netpoint-media.de/werbeformen/spezifikationen.html");
			//change the data type of the cell
			$objPHPExcel->getActiveSheet()->getCell('B'.$r)->setDataType(PHPExcel_Cell_DataType::TYPE_STRING2);
			//now set the link
			$objPHPExcel->getActiveSheet()->getCell('B'.$r)->getHyperlink()->setUrl(strip_tags("http://www.netpoint-media.de/werbeformen/spezifikationen.html"));

			$r++;
			$objPHPExcel->getActiveSheet()->getCell('B'.$r)->setValue("banner@netpoint-media.de");
			//change the data type of the cell
			$objPHPExcel->getActiveSheet()->getCell('B'.$r)->setDataType(PHPExcel_Cell_DataType::TYPE_STRING2);
			//now set the link
			$objPHPExcel->getActiveSheet()->getCell('B'.$r)->getHyperlink()->setUrl(strip_tags("mailto:banner@netpoint-media.de"));
			// Set url color
			$objPHPExcel->getActiveSheet()->getStyle('B'.$r)->applyFromArray($linkStyle);

			$r++;
			$objPHPExcel->getActiveSheet()->getCell('B'.$r)->setValue("ADTECH HELIOS IQ");

			$r=8;
			$objPHPExcel->getActiveSheet()->getCell('B'.$r++)->setValue("Max. Dateigewicht in KB");
			$objPHPExcel->getActiveSheet()->getCell('B'.$r++)->setValue("Größe in Pixel");
			$objPHPExcel->getActiveSheet()->getCell('C9')->setValue("GIF, JPG");
			$objPHPExcel->getActiveSheet()->getCell('D9')->setValue("Flash");

			$objPHPExcel->getActiveSheet()->getStyle("A3:D6")->applyFromArray($addFont);
			$objPHPExcel->getActiveSheet()->getStyle("A8:D9")->applyFromArray($center);

			$result = mysql_query("SELECT * FROM werbeformen WHERE werbeformen.online = '1' ORDER BY werbeformen.sort");

			$counter = 0;
			while($row = @mysql_fetch_array($result))
			{
				$objPHPExcel->getActiveSheet()->getCell('A'.$r++)->setValue($row['name']);
				$objPHPExcel->getActiveSheet()->getCell('B'.($r-1))->setValue($row['format']);
				$objPHPExcel->getActiveSheet()->getCell('C'.($r-1))->setValue($row['gew']);
				$objPHPExcel->getActiveSheet()->getCell('D'.($r-1))->setValue($row['gewflash']);
				$counter++;
			}

			$objPHPExcel->getActiveSheet()->getStyle("C10:D56")->applyFromArray($center);

			$maxrow = $counter + 9;
			$objPHPExcel->getActiveSheet()->getStyle("A8:D".$maxrow)->applyFromArray($greyCellBackroundStyle);

		}

		//Add worksheet Portfolio
		$sheet3 = clone $clone;
		$sheet3->setTitle('portfolio');
		$objPHPExcel->addSheet($sheet3);

		//Arbeitsblatt Portfolio
		$objPHPExcel->setActiveSheetIndex(4);
		{
			$objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight(36);
			$objPHPExcel->getActiveSheet()->getRowDimension(2)->setRowHeight(25.5);

			$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(36);
			$objPHPExcel->getActiveSheet()->getCell('A2')->setValue("Site Name");
			$c = 0;
			$r = 2;
			if($show_imp)
			{
				$c++;
				$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
			  	$objPHPExcel->getActiveSheet()->getColumnDimension($colIndex)->setWidth(16);
				$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue("PageImpressions pro Monat");
				$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getAlignment()->setWrapText(true);
			}
			if($show_view)
			{
				$c++;
				$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
				$objPHPExcel->getActiveSheet()->getColumnDimension($colIndex)->setWidth(16);
				$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue("Visits pro Monat");
			}
			if($show_unique)
			{
				$c++;
				$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
				$objPHPExcel->getActiveSheet()->getColumnDimension($colIndex)->setWidth(16);
				$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue("Unique User pro Monat");
				$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getAlignment()->setWrapText(true);
			}
			$c++;
			$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
			$objPHPExcel->getActiveSheet()->getColumnDimension($colIndex)->setWidth(40);
			$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue("Website- & Zielgruppenbeschreibung / Buchungsmöglichkeiten & Preise");
			$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getAlignment()->setWrapText(true);

			$objPHPExcel->getActiveSheet()->getStyle("A2:E2")->applyFromArray($center);

			$result = mysql_query("SELECT * FROM portfolio WHERE portfolio.status = 'online' ORDER BY Website");

			$counter = 0;
			$r=3;
			//$r++;
			while($row = @mysql_fetch_array($result))
			{
				$c=0;
				$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row['Website']);
				if($show_imp)
				{
					$c++;
					$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row['PI']);
					$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
					$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getNumberFormat()->setFormatCode('#,###');

				}
				if($show_view)
				{
					$c++;
					$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row['visits']);
					$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
					$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getNumberFormat()->setFormatCode('#,###');
				}
				if($show_unique)
				{
					$c++;
					$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row['uniqueuser']);
					$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
					$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getNumberFormat()->setFormatCode('#,###');
				}
				$c++;
				$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row['Website']);
				//change the data type of the cell
				$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setDataType(PHPExcel_Cell_DataType::TYPE_STRING2);
				//now set the link
				$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->getHyperlink()->setUrl(strip_tags('http://www.netpoint-media.de/portfolio/'.$row['Website'].'.html'));
				// Set url color
				$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
				$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->applyFromArray($linkStyle);
				$r++;
				$counter++;
			}
			$maxrow = $counter + 2;
			$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
			$objPHPExcel->getActiveSheet()->getStyle("A3:".$colIndex.$maxrow)->applyFromArray($greyCellBackroundStyle);

		 	$objPHPExcel->getActiveSheet()->getStyle('A2:'.$colIndex.'2')->applyFromArray($menuStyle);
			$objPHPExcel->getActiveSheet()->getStyle('A2:'.$colIndex.'2')->applyFromArray($center);

			/* Summe */
			$c = 0;
			if($show_imp)
			{
				$c++;
				$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
				$objPHPExcel->getActiveSheet()->getCell($colIndex.($maxrow+1))->setValue('=SUM('.$colIndex.'3:'.$colIndex.$maxrow.')');
				$objPHPExcel->getActiveSheet()->getStyle($colIndex.($maxrow+1))->applyFromArray($styleUnderLine);
				$objPHPExcel->getActiveSheet()->getStyle($colIndex.($maxrow+1))->getNumberFormat()->setFormatCode('#,###');
			}
			if($show_view)
			{
				$c++;
				$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
				$objPHPExcel->getActiveSheet()->getCell($colIndex.($maxrow+1))->setValue('=SUM('.$colIndex.'3:'.$colIndex.$maxrow.')');
				$objPHPExcel->getActiveSheet()->getStyle($colIndex.($maxrow+1))->applyFromArray($styleUnderLine);
				$objPHPExcel->getActiveSheet()->getStyle($colIndex.($maxrow+1))->getNumberFormat()->setFormatCode('#,###');
			}
			if($show_unique)
			{
				$c++;
				$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
				$objPHPExcel->getActiveSheet()->getCell($colIndex.($maxrow+1))->setValue('=SUM('.$colIndex.'3:'.$colIndex.$maxrow.')');
				$objPHPExcel->getActiveSheet()->getStyle($colIndex.($maxrow+1))->applyFromArray($styleUnderLine);
				$objPHPExcel->getActiveSheet()->getStyle($colIndex.($maxrow+1))->getNumberFormat()->setFormatCode('#,###');
			}
		}

		//Add worksheet netpoint-rotation
		$sheet4 = clone $clone;
		$sheet4->setTitle('channel');
		$objPHPExcel->addSheet($sheet4);

		//Arbeitsblatt Channel
		$objPHPExcel->setActiveSheetIndex(5);
		{
			$objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight(36);
			$objPHPExcel->getActiveSheet()->getRowDimension(2)->setRowHeight(25.5);
			$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(36);
		  $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(16);
		  $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(9);
		  $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(22);

		  $objPHPExcel->getActiveSheet()->getCell('A2')->setValue('Themen-Rotation');
			$objPHPExcel->getActiveSheet()->getCell('B2')->setValue('PI in Mio. / Monat');
		  $objPHPExcel->getActiveSheet()->getStyle('A2:B2')->applyFromArray($menuStyle);
			$objPHPExcel->getActiveSheet()->getStyle('A2:B2')->applyFromArray($center);

			$result = mysql_query("SELECT SUM(PI_Rubrik) sum,Name,Linkname,vermarktung,Website,portfolio.status,rotation.Linkname FROM rotation,rubriken,portfolio WHERE Art='thema' AND portfolio.Port_ID=rubriken.Port_ID AND portfolio.status='online' AND rubriken.Rot_ID=rotation.Rot_ID AND rotation.Name NOT LIKE '%_neu_%' AND rotation.status = '1' GROUP BY Name");


			$counter = 0;
			$r=2;
			$r++;
			$sheetEx = 6;

			while($row = @mysql_fetch_array($result))
			{
				$objPHPExcel->setActiveSheetIndex(5);
				$objPHPExcel->getActiveSheet()->getCell('A'.$r)->setValue($row['Name']);
				//change the data type of the cell
				$objPHPExcel->getActiveSheet()->getCell('A'.$r)->setDataType(PHPExcel_Cell_DataType::TYPE_STRING2);
				//now set the link
				$objPHPExcel->getActiveSheet()->getCell('A'.$r)->getHyperlink()->setUrl(strip_tags('http://www.netpoint-media.de/rotation/'.$row['Linkname'].'.html'));
				// Set url color
				$objPHPExcel->getActiveSheet()->getStyle('A'.$r)->applyFromArray($linkStyle);

				$objPHPExcel->getActiveSheet()->getCell('B'.$r)->setValue(round($row['sum']/1000000,2));

				$counter++;
				$r++;
				//print_r($row['Linkname']);
				//echo $show_imp,EOL,$show_view,EOL,$show_unique,EOL;
				rotation($sheetEx, $row['Linkname'],$show_imp,$show_view,$show_unique);
				$sheetEx++;
			}
			$objPHPExcel->setActiveSheetIndex(5);

			$maxrow = $counter + 2;
			$objPHPExcel->getActiveSheet()->getStyle("A3:B".$maxrow)->applyFromArray($greyCellBackroundStyle);

		}

		// Set active sheet index to the first sheet, so Excel opens this as the first sheet
		$objPHPExcel->setActiveSheetIndex(1);

		//Remove default worksheet
		$objPHPExcel->removeSheetByIndex(0);

		//Hide worksheet clone
		if($objPHPExcel->getSheetByName('clone'))
		{
			$objPHPExcel->getSheetByName('clone')->setSheetState(PHPExcel_Worksheet::SHEETSTATE_VERYHIDDEN);
		}

		// Save Excel 2007 file
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		//echo $strPath,EOL;
		if (file_exists($strPath))
		{
			chmod($strPath, 0777);
			unlink($strPath);
		}
		$objWriter->save(str_replace('.php', '.xlsx', $strPath));

		$status = false;
		//echo $strPath,EOL;
}
function rotation($sheetEx, $rotid,$show_imp = false,$show_view = false,$show_unique = false)
{
	global $clone;
	global $objPHPExcel;
	global $subchannel;
	global $portid;
	global $portcount2;
	global $portcount;
	global $menuStyle;
	global $linkStyle;
	global $greyCellBackroundStyle;
	global $styleUnderLine;
	global $center;

	//Add worksheet netpoint-rotation
	$sheet5 = clone $clone;
	$sheet5->setTitle("Worksheet");
	$objPHPExcel->addSheet($sheet5);
	//echo $sheetEx,EOL;
	$objPHPExcel->setActiveSheetIndex($sheetEx);

	$objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight(36);
	$objPHPExcel->getActiveSheet()->getRowDimension(2)->setRowHeight(25.5);

	$c = 0;
	$r = 2;

	$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
	$objPHPExcel->getActiveSheet()->getColumnDimension($colIndex)->setWidth(36);
	$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue("Portfolio / Website");
	//echo "Portfolio / Website: ".$colIndex,EOL;

	$c++;
	$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
	$objPHPExcel->getActiveSheet()->getColumnDimension($colIndex)->setWidth(25);
	$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue("Platzierung");

	if($show_imp)
	{
		$c++;
		$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
		$objPHPExcel->getActiveSheet()->getColumnDimension($colIndex)->setWidth(16);
		if($rotid==45)
		{
			$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue("Reichweite");
		}else{
			$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue("PageImpressions pro Monat");
			$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getAlignment()->setWrapText(true);
		}
	}
	if($show_view && $rotid!=45)
	{
		$c++;
		$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
		$objPHPExcel->getActiveSheet()->getColumnDimension($colIndex)->setWidth(16);
		$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue("Visits pro Monat");
		$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getAlignment()->setWrapText(true);
	}
	if($show_unique && $rotid!=45)
	{
		$c++;
		$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
		$objPHPExcel->getActiveSheet()->getColumnDimension($colIndex)->setWidth(16);
		$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue("Unique User pro Monat");
		$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getAlignment()->setWrapText(true);
	}
	$c++;
	$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
	$objPHPExcel->getActiveSheet()->getColumnDimension($colIndex)->setWidth(35);
	$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue("Website- & Zielgruppenbeschreibung");
	$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getAlignment()->setWrapText(true);

	$c++;
	$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
	$objPHPExcel->getActiveSheet()->getColumnDimension($colIndex)->setWidth(35);
	$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue("Channel-Beschreibung / Buchungsmöglichkeiten & Preise");
	$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getAlignment()->setWrapText(true);

	$objPHPExcel->getActiveSheet()->getStyle("A2:".$colIndex."2")->applyFromArray($menuStyle);
	$objPHPExcel->getActiveSheet()->getStyle("A2:".$colIndex."2")->applyFromArray($center);

	if($rotid == 'agof_titel'){
		$query = @mysql_query("SELECT visits Visits_Rubrik,uniqueuser uniqueuser_Rubrik,portfolio.Website,portfolio.Port_ID,'Titel-Rotation' Rubrik, PI PI_Rubrik,'agof-titel' RotName,'agof_titel' Linkname,'' Sublevel FROM portfolio WHERE portfolio.status='online' AND portfolio.agof = '1' ORDER BY Website;");
	}
	else {
	  	$query = @mysql_query("SELECT ROUND(portfolio.visits*rubriken.PI_Rubrik/portfolio.PI) Visits_Rubrik,ROUND(portfolio.uniqueuser*rubriken.PI_Rubrik/portfolio.PI) uniqueuser_Rubrik,portfolio.Website,portfolio.Port_ID,rubriken.Rubrik,rubriken.PI_Rubrik,rotation.Name RotName,rotation.Linkname,rotation.Sublevel FROM rotation,rubriken,portfolio WHERE rubriken.Rot_ID=rotation.Rot_ID AND portfolio.Port_ID=rubriken.Port_ID AND portfolio.status='online' AND rotation.Linkname='".$rotid."' ORDER BY sort,Sublevel,Website,Rubrik");
	}
	$c = 0;
	$r = 3;
	$counter = 0;
	while($row = @mysql_fetch_array($query))
	{
		if($rotid == 'agof_titel'){
			$query3 = mysql_query("SELECT visits Visits_Rubrik,uniqueuser uniqueuser_Rubrik,portfolio.Website,portfolio.Port_ID,'Titel-Rotation' Rubrik, PI PI_Rubrik,'agof-titel' RotName,'agof_titel' Linkname,'' Sublevel FROM verticalnetwork,portfolio WHERE slave=Port_ID AND master='".$result[Port_ID]."'");
			echo "agof ".$rotid,EOL;
		}
		//echo $row['Sublevel'],EOL;
		if($subchannel!=$row['Sublevel'])
		{
			$c++;
			$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row['Sublevel']);
		}

		$c = 0;
		if($portid!=$row['Port_ID'])
		{
			$portcount=0;
			$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row['Website']);
		}else{
			$portcount++;
			if($portcount2 && !$portcount)
			{
			}
		}

		$c++;
		$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row['Rubrik']);
		if($show_imp || $rotid==45)
		{
			//$worksheet4->Cells($r,$c++)->value = $result[PI_Rubrik];
			$c++;
			$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row['PI_Rubrik']);
			$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
			$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getNumberFormat()->setFormatCode('#,###');
		}
		if($show_view && $rotid!=45)
		{
			//$worksheet4->Cells($r,$c++)->value = $result[Visits_Rubrik];
			$c++;
			$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row['Visits_Rubrik']);
			$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
			$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getNumberFormat()->setFormatCode('#,###');
		}
		if($show_unique && $rotid!=45)
		{
			//$worksheet4->Cells($r,$c++)->value = $result[uniqueuser_Rubrik];
			$c++;
			$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row['uniqueuser_Rubrik']);
			$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
			$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getNumberFormat()->setFormatCode('#,###');

			/*if(mysql_num_rows($query3)>0){
				$worksheet4->Range($worksheet4->Cells($r,$c-1),$worksheet4->Cells($r+mysql_num_rows($query3),$c-1))->MergeCells = True;
				$worksheet4->Range($worksheet4->Cells($r,$c-1),$worksheet4->Cells($r+mysql_num_rows($query3),$c-1))->VerticalAlignment = 2;
			}*/
		}
		$c++;
		$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row['Website']);
		//change the data type of the cell
		$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setDataType(PHPExcel_Cell_DataType::TYPE_STRING2);
		//now set the link
		$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->getHyperlink()->setUrl(strip_tags('http://www.netpoint-media.de/portfolio/'.$row['Website'].'.html'));

		// Set url color
		$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
		$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->applyFromArray($linkStyle);

		$r++;
		//$c++;
		$portcount2=$portcount;
		$rotname=$row['RotName'];
		$rotname2=$row['RotName'];
		if(strlen($rotname)>31){
			$rotname = substr($rotname,0,30).".";
		}
		$linkname=$row['Linkname'];
		$subchannel=$row['Sublevel'];
		$portid=$row['Port_ID'];
		while($row3 = @mysql_fetch_array($query3))
		{
			print_r($row3);
			if($subchannel!=$row3['Sublevel'])
			{
				$c++;
				$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row3['Sublevel']);
			}

			$c = 0;
			if($portid!=$row3['Port_ID'])
			{
				$portcount=0;
				$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row3['Website']);
			}else{
				$portcount++;
				if($portcount2 && !$portcount)
				{
				}
			}

			$c++;
			$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row3['Rubrik']);
			if($show_imp || $rotid==45)
			{
				//$worksheet4->Cells($r,$c++)->value = $result[PI_Rubrik];
				$c++;
				$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row3['PI_Rubrik']);
				$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
				$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getNumberFormat()->setFormatCode('#,###');
			}
			if($show_view && $rotid!=45)
			{
				//$worksheet4->Cells($r,$c++)->value = $result[Visits_Rubrik];
				$c++;
				$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row3['Visits_Rubrik']);
				$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
				$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getNumberFormat()->setFormatCode('#,###');
			}
			if($show_unique && $rotid!=45)
			{
				//$worksheet4->Cells($r,$c++)->value = $result[uniqueuser_Rubrik];
				$c++;
				$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row3['uniqueuser_Rubrik']);
				$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
				$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->getNumberFormat()->setFormatCode('#,###');
				/*if(mysql_num_rows($query3)>0){
					$worksheet4->Range($worksheet4->Cells($r,$c-1),$worksheet4->Cells($r+mysql_num_rows($query3),$c-1))->MergeCells = True;
					$worksheet4->Range($worksheet4->Cells($r,$c-1),$worksheet4->Cells($r+mysql_num_rows($query3),$c-1))->VerticalAlignment = 2;
				}*/
			}
			$c++;
			$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row3['Website']);
			//change the data type of the cell
			$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setDataType(PHPExcel_Cell_DataType::TYPE_STRING2);
			//now set the link
			$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->getHyperlink()->setUrl(strip_tags('http://www.netpoint-media.de/portfolio/'.$row3['Website'].'.html'));

			// Set url color
			$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
			$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->applyFromArray($linkStyle);

			$r++;
			$portcount2=$portcount;
		}
		$counter++;
	}
	//$objPHPExcel->setActiveSheetIndex($sheetEx);
	$objPHPExcel->getActiveSheet()->setTitle($rotname);

	$r++;
	$c++;

	$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,3)->setValue($rotname2);
	//change the data type of the cell
	$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,3)->setDataType(PHPExcel_Cell_DataType::TYPE_STRING2);
	//now set the link
	$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,3)->getHyperlink()->setUrl(strip_tags('http://www.netpoint-media.de/portfolio/verticals/'.$linkname));
	// Set url color
	$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
	$objPHPExcel->getActiveSheet()->getStyle($colIndex.'3')->applyFromArray($linkStyle);
	$objPHPExcel->getActiveSheet()->getStyle($colIndex.'3')->applyFromArray($center);

	$maxrow = $counter + 2;
	$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
	$objPHPExcel->getActiveSheet()->getStyle("A3:".$colIndex.$maxrow)->applyFromArray($greyCellBackroundStyle);

	$objPHPExcel->getActiveSheet()->mergeCells($colIndex.'3:'.$colIndex.$maxrow);

	/* Summe */
	$c = 1;
	if($show_imp)
	{
		$c++;
		$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
		$objPHPExcel->getActiveSheet()->getCell($colIndex.($maxrow+1))->setValue('=SUM('.$colIndex.'3:'.$colIndex.$maxrow.')');
		$objPHPExcel->getActiveSheet()->getStyle($colIndex.($maxrow+1))->applyFromArray($styleUnderLine);
		$objPHPExcel->getActiveSheet()->getStyle($colIndex.($maxrow+1))->getNumberFormat()->setFormatCode('#,###');
	}
	if($show_view)
	{
		$c++;
		$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
		$objPHPExcel->getActiveSheet()->getCell($colIndex.($maxrow+1))->setValue('=SUM('.$colIndex.'3:'.$colIndex.$maxrow.')');
		$objPHPExcel->getActiveSheet()->getStyle($colIndex.($maxrow+1))->applyFromArray($styleUnderLine);
		$objPHPExcel->getActiveSheet()->getStyle($colIndex.($maxrow+1))->getNumberFormat()->setFormatCode('#,###');
	}
	if($show_unique)
	{
		$c++;
		$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
		$objPHPExcel->getActiveSheet()->getCell($colIndex.($maxrow+1))->setValue('=SUM('.$colIndex.'3:'.$colIndex.$maxrow.')');
		$objPHPExcel->getActiveSheet()->getStyle($colIndex.($maxrow+1))->applyFromArray($styleUnderLine);
		$objPHPExcel->getActiveSheet()->getStyle($colIndex.($maxrow+1))->getNumberFormat()->setFormatCode('#,###');

	}
}
function getSchriftart(){
	global $objPHPExcel;
	$font = $objPHPExcel->getSheetByName('kontaktdaten')->getStyle("B3")->getFont()->getName();
	return $font;
}
function getMenuColor(){
	global $objPHPExcel;
	$color = $objPHPExcel->getSheetByName('kontaktdaten')->getStyle('B3')->getFill()->getStartColor()->getRGB();
	return $color;
}

echo "Impressions, Visits, Unique<br/>";
flush();
createxls(1,1,1,getcwd()."/excel/ivu.xlsx");
flush();
echo "Impressions, Visits<br/>";
createxls(1,1,0,getcwd()."/excel/iv.xlsx");
flush();
echo "Impressions, Unique<br/>";
createxls(1,0,1,getcwd()."/excel/iu.xlsx");
flush();
echo "Visits, Unique<br/>";
createxls(0,1,1,getcwd()."/excel/vu.xlsx");
flush();
echo "Impressions<br/>";
createxls(1,0,0,getcwd()."/excel/i.xlsx");
flush();
echo "Visits<br/>";
createxls(0,1,0,getcwd()."/excel/v.xlsx");
flush();
echo "Unique<br/>";
createxls(0,0,1,getcwd()."/excel/u.xlsx");
echo "fertsch!";

?>
