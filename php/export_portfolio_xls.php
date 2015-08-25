<?php
/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

/** Include PHPExcel */
require_once '../tools/PHPExcel/Classes/PHPExcel.php';
require_once '../tools/PHPExcel/Classes/PHPExcel/IOFactory.php';

/** Include Define **/
require_once("inc_db.php");

$db_selected = mysql_select_db('netpoint_media', $conn);
if (!$db_selected) {
	die ('Kann netpoint_online nicht benutzen : ' . mysql_error());
}

//ToDo
function createxls($show_imp = false,$show_view = false,$show_unique = false,	$strPath)
{
		// Create new PHPExcel object
		//echo date('H:i:s') , " Create new PHPExcel object" , EOL;
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

			$styleArray = array(
		    'font'  => array(
		        'bold'  => false,
		        'color' => array('rgb' => 'FFFFFF')
		    ),
				'fill' => array(
		            'type' => PHPExcel_Style_Fill::FILL_SOLID,
		            'color' => array('rgb' => '000000')
		        ),
				'borders' => array(
				        'allborders' => array(
				            'style' => PHPExcel_Style_Border::BORDER_THIN
					          )
				    )
			);
		  $objPHPExcel->getActiveSheet()->getCell('A2')->setValue("SPEZIFIKATION & ADSERVER");
		  $objPHPExcel->getActiveSheet()->getStyle('A2:D2')->applyFromArray($styleArray);

			$r = 2;
			$r++;
			$objPHPExcel->getActiveSheet()->getCell('A'.$r++)->setValue("URL Werbemittelspezifikationen");
			$objPHPExcel->getActiveSheet()->getCell('A'.$r++)->setValue("E-Mail Banner-Anlieferungsadresse");
			$objPHPExcel->getActiveSheet()->getCell('A'.$r++)->setValue("Adserver Hersteller / Typ / Version");

			$r++;
			$objPHPExcel->getActiveSheet()->getCell('A'.$r)->setValue("SPEZIFIKATIONEN DER WICHTIGSTEN STANDARD-FORMATE");
			$objPHPExcel->getActiveSheet()->getStyle('A7:D7')->applyFromArray($styleArray);

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
			// Config
			$link_style_array = array(
				'font' => array(
						'underline' => 'single',
						'color' => array ('rgb' => '0000FF')
				)
			);
			$objPHPExcel->getActiveSheet()->getStyle('B'.$r)->applyFromArray($link_style_array);

			$r++;
			$objPHPExcel->getActiveSheet()->getCell('B'.$r)->setValue("ADTECH HELIOS IQ");

			$r=8;
			$objPHPExcel->getActiveSheet()->getCell('B'.$r++)->setValue("Max. Dateigewicht in KB");
			$objPHPExcel->getActiveSheet()->getCell('B'.$r++)->setValue("Größe in Pixel");
			$objPHPExcel->getActiveSheet()->getCell('C9')->setValue("GIF, JPG");
			$objPHPExcel->getActiveSheet()->getCell('D9')->setValue("Flash");

			$center = array(
				'font'  => array(
		        'bold'  => true,
		        'color' => array('rgb' => '000000')
		    ),
				'alignment' => array(
						'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
				)
			);
			$objPHPExcel->getActiveSheet()->getStyle("A8:D9")->applyFromArray($center);

			$result = mysql_query("SELECT * FROM werbeformen WHERE werbeformen.online = '1' ORDER BY werbeformen.sort");

			$counter = 0;
			while($row = @mysql_fetch_array($result))
			{
				$objPHPExcel->getActiveSheet()->getCell('A'.$r++)->setValue($row['name']);
				$objPHPExcel->getActiveSheet()->getCell('B'.$r)->setValue($row['format']);
				$objPHPExcel->getActiveSheet()->getCell('C'.($r-1))->setValue($row['gew']);
				$objPHPExcel->getActiveSheet()->getCell('D'.($r-1))->setValue($row['gewflash']);
				$counter++;
			}

			$center2 = array(
				'alignment' => array(
						'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
				)
			);
			$objPHPExcel->getActiveSheet()->getStyle("C10:D56")->applyFromArray($center);

			$styleArray2 = array(
				'fill' => array(
		            'type' => PHPExcel_Style_Fill::FILL_SOLID,
		            'color' => array('rgb' => 'E8E8E8')
		        ),
				'borders' => array(
				        'allborders' => array(
				        		'style' => PHPExcel_Style_Border::BORDER_THIN,
		                'color' => array('rgb' => 'FFFFFF')
					          )
				)
			);
			$maxrow = $counter + 10;
			$objPHPExcel->getActiveSheet()->getStyle("A8:D".$maxrow)->applyFromArray($styleArray2);

		}

		//Add worksheet Portfolio
		$sheet3 = clone $clone;
		$sheet3->setTitle('Portfolio');
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

			$center = array(
				'alignment' => array(
						'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
				)
			);
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
				}
				if($show_view)
				{
					$c++;
					$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row['visits']);
				}
				if($show_unique)
				{
					$c++;
					$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row['uniqueuser']);
				}
				$c++;
				$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setValue($row['Website']);
				//change the data type of the cell
				$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->setDataType(PHPExcel_Cell_DataType::TYPE_STRING2);
				//now set the link
				$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($c,$r)->getHyperlink()->setUrl(strip_tags('http://www.netpoint-media.de/portfolio/'.$row['Website'].'.html'));
				// Set url color
				// Config
				$link_style_array = array(
					'font' => array(
							'underline' => 'single',
							'color' => array ('rgb' => '0000FF')
					)
				);
				$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
				$objPHPExcel->getActiveSheet()->getStyle($colIndex.$r)->applyFromArray($link_style_array);
				$r++;
				$counter++;
			}
			/* Grey Cells Background  */
			$styleArray2 = array(
				'fill' => array(
		            'type' => PHPExcel_Style_Fill::FILL_SOLID,
		            'color' => array('rgb' => 'E8E8E8')
		        ),
				'borders' => array(
				        'allborders' => array(
				        		'style' => PHPExcel_Style_Border::BORDER_THIN,
		                'color' => array('rgb' => 'FFFFFF')
					          )
				)
			);
			$maxrow = $counter + 2;
			$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
			$objPHPExcel->getActiveSheet()->getStyle("A3:".$colIndex.$maxrow)->applyFromArray($styleArray2);

			/* Schwarze Menu  */
			$styleArray = array(
		    'font'  => array(
		        'bold'  => false,
		        'color' => array('rgb' => 'FFFFFF')
		    ),
				'fill' => array(
		            'type' => PHPExcel_Style_Fill::FILL_SOLID,
		            'color' => array('rgb' => '000000')
		        ),
				'borders' => array(
				        'allborders' => array(
				            'style' => PHPExcel_Style_Border::BORDER_THIN,
										'color' => array('rgb' => 'FFFFFF')
					          )
				    )
			);
		 	$objPHPExcel->getActiveSheet()->getStyle('A2:'.$colIndex.'2')->applyFromArray($styleArray);

			$styleUnderLine = array(
				'font' => array(
					'underline' => PHPExcel_Style_Font::UNDERLINE_DOUBLE
				)
			);

			/* Summe */
			$c = 0;
			if($show_imp)
			{
				$c++;
				$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
				$objPHPExcel->getActiveSheet()->getCell($colIndex.($maxrow+1))->setValue('=SUM(B3:B'.$maxrow.')');
				$objPHPExcel->getActiveSheet()->getStyle($colIndex.($maxrow+1))->applyFromArray($styleUnderLine);
			}
			if($show_view)
			{
				$c++;
				$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
				$objPHPExcel->getActiveSheet()->getCell($colIndex.($maxrow+1))->setValue('=SUM(C3:B'.$maxrow.')');
				$objPHPExcel->getActiveSheet()->getStyle($colIndex.($maxrow+1))->applyFromArray($styleUnderLine);
			}
			if($show_unique)
			{
				$c++;
				$colIndex = PHPExcel_Cell::stringFromColumnIndex($c);
				$objPHPExcel->getActiveSheet()->getCell($colIndex.($maxrow+1))->setValue('=SUM(D3:B'.$maxrow.')');
				$objPHPExcel->getActiveSheet()->getStyle($colIndex.($maxrow+1))->applyFromArray($styleUnderLine);
			}
		}

		//Add worksheet netpoint-rotation
		$sheet5 = clone $clone;
		$sheet5->setTitle('Channel');
		$objPHPExcel->addSheet($sheet5);

		//Arbeitsblatt Channel
		$objPHPExcel->setActiveSheetIndex(5);
		{
			$objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight(36);
			$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(36);
		  $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(16);
		  $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(9);
		  $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(22);

			$styleArray = array(
		    'font'  => array(
		        'bold'  => false,
		        'color' => array('rgb' => 'FFFFFF')
		    ),
				'fill' => array(
		            'type' => PHPExcel_Style_Fill::FILL_SOLID,
		            'color' => array('rgb' => '000000')
		        ),
				'borders' => array(
				        'allborders' => array(
				            'style' => PHPExcel_Style_Border::BORDER_THIN
					          )
				    )
			);
		  $objPHPExcel->getActiveSheet()->getCell('A2')->setValue('Themen-Rotation');
			$objPHPExcel->getActiveSheet()->getCell('B2')->setValue('PI in Mio. / Monat');
		  $objPHPExcel->getActiveSheet()->getStyle('A2:B2')->applyFromArray($styleArray);

			$result = mysql_query("SELECT SUM(PI_Rubrik) sum,Name,Linkname,vermarktung,Website,portfolio.status,rotation.Linkname FROM rotation,rubriken,portfolio WHERE Art='thema' AND portfolio.Port_ID=rubriken.Port_ID AND portfolio.status='online' AND rubriken.Rot_ID=rotation.Rot_ID AND rotation.Name NOT LIKE '%_neu_%' AND rotation.status = '1' GROUP BY Name");


			$counter = 0;
			$r=2;
			$r++;
			while($row = @mysql_fetch_array($result))
			{
				$objPHPExcel->getActiveSheet()->getCell('A'.$r)->setValue($row['Name']);
				//change the data type of the cell
				$objPHPExcel->getActiveSheet()->getCell('A'.$r)->setDataType(PHPExcel_Cell_DataType::TYPE_STRING2);
				//now set the link
				$objPHPExcel->getActiveSheet()->getCell('A'.$r)->getHyperlink()->setUrl(strip_tags('http://www.netpoint-media.de/rotation/'.$row['Linkname'].'.html'));
				// Set url color
				// Config
				$link_style_array = array(
					'font' => array(
							'underline' => 'single',
							'color' => array ('rgb' => '0000FF')
					)
				);
				$objPHPExcel->getActiveSheet()->getStyle('A'.$r)->applyFromArray($link_style_array);

				$objPHPExcel->getActiveSheet()->getCell('B'.$r)->setValue(round($row['sum']/1000000,2));

				$counter++;
				$r++;
				//print_r($row['Linkname']);
				//echo $show_imp,EOL,$show_view,EOL,$show_unique,EOL;
				//rotation($row['Linkname'],$show_imp,$show_view,$show_unique);
			}
			$styleArray2 = array(
				'fill' => array(
		            'type' => PHPExcel_Style_Fill::FILL_SOLID,
		            'color' => array('rgb' => 'E8E8E8')
		        ),
				'borders' => array(
				        'allborders' => array(
				        		'style' => PHPExcel_Style_Border::BORDER_THIN,
		                'color' => array('rgb' => 'FFFFFF')
					          )
				)
			);
			$maxrow = $counter + 2;
			$objPHPExcel->getActiveSheet()->getStyle("A3:B".$maxrow)->applyFromArray($styleArray2);

		}

		//Add worksheet netpoint-rotation
		$sheet6 = clone $clone;
		$sheet6->setTitle('netpoint-rotation');
		$objPHPExcel->addSheet($sheet6);

		//Arbeitsblatt netpoint-rotation
		$objPHPExcel->setActiveSheetIndex(6);
		{
			$objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight(36);
			$objPHPExcel->getActiveSheet()->getRowDimension(2)->setRowHeight(25.5);

			$styleArray = array(
		    'font'  => array(
		        'bold'  => false,
		        'color' => array('rgb' => 'FFFFFF')
		    ),
				'fill' => array(
		            'type' => PHPExcel_Style_Fill::FILL_SOLID,
		            'color' => array('rgb' => '000000')
		        ),
				'borders' => array(
				        'allborders' => array(
				            'style' => PHPExcel_Style_Border::BORDER_THIN,
										'color' => array('rgb' => 'FFFFFF')
					          )
				    )
			);
		  //$objPHPExcel->getActiveSheet()->getCell('A2')->setValue("SPEZIFIKATION & ADSERVER");
		  $objPHPExcel->getActiveSheet()->getStyle('A2:G2')->applyFromArray($styleArray);

			$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(36);
			$objPHPExcel->getActiveSheet()->getCell('A2')->setValue("Portfolio / Website");
		  $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(25);
			$objPHPExcel->getActiveSheet()->getCell('B2')->setValue("Platzierung");
			$objPHPExcel->getActiveSheet()->getStyle('B2')->getAlignment()->setWrapText(true);
			$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(16);
			$objPHPExcel->getActiveSheet()->getCell('C2')->setValue("PageImpressions pro Monat");
			$objPHPExcel->getActiveSheet()->getStyle('C2')->getAlignment()->setWrapText(true);
		  $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(16);
			$objPHPExcel->getActiveSheet()->getCell('D2')->setValue("Visits pro Monat");
			$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(16);
			$objPHPExcel->getActiveSheet()->getCell('E2')->setValue("Unique User pro Monat");
			$objPHPExcel->getActiveSheet()->getStyle('E2')->getAlignment()->setWrapText(true);
			$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(35);
			$objPHPExcel->getActiveSheet()->getCell('F2')->setValue("Website- & Zielgruppenbeschreibung");
			$objPHPExcel->getActiveSheet()->getStyle('F2')->getAlignment()->setWrapText(true);
		  $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(35);
			$objPHPExcel->getActiveSheet()->getCell('G2')->setValue("Channel-Beschreibung / Buchungsmöglichkeiten & Preise");
			$objPHPExcel->getActiveSheet()->getStyle('G2')->getAlignment()->setWrapText(true);

			$center = array(
				'alignment' => array(
						'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
				)
			);
			$objPHPExcel->getActiveSheet()->getStyle("A2:G2")->applyFromArray($center);

		}
		// Set active sheet index to the first sheet, so Excel opens this as the first sheet
		$objPHPExcel->setActiveSheetIndex(1);

		//Remove default worksheet
		$objPHPExcel->removeSheetByIndex(0);
		//Hide worksheet clone
		$objPHPExcel->getSheetByName('clone')->setSheetState(PHPExcel_Worksheet::SHEETSTATE_VERYHIDDEN);

		// Save Excel 2007 file
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		//echo $strPath,EOL;
		if (file_exists($strPath))
		{
			chmod($strPath, 0777);
			unlink($strPath);
		}
		$objWriter->save(str_replace('.php', '.xlsx', $strPath));

		echo $strPath,EOL;
}
function rotation($rotid,$show_imp = false,$show_view = false,$show_unique = false)
{
	/*$i = 5;
	$j = 4;
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
	//echo $rotid.'_'.$show_imp.'_'.$show_view.'_'.$show_unique,EOL;
	$objPHPExcel->setActiveSheetIndex(2);
	//Clone worksheet index 2
	$sheet4 = $objPHPExcel->getActiveSheet()->copy();
	//Add worksheet Portfolio
	$sheet5 = clone $sheet4;
	//$sheet2->setTitle($rotid);
	$objPHPExcel->addSheet($sheet4);
	$i++;*/
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
