<?
ini_set("max_execution_time","0");
error_reporting('e_all');


function createxls($show_imp = false,$show_view = false,$show_unique = false,	$strPath)
{
	global $wkb;
	//Open Excel, Hit Alt+F11, then Hit F2 -- this is your COM bible

// Start
{
	//starting excel
	$excel = new COM("excel.application") or die("Unable to instanciate excel");
//	echo "Loaded excel, version {$excel->Version}\n";

	//bring it to front
	 $excel->Visible = 1;//NOT

	//dont want alerts ... run silent
	$excel->DisplayAlerts = 1;
}
	//Arbeitsmappe erstellen
	$wkb = $excel->Workbooks->Add();

		$wkb2 = $excel->Workbooks->Open("C:/apachefriends/xampp/htdocs/g-tool/excel/deckblatt.xls");
		$sheet=$wkb->Worksheets(1);
		$sheet2=$wkb2->Worksheets(1);
		$sheet2->Copy($sheet);
		$wkb2->Close(false);
		$worksheet=$wkb->Worksheets(1);
		$worksheet->Name = "kontaktdaten";
		$wsnr = 1;
$wsnr++;
// Arbeitsblatt Technik
{
	$wkb->Worksheets($wkb->Worksheets->Count())->activate;
	$wkb2 = $excel->Workbooks->Open("C:/apachefriends/xampp/htdocs/g-tool/excel/deckblatt.xls");
	$sheet=$wkb->Worksheets($wsnr);
	$sheet2=$wkb2->Worksheets(2);
	$sheet2->Copy($sheet);
	$wkb2->Close(false);
	$worksheet3=$wkb->Worksheets($wsnr);
	$worksheet3->Name = "technik";
	// Hintergrund wei� + breiten + h�hen
	{
		$worksheet3->Cells->Interior->ColorIndex = 2;
		$worksheet3->Columns("A:A")->ColumnWidth = 36;
		$worksheet3->Columns("B:B")->ColumnWidth = 58;
		$worksheet3->Columns("C:C")->ColumnWidth = 11;
		$worksheet3->Columns("D:D")->ColumnWidth = 11;
		$worksheet3->Rows("1:1")->RowHeight = 36;
	}
	// Spezifikationen & Adserver
	{
		$worksheet3->Range("A2:D2")->MergeCells = True;
		$worksheet3->Range("B3:D3")->MergeCells = True;
		$worksheet3->Range("B4:D4")->MergeCells = True;
		$worksheet3->Range("B5:D5")->MergeCells = True;
		$worksheet3->Range("A7:D7")->MergeCells = True;
		$worksheet3->Range("B8:D8")->MergeCells = True;
		$worksheet3->Range("A8:A9")->MergeCells = True;

		$c=1;
		$r=2;
		$worksheet3->Cells($r,$c)->value = "Spezifikationen & Adserver";
		$worksheet3->Cells($r,$c)->Font->Bold = True;
		$worksheet3->Cells($r,$c)->Interior->ColorIndex = 46;
		$worksheet3->Cells($r,$c)->Font->ColorIndex = 2;
		$worksheet3->Cells($r,$c)->HorizontalAlignment = 3;
		$r++;
		$worksheet3->Cells($r++,$c)->value = "URL Werbemittelspezifikationen";
		$worksheet3->Cells($r++,$c)->value = "E-Mail Banner-Anlieferungsadresse";
		$worksheet3->Cells($r++,$c)->value = "Adserver Hersteller / Typ / Version";
		$r++;
		$worksheet3->Cells($r,$c)->value = "Spezifikationen der wichtigsten Standard-Formate";
		$worksheet3->Cells($r,$c)->Font->Bold = True;
		$worksheet3->Cells($r,$c)->Interior->ColorIndex = 46;
		$worksheet3->Cells($r,$c)->Font->ColorIndex = 2;
		$worksheet3->Cells($r,$c)->HorizontalAlignment = 3;
		$r++;
		$worksheet3->Cells($r++,$c)->value = "Format";
		$c=2;
		$r=3;
		$worksheet3->Cells($r,$c)->value = "http://www.netpoint-media.de/werbeformen/spezifikationen.html";
		$worksheet3->Cells($r,$c)->Hyperlinks->Add($worksheet3->Cells($r,$c) ,'http://www.netpoint-media.de/werbeformen/spezifikationen.html');
		$r++;
		$worksheet3->Cells($r,$c)->value = "banner@netpoint-media.de";
		$worksheet3->Cells($r,$c)->Hyperlinks->Add($worksheet3->Cells($r,$c) ,'mailto:banner@netpoint-media.de');
		$r++;
		$worksheet3->Cells($r++,$c)->value = "ADTECH HELIOS IQ";

		$worksheet3->Range("A8:D9")->Font->Bold = True;
		$worksheet3->Range("A8:D9")->HorizontalAlignment = 3;
		$worksheet3->Range("A8:D9")->VerticalAlignment = 1;

		$r=8;
		$worksheet3->Cells($r++,$c)->value = "Max. Dateigewicht in KB";
		$worksheet3->Cells($r,$c++)->value = "Gr��e in Pixel";
		$worksheet3->Cells($r,$c++)->value = "GIF, JPG";
		$worksheet3->Cells($r++,$c)->value = "Flash";

	$dbHost = "db.netpoint-media.de";
	$dbUser = "tm";
	$dbPass = "dbpasswort";
	$dbName = "dbtm_netpoint";
		$connect = @mysql_connect($dbHost, $dbUser, $dbPass) or die(mysql_error());
		$selectDB = @mysql_select_db($dbName, $connect);

		$r=10;
		$c=1;
		$query2 = @mysql_query("SELECT * FROM werbeformen WHERE werbeformen.online = '1' ORDER BY werbeformen.sort;");

		while($result2 = @mysql_fetch_array($query2))
		{
			$worksheet3->Cells($r,$c++)->value = $result2[name];
			$worksheet3->Cells($r,$c++)->value = $result2[format];
			$worksheet3->Cells($r,$c++)->value = $result2[gew];
			$worksheet3->Cells($r++,$c)->value = $result2[gewflash];
			$c=1;
		}
		$bisborder=$r-1;
		$worksheet3->Range("C10:D".$bisborder)->HorizontalAlignment = 3;
		for($i=1;$i<5;$i++)
		{
			$worksheet3->Range("A7:D".$bisborder)->Borders($i)->LineStyle = 1;
			$worksheet3->Range("A7:D".$bisborder)->Borders($i)->Weight = 2;
			$worksheet3->Range("A7:D".$bisborder)->Borders($i)->ColorIndex = 1;
			$worksheet3->Range("A2:D5")->Borders($i)->LineStyle = 1;
			$worksheet3->Range("A2:D5")->Borders($i)->Weight = 2;
			$worksheet3->Range("A2:D5")->Borders($i)->ColorIndex = 1;
		}
	}
}
$wsnr++;
// Arbeitsblatt Portfolio
{
	$wkb->Worksheets($wkb->Worksheets->Count())->activate;
	$wkb2 = $excel->Workbooks->Open("C:/apachefriends/xampp/htdocs/g-tool/excel/deckblatt.xls");
	$sheet=$wkb->Worksheets($wsnr);
	$sheet2=$wkb2->Worksheets(2);
	$sheet2->Copy($sheet);
	$wkb2->Close(false);
	$worksheet4=$wkb->Worksheets($wsnr);
	$worksheet4->Name = "Portfolio";
	// Hintergrund wei� + breiten + h�hen
	{
		$worksheet4->Cells->Interior->ColorIndex = 2;
		$worksheet4->Rows("1:1")->RowHeight = 36;
		$worksheet4->Rows("2:2")->RowHeight = 25.5;
	}
	$r=2;
	$c=1;

	$worksheet4->Columns(chr($c+64).":".chr($c+64))->ColumnWidth = 36;
	$worksheet4->Cells($r,$c++)->value = 'Site Name';
	if($show_imp)
	{
		$worksheet4->Columns(chr($c+64).":".chr($c+64))->ColumnWidth = 16;
		$worksheet4->Cells($r,$c++)->value = 'PageImpressions pro Monat';
	}
	if($show_view)
	{
		$worksheet4->Columns(chr($c+64).":".chr($c+64))->ColumnWidth = 16;
		$worksheet4->Cells($r,$c++)->value = 'Visits               pro Monat';
	}
	if($show_unique)
	{
		$worksheet4->Columns(chr($c+64).":".chr($c+64))->ColumnWidth = 16;
		$worksheet4->Cells($r,$c++)->value = 'Unique User               pro Monat';
	}
	$worksheet4->Columns(chr($c+64).":".chr($c+64))->ColumnWidth = 40;
	$worksheet4->Cells($r++,$c++)->value = 'Website- & Zielgruppenbeschreibung / Buchungsm�glichkeiten & Preise';

	$c--;
	$worksheet4->Range("A2:".chr($c+64)."2")->Font->Bold = True;
	$worksheet4->Range("A2:".chr($c+64)."2")->Interior->ColorIndex = 46;
	$worksheet4->Range("A2:".chr($c+64)."2")->Font->ColorIndex = 2;
	$worksheet4->Range("A2:".chr($c+64)."2")->HorizontalAlignment = 3;
	$worksheet4->Range("A2:".chr($c+64)."2")->WrapText = True;
	$worksheet4->Range("A2:".chr($c+64)."2")->VerticalAlignment = 1;

	$query = @mysql_query("SELECT * FROM portfolio WHERE portfolio.status = 'online' ORDER BY Website");
	while($result = @mysql_fetch_array($query))
	{
		$c=1;
		$worksheet4->Cells($r,$c++)->value = $result[Website];
		if($show_imp)
		{
			$worksheet4->Cells($r,$c++)->value = $result[PI];
		}
		if($show_view)
		{
			$worksheet4->Cells($r,$c++)->value = $result[visits];
		}
		if($show_unique)
		{
			$worksheet4->Cells($r,$c++)->value = $result[uniqueuser];
		}
		$worksheet4->Cells($r,$c)->value = $result[Website];
		$worksheet4->Cells($r,$c)->Hyperlinks->Add($worksheet4->Cells($r,$c) ,'http://www.netpoint-media.de/portfolio/'.$result[Website].'.html');
		$r++;
		$c++;
	}
	$r--;
	$c--;
		$worksheet4->Range("A2:".chr($c+64)."".$r)->NumberFormat = '#.###';
		for($i=1;$i<5;$i++)
		{
			$worksheet4->Range("A2:".chr($c+64)."".$r)->Borders($i)->LineStyle = 1;
			$worksheet4->Range("A2:".chr($c+64)."".$r)->Borders($i)->Weight = 2;
			$worksheet4->Range("A2:".chr($c+64)."".$r)->Borders($i)->ColorIndex = 1;
		}
		$c=2;
		$r++;
		if($show_imp)
		{
			$worksheet4->Cells($r,$c)->value = "=SUMME(".chr($c+64)."3:".chr($c+64)."".($r-1);
			$c++;
		}
		if($show_view)
		{
			$worksheet4->Cells($r,$c)->value = "=SUMME(".chr($c+64)."3:".chr($c+64)."".($r-1);
			$c++;
		}
		if($show_unique)
		{
			$worksheet4->Cells($r,$c)->value = "=SUMME(".chr($c+64)."3:".chr($c+64)."".($r-1);
			$c++;
		}
		$worksheet4->Range("B".$r.":".chr($c+64)."".$r)->Font->Bold = True;
		$worksheet4->Range("B".$r.":".chr($c+64)."".$r)->Font->Underline = 5;
}
$wsnr++;
// Arbeitsblatt Channel
{
	$wkb->Worksheets($wkb->Worksheets->Count())->activate;
	$wkb2 = $excel->Workbooks->Open("C:/apachefriends/xampp/htdocs/g-tool/excel/deckblatt.xls");
	$sheet=$wkb->Worksheets($wsnr);
	$sheet2=$wkb2->Worksheets(2);
	$sheet2->Copy($sheet);
	$wkb2->Close(false);
	$wkb->Worksheets($wkb->Worksheets->Count())->Delete();
	$wkb->Worksheets($wkb->Worksheets->Count())->Delete();
	$worksheet5=$wkb->Worksheets($wsnr);
	$worksheet5->Name = "Channel";
	// Hintergrund wei� + breiten + h�hen
	{
		$worksheet5->Cells->Interior->ColorIndex = 2;
		$worksheet5->Columns("A:A")->ColumnWidth = 36;
		$worksheet5->Columns("B:B")->ColumnWidth = 16;
		$worksheet5->Columns("C:C")->ColumnWidth = 9;
		$worksheet5->Columns("D:D")->ColumnWidth = 22;
		$worksheet5->Columns("E:E")->ColumnWidth = 16;
		$worksheet5->Columns("F:F")->ColumnWidth = 5;
		$worksheet5->Columns("G:G")->ColumnWidth = 36;
		$worksheet5->Rows("1:1")->RowHeight = 36;
	}
	$r=2;
	$c=1;

	$worksheet5->Cells($r,$c++)->value = 'Themen-Rotation';
	$worksheet5->Cells($r++,$c++)->value = 'PI in Mio. / Monat';
	$worksheet5->Range("A2:B2")->Font->Bold = True;
	$worksheet5->Range("A2:B2")->Interior->ColorIndex = 46;
	$worksheet5->Range("A2:B2")->Font->ColorIndex = 2;
	$worksheet5->Range("A2:B2")->HorizontalAlignment = 1;
	$worksheet5->Range("A2:B2")->VerticalAlignment = 1;

//  $query = @mysql_query("SELECT SUM(PI_Rubrik) sum,Name,Linkname,vermarktung,Website,status,rotation.Linkname FROM rotation,rubriken,portfolio WHERE Art='thema' AND portfolio.Port_ID=rubriken.Port_ID AND portfolio.status='online' AND rubriken.Rot_ID=rotation.Rot_ID AND (rotation.Rot_ID < 45 OR rotation.Linkname = 'musik_mp3_popkultur') AND rotation.Rot_ID != 7 AND rotation.Rot_ID != 61 GROUP BY Name");    //rotation.Rot_ID NOT IN (7,45,61,56,58) OR rotation.Rot_ID IN (61,56)
  $query = @mysql_query("SELECT SUM(PI_Rubrik) sum,Name,Linkname,vermarktung,Website,portfolio.status,rotation.Linkname FROM rotation,rubriken,portfolio WHERE Art='thema' AND portfolio.Port_ID=rubriken.Port_ID AND portfolio.status='online' AND rubriken.Rot_ID=rotation.Rot_ID AND rotation.Name NOT LIKE '%_neu_%' AND rotation.status = '1' GROUP BY Name");

	while($result = @mysql_fetch_array($query))
	{
		$c=1;
		$worksheet5->Cells($r,$c)->value = $result[Name];
		$worksheet5->Cells($r,$c)->Hyperlinks->Add($worksheet5->Cells($r,$c) ,'http://www.netpoint-media.de/rotation/'.$result[Linkname].'.html');
		$c++;
		$worksheet5->Cells($r,$c)->value = round($result[sum]/1000000,2);
		$r++;
		$wkb2 = $excel->Workbooks->Open("C:/apachefriends/xampp/htdocs/g-tool/excel/deckblatt.xls");
		$sheet=$wkb->Worksheets(($wkb->Worksheets->Count()));
		$sheet2=$wkb2->Worksheets(2);
		$sheet2->Copy($sheet);
		$wkb2->Close(false);

		rotation($result[Linkname],$show_imp,$show_view,$show_unique);
	}
	$r--;
		$worksheet5->Range("A2:B".$r)->NumberFormat = '#.###0,00';
		for($i=1;$i<5;$i++)
		{
			$worksheet5->Range("A2:B".$r)->Borders($i)->LineStyle = 1;
			$worksheet5->Range("A2:B".$r)->Borders($i)->Weight = 2;
			$worksheet5->Range("A2:B".$r)->Borders($i)->ColorIndex = 1;
		}

	$r=2;
	$c=4;

/*
	$worksheet5->Cells($r,$c++)->value = 'Regional-Rotation';
	$worksheet5->Cells($r++,$c++)->value = 'PI in Mio. / Monat';
	$worksheet5->Range("D2:E2")->Font->Bold = True;
	$worksheet5->Range("D2:E2")->Interior->ColorIndex = 46;
	$worksheet5->Range("D2:E2")->Font->ColorIndex = 2;
	$worksheet5->Range("D2:E2")->HorizontalAlignment = 1;
	$worksheet5->Range("D2:E2")->VerticalAlignment = 1;

  $query = @mysql_query("SELECT SUM(PI_Rubrik) sum,Name,Linkname,vermarktung,Website,status,rotation.Linkname FROM rotation,rubriken,portfolio WHERE Art='regional' AND portfolio.Port_ID=rubriken.Port_ID AND portfolio.status='online' AND rubriken.Rot_ID=rotation.Rot_ID GROUP BY Name");
	while($result = @mysql_fetch_array($query))
	{
		$c=4;
		$worksheet5->Cells($r,$c)->value = $result[Name];
		$worksheet5->Cells($r,$c)->Hyperlinks->Add($worksheet5->Cells($r,$c) ,'http://www.netpoint-media.de/rotation/'.$result[Linkname].'.html');
		$c++;
		$worksheet5->Cells($r,$c)->value = round($result[sum]/1000000,2);
		$r++;
		rotation($result[Linkname],$show_imp,$show_view,$show_unique);
	}
	$r--;
		$worksheet5->Range("D2:E".$r)->NumberFormat = '#.###0,00';
		for($i=1;$i<5;$i++)
		{
			$worksheet5->Range("D2:E".$r)->Borders($i)->LineStyle = 1;
			$worksheet5->Range("D2:E".$r)->Borders($i)->Weight = 2;
			$worksheet5->Range("D2:E".$r)->Borders($i)->ColorIndex = 1;
		}

		$worksheet5->Range("G2:G9")->MergeCells = True;
		$worksheet5->Range("G2:G9")->VerticalAlignment = 1;
		$worksheet5->Range("G2:G9")->WrapText = True;
		$worksheet5->Cells(2,7)->value = 'Optional kann die regionale Werbung durch eine Netzwerkschaltung mit IP-Targeting verst�rkt werden und �ber ADTECH "AdLocal" nach ACNielsen Gebieten- und Ballungsr�umen, Postleitzahlengebieten, Stadt- und Landkreisen, bis hin auf Gemeinde-Ebenen punktgenau ausgeliefert werden.';
		for($i=1;$i<5;$i++)
		{
			$worksheet5->Range("G2:G9")->Borders($i)->LineStyle = 1;
			$worksheet5->Range("G2:G9")->Borders($i)->Weight = 3;
			$worksheet5->Range("G2:G9")->Borders($i)->ColorIndex = 1;
		}
		*/
}


// Logo einf�gen
/*
{
	for($i=2;$i<$wkb->Worksheets->Count();$i++)
	{
		$wkb->Worksheets($i)->Shapes->AddPicture("C:/netpoint.gif",0,1,13, 6, 138, 65);
	}

}
*/
$wkb->Worksheets(1)->activate;
// Abschluss
{
	$wkb->Worksheets($wkb->Worksheets->Count())->Delete();
	// XLS abspeichern
	if (file_exists($strPath)) {unlink($strPath);}
	$wkb->SaveAs($strPath,51);

	// schliessen
	$wkb->Close(false);
	$excel->Workbooks->Close();
	unset($worksheet);
	unset($worksheet2);
	unset($worksheet3);
	unset($worksheet4);
	unset($worksheet5);
	$excel->Quit();
	$excel = null;
}
}
// Arbeitsbl�tter Rotationen
function rotation($rotid,$show_imp = false,$show_view = false,$show_unique = false)
{
	global $wkb;
	$wkb->Worksheets($wkb->Worksheets->Count())->activate;
//	$wkb->Worksheets->Add();
	$worksheet4=$wkb->Worksheets(($wkb->Worksheets->Count())-1);
	$worksheet4->Name = time();
	// Hintergrund wei� + breiten + h�hen
	{
	$worksheet4->Cells->Interior->ColorIndex = 2;
		$worksheet4->Rows("1:1")->RowHeight = 36;
		$worksheet4->Rows("2:2")->RowHeight = 25.5;
	}
	$r=2;
	$c=1;

	$worksheet4->Columns(chr($c+64).":".chr($c+64))->ColumnWidth = 36;
	$worksheet4->Cells($r,$c++)->value = 'Portfolio / Website';
	$worksheet4->Columns(chr($c+64).":".chr($c+64))->ColumnWidth = 25;
	$worksheet4->Cells($r,$c++)->value = 'Platzierung';
	if($show_imp)
	{
		$worksheet4->Columns(chr($c+64).":".chr($c+64))->ColumnWidth = 16;
		if($rotid==45)
			$worksheet4->Cells($r,$c++)->value = 'Reichweite';
		else
			$worksheet4->Cells($r,$c++)->value = 'PageImpressions pro Monat';
	}
	if($show_view && $rotid!=45)
	{
		$worksheet4->Columns(chr($c+64).":".chr($c+64))->ColumnWidth = 16;
		$worksheet4->Cells($r,$c++)->value = 'Visits               pro Monat';
	}
	if($show_unique && $rotid!=45)
	{
		$worksheet4->Columns(chr($c+64).":".chr($c+64))->ColumnWidth = 16;
		$worksheet4->Cells($r,$c++)->value = 'Unique User               pro Monat';
	}
	$worksheet4->Columns(chr($c+64).":".chr($c+64))->ColumnWidth = 35;
	$worksheet4->Cells($r,$c++)->value = 'Website- & Zielgruppenbeschreibung';
	$worksheet4->Columns(chr($c+64).":".chr($c+64))->ColumnWidth = 35;
	$worksheet4->Cells($r++,$c)->value = 'Channel-Beschreibung / Buchungsm�glichkeiten & Preise';


	$worksheet4->Range("A2:".chr($c+64)."2")->Font->Bold = True;
	$worksheet4->Range("A2:".chr($c+64)."2")->Interior->ColorIndex = 46;
	$worksheet4->Range("A2:".chr($c+64)."2")->Font->ColorIndex = 2;
	$worksheet4->Range("A2:".chr($c+64)."2")->HorizontalAlignment = 3;
	$worksheet4->Range("A2:".chr($c+64)."2")->WrapText = True;
	$worksheet4->Range("A2:".chr($c+64)."2")->VerticalAlignment = 1;

	if($rotid == 'agof_titel'){
		$query = @mysql_query("SELECT visits Visits_Rubrik,uniqueuser uniqueuser_Rubrik,portfolio.Website,portfolio.Port_ID,'Titel-Rotation' Rubrik, PI PI_Rubrik,'agof-titel' RotName,'agof_titel' Linkname,'' Sublevel FROM portfolio WHERE portfolio.status='online' AND portfolio.agof = '1' ORDER BY Website;");
	}
	else {
	  $query = @mysql_query("SELECT ROUND(portfolio.visits*rubriken.PI_Rubrik/portfolio.PI) Visits_Rubrik,ROUND(portfolio.uniqueuser*rubriken.PI_Rubrik/portfolio.PI) uniqueuser_Rubrik,portfolio.Website,portfolio.Port_ID,rubriken.Rubrik,rubriken.PI_Rubrik,rotation.Name RotName,rotation.Linkname,rotation.Sublevel FROM rotation,rubriken,portfolio WHERE rubriken.Rot_ID=rotation.Rot_ID AND portfolio.Port_ID=rubriken.Port_ID AND portfolio.status='online' AND rotation.Linkname='".$rotid."' ORDER BY sort,Sublevel,Website,Rubrik");
	}
	while($result = @mysql_fetch_array($query))
	{
		if($rotid == 'agof_titel'){
			$query3 = mysql_query("SELECT visits Visits_Rubrik,uniqueuser uniqueuser_Rubrik,portfolio.Website,portfolio.Port_ID,'Titel-Rotation' Rubrik, PI PI_Rubrik,'agof-titel' RotName,'agof_titel' Linkname,'' Sublevel FROM verticalnetwork,portfolio WHERE slave=Port_ID AND master='".$result[Port_ID]."'");
		}
		if($subchannel!=$result[Sublevel])
		{
			$worksheet4->Range("A".$r.":".chr($c+63)."".$r)->MergeCells = True;
			$worksheet4->Range("A".$r.":".chr($c+63)."".$r)->Font->Bold = True;
			$worksheet4->Cells($r++,1)->value = $result[Sublevel];
		}
		$c=1;
		if($portid!=$result[Port_ID])
		{
			$portcount=0;
			$worksheet4->Cells($r,$c)->value = $result[Website];
		}
		else
			$portcount++;
		if($portcount2 && !$portcount)
		{
			$worksheet4->Range("A".($r-$portcount2-1).":A".($r-1))->MergeCells = True;
			$worksheet4->Range("A".($r-$portcount2-1).":A".($r-1))->VerticalAlignment = 1;
		}
		$c++;
		$worksheet4->Cells($r,$c++)->value = $result[Rubrik];
		if($show_imp || $rotid==45)
			$worksheet4->Cells($r,$c++)->value = $result[PI_Rubrik];
		if($show_view && $rotid!=45)
			$worksheet4->Cells($r,$c++)->value = $result[Visits_Rubrik];
		if($show_unique && $rotid!=45){
			$worksheet4->Cells($r,$c++)->value = $result[uniqueuser_Rubrik];
			if(mysql_num_rows($query3)>0){
			$worksheet4->Range($worksheet4->Cells($r,$c-1),$worksheet4->Cells($r+mysql_num_rows($query3),$c-1))->MergeCells = True;
			$worksheet4->Range($worksheet4->Cells($r,$c-1),$worksheet4->Cells($r+mysql_num_rows($query3),$c-1))->VerticalAlignment = 2;
			}
		}
		$worksheet4->Cells($r,$c)->value = $result[Website];
		$worksheet4->Cells($r,$c)->Hyperlinks->Add($worksheet4->Cells($r,$c) ,'http://www.netpoint-media.de/portfolio/'.$result[Website].'.html');
		$r++;
		$c++;
		$portcount2=$portcount;
		$rotname=$result[RotName];
		$rotname2=$result[RotName];
		if(strlen($rotname)>31)
			$rotname = substr($rotname,0,30).".";
		$linkname=$result[Linkname];
		$subchannel=$result[Sublevel];
		$portid=$result[Port_ID];
		while($result3 = @mysql_fetch_array($query3))
		{
			if($subchannel!=$result3[Sublevel])
			{
				$worksheet4->Range("A".$r.":".chr($c+63)."".$r)->MergeCells = True;
				$worksheet4->Range("A".$r.":".chr($c+63)."".$r)->Font->Bold = True;
				$worksheet4->Cells($r++,1)->value = $result3[Sublevel];
			}
			$c=1;
			if($portid!=$result3[Port_ID])
			{
				$portcount=0;
				$worksheet4->Cells($r,$c)->value = $result3[Website];
			}
			else
				$portcount++;
			if($portcount2 && !$portcount)
			{
				$worksheet4->Range("A".($r-$portcount2-1).":A".($r-1))->MergeCells = True;
				$worksheet4->Range("A".($r-$portcount2-1).":A".($r-1))->VerticalAlignment = 1;
			}
			$c++;
			$worksheet4->Cells($r,$c++)->value = $result3[Rubrik];
			if($show_imp || $rotid==45)
				$worksheet4->Cells($r,$c++)->value = $result3[PI_Rubrik];
			if($show_view && $rotid!=45)
				$worksheet4->Cells($r,$c++)->value = $result3[Visits_Rubrik];
			if($show_unique && $rotid!=45){
//				$worksheet4->Cells($r,$c)->value = $result3[uniqueuser_Rubrik];
				$c++;
			}
			$worksheet4->Cells($r,$c)->value = $result3[Website];
			$worksheet4->Cells($r,$c)->Hyperlinks->Add($worksheet4->Cells($r,$c) ,'http://www.netpoint-media.de/portfolio/'.$result3[Website].'.html');
			$r++;
			$c++;
			$portcount2=$portcount;
//			$rotname=$result[RotName];
//			$rotname2=$result[RotName];
//			if(strlen($rotname)>31)
//				$rotname = substr($rotname,0,30).".";
//			$linkname=$result[Linkname];
//			$subchannel=$result[Sublevel];
//			$portid=$result[Port_ID];
		}
	}
//	if($worksheet4->Cells(($r-1),($c-1))->value=="")
//		$worksheet4->Range("A".($r-2).":A".($r-1))->MergeCells = True;
	$r--;
		$worksheet4->Range("".chr($c+64)."3:".chr($c+64)."".$r)->MergeCells = True;
		$worksheet4->Range("".chr($c+64)."3:".chr($c+64)."".$r)->HorizontalAlignment = 3;
		$worksheet4->Range("".chr($c+64)."3:".chr($c+64)."".$r)->VerticalAlignment = 2;
		$worksheet4->Name = $rotname;
		$worksheet4->Range("".chr($c+64)."3:".chr($c+64)."".$r)->Font->Bold = True;
		$worksheet4->Range("C2:".chr($c+62)."".$r)->NumberFormat = '#.###';
		for($i=1;$i<5;$i++)
		{
			$worksheet4->Range("A2:".chr($c+64)."".$r)->Borders($i)->LineStyle = 1;
			$worksheet4->Range("A2:".chr($c+64)."".$r)->Borders($i)->Weight = 2;
			$worksheet4->Range("A2:".chr($c+64)."".$r)->Borders($i)->ColorIndex = 1;
		}
		$worksheet4->Cells(3,$c)->value = $rotname2;
		$worksheet4->Cells(3,$c)->Hyperlinks->Add($worksheet4->Cells(3,$c) ,'http://www.netpoint-media.de/rotation/'.$linkname.'.html');
		$c=3;
		$r++;
		if($show_imp || $rotid==45)
		{
			$worksheet4->Cells($r,$c)->value = "=SUMME(".chr($c+64)."3:".chr($c+64)."".($r-1);
			$c++;
		}
		if($show_view && $rotid!=45)
		{
			$worksheet4->Cells($r,$c)->value = "=SUMME(".chr($c+64)."3:".chr($c+64)."".($r-1);
			$c++;
		}
		if($show_unique && $rotid!=45)
		{
			if($rotid == 'agof_titel') {
				$query12 = @mysql_query("SELECT uuser*1000000 UU,datum FROM factsheet_inhalt_neu WHERE filter1='Basis' AND filter1='Basis' AND filter2='' AND filter3='' AND filter4='' AND filter5='' AND website = 'netpoint_media_Vermarkterreichweite';");
				$result12 = @mysql_fetch_array($query12);
				$worksheet4->Cells($r,$c)->value = $result12[UU];
				$worksheet4->Cells($r,$c)->NumberFormat = '#.###';
				$c++;
			}
			else {
				$worksheet4->Cells($r,$c)->value = "=SUMME(".chr($c+64)."3:".chr($c+64)."".($r-1);
				$c++;
			}
		}
		$worksheet4->Range("C".$r.":".chr($c+64)."".$r)->Font->Bold = True;
		$worksheet4->Range("C".$r.":".chr($c+64)."".$r)->Font->Underline = 5;
}
echo "Impressions, Visits, Unique<br/>";
flush();
createxls(1,1,1,"C:/apachefriends/xampp/htdocs/g-tool/excel/ivu.xlsx");
flush();
echo "Impressions, Visits<br/>";
createxls(1,1,0,"C:/apachefriends/xampp/htdocs/g-tool/excel/iv.xlsx");
flush();
echo "Impressions, Unique<br/>";
createxls(1,0,1,"C:/apachefriends/xampp/htdocs/g-tool/excel/iu.xlsx");
flush();
echo "Visits, Unique<br/>";
createxls(0,1,1,"C:/apachefriends/xampp/htdocs/g-tool/excel/vu.xlsx");
flush();
echo "Impressions<br/>";
createxls(1,0,0,"C:/apachefriends/xampp/htdocs/g-tool/excel/i.xlsx");
flush();
echo "Visits<br/>";
createxls(0,1,0,"C:/apachefriends/xampp/htdocs/g-tool/excel/v.xlsx");
flush();
echo "Unique<br/>";
createxls(0,0,1,"C:/apachefriends/xampp/htdocs/g-tool/excel/u.xlsx");
echo "fertsch!";
?>
