<?php
$no_status=1;
error_reporting(none);
$dir=getcwd()."/excel/";
if(!$style)
{
	echo "Vorhandene Versionen:<br/><table>";
				$dh = opendir($dir);
				while ($file = readdir ($dh))
				{
					if(substr($file,0,1)!=".")
					$filenames[] = $file;
				}
				closedir($dh);
				if($filenames)
				{
					natcasesort($filenames);
//					$filenames=array_reverse($filenames);
//					echo "Newsletterdatei: <select name='dateiname'>";
					foreach($filenames as $file)
					{
							if($file != 'deckblatt.xlsx')
							echo '<tr><td><a href="export_excel.php?style='.$file.'">'.str_replace(Array("i","v","u",".xlsx"),Array("impressions ","visits ","uniques ",""),$file).'</a></td><td>'.date (" d.m.Y H:i:s", filemtime($dir."/".$file)).'</td></tr>';
					}
				}
	echo "</table><br/><a id='link' href='#'><span id='text'>Update</span></a>";
  ?>
  <script src="//code.jquery.com/jquery-1.11.3.min.js"></script>
  <script type="text/javascript">
  $(function(){
   $("#link").click(function(){
     $("#text").html('<span id ="blink">Updating...</span>');
     var element = $("#blink");
     var shown = true;
     setInterval(toggle, 500);

     function toggle() {
         if(shown) {
              element.hide();
              shown = false;
          } else {
              element.show();
              shown = true;
          }
      }
      $.post("export_portfolio_xls.php",function(data){
        console.log('ok');
        location.reload();
      });
   });
  return false;
  });
  </script>
  <?php
}
else
{
	header("Content-Type: application/vnd.ms-excel; charset='iso-8859-15'");
	header('Content-Disposition: attachment; filename="npm_portfolio_'.date ("m_Y", filemtime($dir."/".$style)).'.xlsx"');
	$fh=fopen($dir."/".$style."", "rb");
	fpassthru($fh);
	unlink($fname);
}
?>
