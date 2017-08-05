<!DOCTYPE>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

</head>
<body>
  
<?php 
$doc = new COM("word.application");
if (!$doc) {
  echo ('Could not initialise MS word object.\n'); 
  exit(1);
}

$doc->Documents->Open('C:\xampp\htdocs\test\game.docx'); 
$tpage = $doc->ActiveDocument->ComputeStatistics(2);
$tdoc = $doc->ActiveDocument->ComputeStatistics(0);
$a = 0;
for ($i=1; $i < $tpage; $i++) { 
     	
     $temp = str_word_count( trim($doc->Documents[1]->Range($doc->Documents[1]->GoTo(1,1,$i)->Start , $doc->Documents[1]->GoTo(1,1,$i+1)->End )->Text ) , 0 , "0123456789/*-+\|]}[{';:?=`~&^%$#@!(),.");
     $a+=$temp;
echo "page ".$i."count = ". $temp ;  
echo "</br>";
  } 

  echo "page 13 count = ".( intval($tdoc) - intval($a) )  ; 
echo "</br>";    

echo "Number of pages: " .$tpage."<br>"; 
echo "Number of words: " .$tdoc ;

// $tt = $doc->ActiveDocument->docs->Count; 
// echo "==".$tt;
$doc->Documents[1]->Close(false); 
$doc->Quit(); 
$doc = null; 
unset($doc);  
?>




</body>
</html>
