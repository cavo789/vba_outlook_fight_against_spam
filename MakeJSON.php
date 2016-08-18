<?php
/** 
 * Author : AVONTURE Christophe - https://www.aesecure.com
 * 
 * This PHP utility will update the SPAM.json file.   The idea is to copy/paste a list of spam domain server from the internet like
 * f.i. the list that we can find at http://www.joewein.net/dl/bl/dom-bl.txt.
 * 
 * The content of such list can be just copy/pasted here, as value of the $LIST variable.  Just copy/paste the list even thousands records.
 * 
 * The PHP script will then convert that list into an array, will remove duplicates, will add the @ prefix before each entries and, finally,
 * will merge that new list with the existing SPAM.json file.   So that file will be bigger and contains new entries.
 * 
 * It's safe to run the script more than once since only new values will be appended.  So, if fired twice, the second time nothing will be added since
 * already processed.
 * 
 */

$LIST='';   // <-- copy/paste here the new list

define('DS',DIRECTORY_SEPARATOR);

   // Open the SPAM.JSON file of the Outlook VBA repository
   $filename=__DIR__.DS.'spam.json';

   if (!file_exists($filename)) $filename='C:\Christophe\Repository\outlook_vba\spam.json';
   if (!file_exists($filename)) { die('<strong>The spam.json file is missing.</strong>'); }
   
   $arrSpam=json_decode(file_get_contents($filename),true);

   echo '<pre>Number of entries in '.$filename.' : '.count($arrSpam,true).'</pre>';
   
   // Convert the big list $LIST into an array.  Each entry should start with a @
   if (trim($LIST)!='') {
      
      $LIST = preg_replace('/[\n\r ;,|]+/', ';@', '@'.$LIST);
      $arrNew=array_unique(explode(';',$LIST));

      echo '<pre>Number of entries to add, before the merge '.count($arrNew,true).'</pre>';

      $arrNew=array_unique(array_merge($arrNew,$arrSpam));

      echo '<pre>After the merge of the two lists '.count($arrNew,true).'</pre>';

      echo '<hr/>';

      echo 'Update the file '.$filename.' with the new list';
      $file = fopen($filename, 'w');
      fwrite($file, json_encode($arrNew,JSON_PRETTY_PRINT));
      fclose($file);

      echo '<h2>New file content :</h2>';
      echo '<pre>'.json_encode(array_unique($arrNew),JSON_PRETTY_PRINT).'</pre>';
      
   } else { // if (trim($LIST)!='') 
      
      echo '<strong>Please edit this script and initialize the $LIST variable.  Read comment at the top of the script to understand why and how.</strong>';
      
   }