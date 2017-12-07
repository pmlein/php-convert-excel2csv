<?php

//include the file that loads the PhpSpreadsheet classes
// Output file download https://preview.c9users.io/plein/php-csv-calendar/events1.csv
require 'vendor/autoload.php';
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);


// Output file charset is UTF-8, it seems to work with google calendar
const OUTPUTFILE = "events.csv";

//create directly an object instance of the IOFactory class, and load the xlsx file
    $fxls ='excel.xlsx';

    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($fxls);

    $ws_count = $spreadsheet->getSheetCount();
    $nameArray = $spreadsheet->getSheetNames();
    $csv_data = null;
    $csv_data ='Subject,Start Date,Start Time,End Date,End Time,All Day Event,Description,Location,Private' . PHP_EOL;
    for($ws=0; $ws<$ws_count; $ws++){
        //echo PHP_EOL;
        //echo 'Worksheet ', $ws, PHP_EOL;
        $sheet = $spreadsheet->getSheet($ws)->toArray(null, true, true, true);
        $worksheet_data = "";
        $csv_data .= handleWorkSheet($sheet, $worksheet_data, $nameArray[$ws]);
    }
    $myfile = fopen(OUTPUTFILE, "w") or die("Unable to open file!");
    fwrite($myfile, $csv_data);
    fclose($myfile);
    
/*
* Handles one worksheet each time called
*/
function handleWorkSheet($xls_data, $csv, $wsName){
    // Date to $date, in excel only if the date changes
    $date = null;
    
    $nr = count($xls_data); //number of rows
    // Handle each row of worksheet
    $date = "";
    $differDate = "";
    $rowCount = 0;
    for($i=2; $i<=$nr; $i++){
        if (!empty($xls_data[$i])) {
        if (!empty($xls_data[$i]['A']) ) {
            $date = formatDateFinnish($xls_data[$i]['A']);
            $differDate = formatDate($xls_data[$i]['A']);
        }
        if (empty($date)) {
            throw new Exception("Error: date missing: " . $xls_data[$i]);
        } else{
            $xls_data[$i]['A'] = $date; 
        }
        // EA tieto voi olla joukkue-kentassa, jolloin otetaan description-kenttään
        $possible_ea = "";
        $team = trim($xls_data[$i]['D']);
        $eaName = "";
        if ((substr($team,0,2) == 'EA') && (strlen($team)>3)) {
             $eaName = $team;
             //print_r($eaName);
            //print_r(PHP_EOL);
        }
        if (trim($xls_data[$i]['D']) == 'IA' or !empty($eaName)) {
            $rowarray = array_values($xls_data[$i]);
           //echo "Rivi: ";
           //print_r($rowarray);
            //print_r(PHP_EOL);
            $row="";
            $row = handleRow($rowarray,$eaName,$differDate); 
            $csv .=$row;
            $csv .=PHP_EOL;
            $rowCount = $rowCount + 1;
        }
        }
    }
   // $csv .=PHP_EOL;
    // Print csv file information of worksheet
    print_r($csv);
    echo PHP_EOL, "Workseet ", $wsName, ": ", $rowCount, " events will be written to the file ", OUTPUTFILE, PHP_EOL, PHP_EOL;
    return $csv;
}

/*
 * Subject,Start Date,Start Time,End Date,End Time,All Day Event,Description,Location,Private
 */
function handleRow($row,$eaName,$differDate){
    $csv_row = "";
    $sDay = "";
    $eDay = "";
    $sTime = "";
    $eTime = "";
    $allDay = 'False';
    $desc = "";
    $loc = "";
    $priv = 'True';
    //$sYearDay = "";
// Subject
    $subject = trim($row[4]). " (" . trim($row[6]). ") " . trim($row[1]). "-" . trim($row[2]);
// Start date
    if (!empty($row[0])) {
        $sDay = trim($row[0]);
    }
// Start time
    // To get day light saving time + local time difference 
    // (google adds this automatically in import to google calendar, so has to be subtracted from time hour)
   
    $timezone = "Europe/Helsinki";
    $localAt21 = new DateTime($differDate, new DateTimeZone($timezone));
    $differ = $localAt21->getOffset() / 3600;

    if (!empty($row[1])) {
        //$sTimeH = csvTime(trim($row[1]),"",":");
        // if need for time zone or day light saving times: use variable differ (+ 3, summer time)
        $sTimeH = csvTime(trim($row[1]),"",":") + $differ;
        $sTimeM = csvTime(trim($row[1]),":","");
        $sTime = $sTimeH . ':' . $sTimeM;
    }
    
// End date
    if (!empty($row[0])) {
        $eDay = trim($row[0]);
    }
// End time
    if (!empty($row[2])) {
       // $eTimeH = csvTime(trim($row[2]),"",":");
        $eTimeH = csvTime(trim($row[2]),"",":") + $differ;
        $eTimeM = csvTime(trim($row[2]),":","");
        $eTime = $eTimeH . ':' . $eTimeM;
    }

//Description (Coach)
    if (!empty($row[3])) {
        if (!empty($eaName)) {
            $desc = trim($row[3]);
        } else {
            $desc = trim($row[6]);
        }
    }
//Location
    if (!empty($row[0])) {
        $loc = trim($row[5]);
    }
    $csv_row .= $subject .',' . $sDay . ',' . $sTime . ',' . $eDay . ',' . $eTime . ',' . $allDay . ',' . $desc . ',' .
        $loc . ',' . $priv;
    return $csv_row;
}

/*
 * Takes hour or minute from time (12:30)
 */
function csvTime($time, $startFlag,$endFlag) {
    //$csvTime = trim($time);
    if (empty($startFlag)) {
        $startFlagPosition = 0;
    } else {
        $startFlagPosition = strpos($time, $startFlag) + 1;
    }
   if (empty($endFlag)) {
        $endFlagPosition = strlen($time);
    } else {
        $endFlagPosition = strpos($time, $endFlag, $startFlagPosition);
    }
// echo "End pos ", $endFlagPosition, PHP_EOL;
     $hourOrMin = substr($time, $startFlagPosition, $endFlagPosition - $startFlagPosition);
     //echo "Tunti tai Min ", $hourOrMin;
     return $hourOrMin;
}

/*
 * Date Ma 5.6 to format 6/5/2017 (month, day, current year)
 */
function formatDate($date){
    $csv_date = "";
    $startFlag = ".";
    if (!empty($date)) {
        // poistetaan mahdollinen piste lopusta
       $month = trim(substr($date, strpos($date, $startFlag) + 1));
       $month = str_replace('.','',$month);
       $startFlag = " ";
       $endFlag = ".";
       $startFlagPosition = strpos($date, $startFlag) + 1;
       $endFlagPosition = strpos($date, $endFlag, $startFlagPosition);
       
       // third parameter in substr is count of characters to return, not final position
       $day = substr($date, $startFlagPosition, $endFlagPosition - $startFlagPosition); 
       $csv_date = $month . '/' . $day . '/' . date("Y");
    } else {
        throw new Exception("Error: date missing. " . $xls_data[$i]);
    }
    return $csv_date;
}

/*
 * Date Ma 5.6 to format 5/6/2017 (month, day, current year)
 */
function formatDateFinnish($date){
    $csv_date = "";
    $startFlag = ".";
    if (!empty($date)) {
       $month = trim(substr($date, strpos($date, $startFlag) + 1));
       $month = str_replace('.','',$month);
       $startFlag = " ";
       $endFlag = ".";
       $startFlagPosition = strpos($date, $startFlag) + 1;
       $endFlagPosition = strpos($date, $endFlag, $startFlagPosition);
       
       // third parameter in substr is count of characters to return, not final position
       $day = substr($date, $startFlagPosition, $endFlagPosition - $startFlagPosition); 
       //$csv_date = $month . '/' . $day . '/' . date("Y");
       $csv_date = $day . '/' . $month . '/' . date("Y");
    } else {
        throw new Exception("Error: date missing. " . $xls_data[$i]);
    }
    return $csv_date;
}

?>
