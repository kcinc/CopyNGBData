<?php

// USAGE: copy_ndb <src> <sheet number> <target>
// sheet number: 0 base
error_reporting(E_ALL);

date_default_timezone_set('Asia/Tokyo');

/** PHPExcel_IOFactory */
require_once 'PHPExcel/Classes/PHPExcel/IOFactory.php';

$cacheMethod = PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
PHPExcel_Settings::setCacheStorageMethod($cacheMethod);

// open source
echo date('H:i:s') . " Load source from $argv[1]\n";
try {
	$objSrc = PHPExcel_IOFactory::load($argv[1]);
} catch(Exception $e) {
    die('Error loading source: '.$e->getMessage());
}

if (!$objSrc) {
	echo "$argv[1] not found\n";
}

// select sheet
$sourceSheet = $argv[2];

// open target
echo date('H:i:s') . " Load target from $argv[3]\n";
try {
	$objTarget = PHPExcel_IOFactory::load($argv[3]);
} catch(Exception $e) {
    die('Error loading target: '.$e->getMessage());
}

if (!$objTarget) {
	echo "$argv[3] not found\n";
}

// search for target key range
echo date('H:i:s') . " find search area\n";
$row = 1;
$col = 'A';
$done = false;
$found = false;

$search_start = '';
$search_count = 0;

while (!$done) {
    $cell = $objTarget->getActiveSheet()->getCell($col . $row);
    echo "debug: cell $col.$row found\n";
    $val = $cell->getValue();

    echo "debug: $val\n";

    if (!$found) {
        if ($val === '全文明細書（US）') {
            $found = true;
            $search_start = $col;
            $search_count ++;
        }
    }
    else {
        if ($val === '全文明細書（US）') {
            $search_count ++;
        }
        else {
            $done = true;
        }
    }

    $col++;
}

echo "debug: search_start=$search_start search_count=$search_count\n";

// loop through sheet
//   search for value in key range
//     copy source to target

//echo date('H:i:s') . " Write to Excel2007 format\n";
//$objWriter = PHPExcel_IOFactory::createWriter($objSrc, 'Excel2007');
//$objWriter->save(str_replace('.php', '.xlsx', __FILE__));

// Echo done
echo date('H:i:s') . " Done writing files.\r\n";

?>
