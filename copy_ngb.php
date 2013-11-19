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
$objSrcSheet = $objSrc->getSheet($argv[2]);

if (!$objSrcSheet) {
	echo "Sheet $argv[2] not found\n";
}

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

$objTargetSheet = $objTarget->getActiveSheet();

// search for target key range
echo date('H:i:s') . " find search area\n";
$row = 1;
$col = 'A';
$done = false;
$found = false;

$search_start = '';
$search_count = 0;

while (!$done) {
    $cell = $objTargetSheet->getCell($col . $row);
    $val = $cell->getValue();

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

// loop through sheet
//   search for value in key range
//     copy source to target
if ($argv[2] == 0) {
    $col = 'P';
}
else {
    $col = 'AA';
}
$row = 2;
$done = false;

while (!$done) {
    $cell = $objSrcSheet->getCell($col . $row);
    $key_list = $cell->getValue();

    if (strlen($key_list) == 0) {
        $done = true;
        break;
    }

    $keys = explode(";", $key_list);
    $keys_found = false;
    foreach ($keys as $rawkey) {
        $key = trim($rawkey);
        $search_done = false;
        $search_row = 4;

        while (!$search_done) {
            $found = false;
            $search_col = $search_start;
            for ($i = 0; $i < $search_count; $i++) {
                $cell = $objTargetSheet->getCell($search_col . $search_row);
                $val = $cell->getValue();

                if (strlen($val) == 0) {
                    if ($i == 0) {
                        $search_done = true;
                    }
                    break;
                }

                if ($val === $key) {
                    $found = true;
                    break;
                }

                $search_col ++;
            }

            if ($found) {
                $keys_found = true;
                // insert from source to target
                for ($copy_col = 'A'; $copy_col != 'Q'; $copy_col ++) {
                    $cell = $objSrcSheet->getCell($copy_col . $row);
                    $val = $cell->getValue();

                    $objTargetSheet->setCellValue($copy_col . $search_row, $val);
                }
            }

            $search_row ++;
        }
    }

    if (!$keys_found) {
        echo "Key list $key_list not found!!!\n";
    }

    $row ++;
}

echo date('H:i:s') . " Write target\n";
$objWriter = PHPExcel_IOFactory::createWriter($objTarget, 'Excel2007');
$objWriter->save(str_replace('.xlsx', '.fixed.xlsx', $argv[3]));

// Echo done
echo date('H:i:s') . " Done writing files.\r\n";

?>
