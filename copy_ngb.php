<?php

// USAGE: copy_ndb <folder>
error_reporting(E_ALL);

date_default_timezone_set('Asia/Tokyo');

/** PHPExcel_IOFactory */
require_once 'PHPExcel/Classes/PHPExcel/IOFactory.php';

$cacheMethod = PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
PHPExcel_Settings::setCacheStorageMethod($cacheMethod);

// get source folder
$company = $argv[1];

// find souce
$fileList = glob($company . "/NGB*");
$src = "";
foreach ($fileList as $file) {
    if (stristr($file, "fixed") === FALSE) {
        $src = $file;
        break;
    }
}

// open source
echo date('H:i:s') . " Load source from $src\n";
try {
	$objSrc = PHPExcel_IOFactory::load($src);
} catch(Exception $e) {
    die('Error loading source: '.$e->getMessage());
}

if (!$objSrc) {
	echo "$src not found\n";
}

// process source file
process($objSrc, $company, "C10G");
process($objSrc, $company, "B01J");
process($objSrc, $company, "C10L");
process($objSrc, $company, "C10M");

$objSrcWriter = PHPExcel_IOFactory::createWriter($objSrc, 'Excel2007');
$objSrcWriter->save("out/" . basename(str_replace('.xlsx', '.fixed.xlsx', $src)));

// Echo done
echo date('H:i:s') . " Done writing files.\r\n";

function process($objSrc, $company, $code)
{
    // select sheet
    $sheetCount = $objSrc->getSheetCount();
    $sheet = 0;
    for ($i = 0; $i < $sheetCount; $i ++) {
        $title = $objSrc->getSheet($i)->getTitle();

        if (stristr($title, $code)) {
            $sheet = $i;
            break;
        }
    }
    $objSrcSheet = $objSrc->getSheet($sheet);

    if (!$objSrcSheet) {
        echo "Sheet $code not found\n";
    }

    // open target
    $spec = $company . "/" . $company . "*" . $code . "*";
    $fileList = glob($spec);

    foreach ($fileList as $file) {
        if (stristr($file, "fixed") === FALSE) {
            copyNGBData($objSrcSheet, $file);
        }
    }
}

function copyNGBData($objSrcSheet, $targetFile) {
        
    echo date('H:i:s') . " Load target from $targetFile\n";
    try {
        $objTarget = PHPExcel_IOFactory::load($targetFile);
    } catch(Exception $e) {
        die('Error loading target: '.$e->getMessage());
    }

    if (!$objTarget) {
        echo "$targetFile not found\n";
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
    $col = 'P';
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

                        if ($cell->hasHyperlink()) {
                            $url = $cell->getHyperlink()->getUrl();

                            $targetCell = $objTargetSheet->getCell($copy_col . $search_row);
                            $targetCell->getHyperlink()->setUrl($url);
                        }
                    }
                }

                $search_row ++;
            }
        }

        if (!$keys_found) {
            $objSrcSheet->getStyle('A' . $row)->applyFromArray(
                array(
                    'fill' => array(
                        'type' => PHPExcel_Style_Fill::FILL_SOLID
                        , 'color' => array('rgb' => 'FF0000')
                    )
                )
            );
        }

        $row ++;
    }

    echo date('H:i:s') . " Write target\n";
    $outFile = "out/" . basename(str_replace('.xlsx', '.fixed.xlsx', $targetFile));
    $objTargetWriter = PHPExcel_IOFactory::createWriter($objTarget, 'Excel2007');
    $objTargetWriter->save($outFile);

    $objTarget->disconnectWorksheets();
    unset($objTarget);
    unset($objTargetWriter);
}

?>
