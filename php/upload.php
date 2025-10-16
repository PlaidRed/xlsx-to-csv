<?php

ob_start();

ini_set('memory_limit', '1024M');
ini_set('display_errors', 0);
error_reporting(E_ALL);

header('Content-Type: application/json');

function sendResponse($success, $data) {
    ob_clean();
    echo json_encode(array_merge(['success' => $success], $data));
    ob_end_flush();
    exit;
}

if (!file_exists(__DIR__ . '/vendor/autoload.php')) {
    sendResponse(false, ['message' => 'PhpSpreadsheet not installed. Run: composer require phpoffice/phpspreadsheet']);
}

try {
    require __DIR__ . '/vendor/autoload.php';
    
    if (!isset($_FILES['excel_file'])) {
        sendResponse(false, ['message' => 'No file uploaded']);
    }

    if ($_FILES['excel_file']['error'] !== UPLOAD_ERR_OK) {
        sendResponse(false, ['message' => 'Upload error code: ' . $_FILES['excel_file']['error']]);
    }

    $targetDir = __DIR__ . '/uploads/';
    if (!is_dir($targetDir)) {
        mkdir($targetDir, 0777, true);
    }

    $fileTmp = $_FILES['excel_file']['tmp_name'];
    $fileName = basename($_FILES['excel_file']['name']);
    $targetFile = $targetDir . uniqid() . '-' . $fileName;

    if (!move_uploaded_file($fileTmp, $targetFile)) {
        sendResponse(false, ['message' => 'Failed to save uploaded file']);
    }

    $originalSize = round(filesize($targetFile) / 1024, 2);

    // Load Excel - setReadDataOnly(true) strips all formatting
    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($targetFile);
    $reader->setReadDataOnly(true);
    $spreadsheet = $reader->load($targetFile);
    $sheet = $spreadsheet->getActiveSheet();
    
    // Get the actual used range (not the full grid)
    $highestRow = $sheet->getHighestDataRow();
    $highestColumn = $sheet->getHighestDataColumn();
    $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
    
    // Get raw data only from the actual used range
    $data = $sheet->rangeToArray(
        'A1:' . $highestColumn . $highestRow,
        null,      // null value
        true,      // calculate formulas
        false,     // format data
        false      // return as indexed array
    );

    // Create CSV
    $csvFileName = $targetDir . uniqid() . '.csv';
    $fp = fopen($csvFileName, 'w');

    if (!$fp) {
        sendResponse(false, ['message' => 'Failed to create CSV file']);
    }

    // Find which columns have ANY data across ALL rows (non-whitespace content)
    $columnsWithData = array();
    foreach ($data as $row) {
        foreach ($row as $colIndex => $cell) {
            $trimmed = trim((string)$cell);
            if ($trimmed !== '' && $trimmed !== null) {
                $columnsWithData[$colIndex] = true;
            }
        }
    }
    
    // Sort column indices
    $activeColumns = array_keys($columnsWithData);
    sort($activeColumns);
    
    $rowCount = 0;
    $emptyRowsSkipped = 0;
    
    foreach ($data as $row) {
        // Check if row has any non-empty content
        $hasContent = false;
        foreach ($activeColumns as $colIndex) {
            $value = isset($row[$colIndex]) ? trim((string)$row[$colIndex]) : '';
            if ($value !== '') {
                $hasContent = true;
                break;
            }
        }
        
        // Skip completely empty rows
        if (!$hasContent) {
            $emptyRowsSkipped++;
            continue;
        }
        
        // Only include columns that have data somewhere in the spreadsheet
        $cleanedRow = array();
        foreach ($activeColumns as $colIndex) {
            $value = isset($row[$colIndex]) ? trim((string)$row[$colIndex]) : '';
            $cleanedRow[] = $value;
        }
        
        // Write the row
        fputcsv($fp, $cleanedRow);
        $rowCount++;
    }

    fclose($fp);
    
    $csvSize = round(filesize($csvFileName) / 1024, 2);
    
    // Create ZIP file
    $zipFileName = $targetDir . uniqid() . '.zip';
    $zip = new ZipArchive();
    
    if ($zip->open($zipFileName, ZipArchive::CREATE) !== TRUE) {
        @unlink($targetFile);
        @unlink($csvFileName);
        sendResponse(false, ['message' => 'Failed to create ZIP file']);
    }
    
    // Add CSV to ZIP with original filename (without UUID)
    $originalBaseName = pathinfo($fileName, PATHINFO_FILENAME);
    $zip->addFile($csvFileName, $originalBaseName . '.csv');
    
    // Set maximum compression
    $zip->setCompressionName($originalBaseName . '.csv', ZipArchive::CM_DEFLATE, 9);
    
    $zip->close();
    
    $zipSize = round(filesize($zipFileName) / 1024, 2);
    
    // Calculate compression ratio
    $compressionRatio = $csvSize > 0 ? round(($zipSize / $csvSize) * 100, 1) : 0;
    
    // Clean up temporary files
    @unlink($targetFile);
    @unlink($csvFileName);

    sendResponse(true, [
        'original_size' => $originalSize,
        'csv_size' => $csvSize,
        'zip_size' => $zipSize,
        'compression_ratio' => $compressionRatio,
        'zip_file' => 'uploads/' . basename($zipFileName),
        'rows_written' => $rowCount,
        'excel_dimensions' => $highestColumn . $highestRow,
        'total_columns_in_excel' => $highestColumnIndex,
        'columns_with_data' => count($activeColumns),
        'empty_rows_skipped' => $emptyRowsSkipped
    ]);

} catch (Exception $e) {
    sendResponse(false, ['message' => 'Error: ' . $e->getMessage(), 'details' => $e->getTraceAsString()]);
} catch (Error $e) {
    sendResponse(false, ['message' => 'Fatal error: ' . $e->getMessage(), 'details' => $e->getTraceAsString()]);
}
?>