<?php

ini_set('memory_limit', '1024M'); // 1 GB

ini_set('display_errors', 1);
error_reporting(E_ALL);

header('Content-Type: application/json');

require __DIR__ . '/vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;

// Check if file was uploaded
if (!isset($_FILES['excel_file'])) {
    echo json_encode(['success' => false, 'message' => 'No file uploaded']);
    exit;
}

// Create uploads folder if it doesn't exist
$targetDir = __DIR__ . '/uploads/';
if (!is_dir($targetDir)) mkdir($targetDir, 0777, true);

// Move uploaded file
$fileTmp = $_FILES['excel_file']['tmp_name'];
$fileName = basename($_FILES['excel_file']['name']);
$targetFile = $targetDir . uniqid() . '-' . $fileName;

if (!move_uploaded_file($fileTmp, $targetFile)) {
    echo json_encode(['success' => false, 'message' => 'Failed to move uploaded file']);
    exit;
}

// Get original file size in KB
$originalSize = round(filesize($targetFile) / 1024, 2);

try {
    // Load Excel file
    $spreadsheet = IOFactory::load($targetFile);
    $data = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);

    // Prepare CSV path
    $csvFileName = $targetDir . uniqid() . '.csv';
    $fp = fopen($csvFileName, 'w');

    foreach ($data as $row) {
        // Keep only non-empty cells
        $rowValues = array_filter($row, fn($cell) => $cell !== null && $cell !== '');
        if (!empty($rowValues)) {
            // Optionally trim spaces
            $rowValues = array_map('trim', $rowValues);
            fputcsv($fp, array_values($rowValues));
        }
    }

    fclose($fp);

    $csvSize = round(filesize($csvFileName) / 1024, 2);

    // Return JSON with download info
    echo json_encode([
        'success' => true,
        'original_size' => $originalSize,
        'csv_size' => $csvSize,
        'csv_file' => 'uploads/' . basename($csvFileName)
    ]);

} catch (Exception $e) {
    echo json_encode(['success' => false, 'message' => $e->getMessage()]);
}
