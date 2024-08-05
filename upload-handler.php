<?php

use MergeExcel\Parser;

define('START', microtime(true)); // Start time tracking

ini_set('memory_limit', '-1'); //# might over use the memory, so remove the limit
date_default_timezone_set('Africa/Nairobi');

header('Content-Type: application/json; charset=utf-8');
header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Methods: GET, POST");


require 'vendor/autoload.php';

$tag_file = 'tag.dat';

$uploadDir = __DIR__ . '/uploads/';

// Ensure the upload directory exists
if (!file_exists($uploadDir)) {
    mkdir($uploadDir, 0777, true);
}

if (!file_exists($tag_file)) {
    touch($tag_file);
}

// Check if files were uploaded
if (!isset($_FILES['files'])) {
    http_response_code(503);
    echo json_encode(['status' => 'error', 'message' => 'No files uploaded.']);
    exit();
}

$files = $_FILES['files'];

if (!isset($requestTag) || $requestTag === '') {
   file_put_contents($tag_file, uniqid());
}
$requestTag = file_get_contents($tag_file);

// Process each file
for ($i = 0; $i < count($files['name']); $i++) {
    $originalName = $files['name'][$i];
    $tmpName = $files['tmp_name'][$i];
    $size = $files['size'][$i];
    $error = $files['error'][$i];

    // Ensure there were no errors with the file upload
    if ($error !== UPLOAD_ERR_OK) {
        http_response_code(500);
        echo json_encode(['status' => 'error', 'message' => 'File upload error code: ' . $error]);
        exit();
    }

    $uniqueName = $requestTag . '_' . basename($originalName);
    $destination = $uploadDir . $uniqueName;

    if (!move_uploaded_file($tmpName, $destination)) {
        http_response_code(500);
        echo json_encode(['status' => 'error', 'message' => 'Failed to move uploaded file.']);
        exit();
    }

    echo json_encode(['status' => 'debug', 'message' => 'uploaded files in .' .   microtime(true) - $start]);

}



$parser = new Parser();


if (!$parser->merge($uploadDir, $requestTag)) {
    http_response_code(500);
    echo json_encode(['status' => 'error', 'message' => 'Could not merge your files.']);
    exit();
}

echo json_encode(['status' => 'debug', 'message' => 'merged .' .   microtime(true) - $start]);


foreach ($parser->get_workbooks('output') as $workbook) {
    // Extract the filename without the path
    $filename = basename($workbook);

    // Check if the filename starts with the request tag
    if (strpos($filename, $requestTag) === 0) {
        http_response_code(200);
        echo json_encode(['status' => 'success', 'message' => 'merged your files to' .  $filename]);
        @unlink($tag_file);
        exit();
    } else {
        http_response_code(500);
        echo json_encode(['status' => 'error', 'message' => 'Can\'t find your file  ' .  $filename . ' with ' . $requestTag]);
        exit();
    }
}
