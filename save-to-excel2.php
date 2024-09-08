<?php
require 'vendor/autoload.php'; // Load PhpSpreadsheet

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Get form data
$name = $_POST['name'];
$phone = $_POST['phone'];
$email = $_POST['email'];
$message = $_POST['message'];

// Load existing spreadsheet or create a new one
$excelFile = 'form-data.xlsx';

if (file_exists($excelFile)) {
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($excelFile);
} else {
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    // Set headers if new file
    $sheet->setCellValue('A1', 'Name')
          ->setCellValue('B1', 'Phone')
          ->setCellValue('C1', 'Email')
          ->setCellValue('D1', 'Message');
}

// Get active sheet
$sheet = $spreadsheet->getActiveSheet();

// Find the next available row
$row = $sheet->getHighestRow() + 1;

// Write form data to the sheet
$sheet->setCellValue('A' . $row, $name)
      ->setCellValue('B' . $row, $phone)
      ->setCellValue('C' . $row, $email)
      ->setCellValue('D' . $row, $message);

// Save to file
$writer = new Xlsx($spreadsheet);
$writer->save($excelFile);

// Redirect or display success message
echo "Form submitted successfully. Data has been saved.";
?>
