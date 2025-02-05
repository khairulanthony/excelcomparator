<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ExcelController extends Controller
{
    public function index()
    {
        return view('excel.index');
    }

    public function compare(Request $request)
    {
        // Add detailed validation messages
        $request->validate([
            'file1' => 'required|file|mimes:xlsx,xls',
            'file2' => 'required|file|mimes:xlsx,xls',
        ]);

        try {
            // Load files with full reading options for better compatibility
            $reader = IOFactory::createReader('Xlsx');
            $reader->setReadDataOnly(false);
            $reader->setLoadAllSheets(true);
            
            $spreadsheet1 = $reader->load($request->file('file1')->getRealPath());
            $spreadsheet2 = $reader->load($request->file('file2')->getRealPath());
            
            $worksheet1 = $spreadsheet1->getActiveSheet();
            $worksheet2 = $spreadsheet2->getActiveSheet();

            // Get and log all column headers from both files for debugging
            $headers1 = $this->getColumnHeaders($worksheet1);
            $headers2 = $this->getColumnHeaders($worksheet2);

            // Log the headers we found
            \Log::info('File 1 Headers:', $headers1);
            \Log::info('File 2 Headers:', $headers2);

            // Find the required columns using more flexible matching
            $columns = $this->findRequiredColumns($worksheet1, $worksheet2);

            if ($columns['error']) {
                return back()->with('error', $columns['message']);
            }

            // Create lookup array from second file
            $lookupData = [];
            $highestRow2 = $worksheet2->getHighestRow();
            
            for ($row = 2; $row <= $highestRow2; $row++) {
                $name = $this->cleanValue($worksheet2->getCell($columns['name2'] . $row)->getValue());
                $ic = $this->cleanValue($worksheet2->getCell($columns['ic2'] . $row)->getValue());
                
                if ($name && $ic) {
                    $lookupData[$name] = $ic;
                    \Log::info("Found mapping: $name -> $ic");
                }
            }

            // Update first file
            $updatedCount = 0;
            $highestRow1 = $worksheet1->getHighestRow();
            
            for ($row = 2; $row <= $highestRow1; $row++) {
                $name = $this->cleanValue($worksheet1->getCell($columns['name1'] . $row)->getValue());
                if (isset($lookupData[$name])) {
                    $worksheet1->setCellValue($columns['ic1'] . $row, $lookupData[$name]);
                    $updatedCount++;
                }
            }

            $writer = new Xlsx($spreadsheet1);
            $fileName = 'updated_excel_' . time() . '.xlsx';
            $savePath = storage_path('app/public/' . $fileName);
            $writer->save($savePath);

            return back()->with([
                'success' => "Updated $updatedCount records successfully",
                'download_file' => $fileName
            ]);

        } catch (\Exception $e) {
            \Log::error('Excel processing error: ' . $e->getMessage());
            \Log::error($e->getTraceAsString());
            return back()->with('error', 'Error processing files: ' . $e->getMessage());
        }
    }

    private function getColumnHeaders($worksheet)
    {
        $headers = [];
        $highestColumn = $worksheet->getHighestColumn();
        
        for ($col = 'A'; $col <= $highestColumn; $col++) {
            $cell = $worksheet->getCell($col . '1');
            $value = $cell->getValue();
            $headers[$col] = [
                'raw_value' => $value,
                'cleaned_value' => $this->cleanValue($value),
                'data_type' => $cell->getDataType(),
            ];
        }
        
        return $headers;
    }

    private function findRequiredColumns($worksheet1, $worksheet2)
    {
        $result = [
            'name1' => null,
            'ic1' => null,
            'name2' => null,
            'ic2' => null,
            'error' => false,
            'message' => ''
        ];

        // Array of possible header variations
        $nameVariations = ['name', 'nama', 'full name', 'employee name'];
        $icVariations = ['i.c. no.', 'ic no', 'ic number', 'ic', 'i.c no', 'i.c. number'];

        // Find columns in first worksheet
        foreach ($this->getColumnHeaders($worksheet1) as $col => $header) {
            $cleanedHeader = $this->cleanValue($header['raw_value']);
            if (in_array($cleanedHeader, $nameVariations)) {
                $result['name1'] = $col;
            }
            if (in_array($cleanedHeader, $icVariations)) {
                $result['ic1'] = $col;
            }
        }

        // Find columns in second worksheet
        foreach ($this->getColumnHeaders($worksheet2) as $col => $header) {
            $cleanedHeader = $this->cleanValue($header['raw_value']);
            if (in_array($cleanedHeader, $nameVariations)) {
                $result['name2'] = $col;
            }
            if (in_array($cleanedHeader, $icVariations)) {
                $result['ic2'] = $col;
            }
        }

        // Validate found columns
        if (!$result['name1'] || !$result['name2']) {
            $result['error'] = true;
            $result['message'] = 'Name column not found in one or both files.';
        } elseif (!$result['ic2']) {
            $result['error'] = true;
            $result['message'] = 'I.C. No. column not found in the second file.';
        }

        if (!$result['ic1']) {
            // If IC column doesn't exist in first file, create it
            $result['ic1'] = ++$result['name1'];
        }

        return $result;
    }

    private function cleanValue($value)
    {
        if (is_null($value)) return '';
        
        $value = (string)$value;
        $value = trim($value);
        $value = strtolower($value);
        $value = preg_replace('/\s+/', ' ', $value);
        $value = str_replace(['．', '。', '．', '・'], '.', $value); // Handle various types of periods
        $value = preg_replace('/[^a-z0-9\s\.]/', '', $value);
        
        return $value;
    }
}