<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

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

            \Log::info('File 1 Headers:', $headers1);
            \Log::info('File 2 Headers:', $headers2);

            // Find the required columns using more flexible matching
            $columns = $this->findRequiredColumns($worksheet1, $worksheet2);

            if ($columns['error']) {
                return back()->with('error', $columns['message']);
            }

            // Create data mapping from second file
            $lookupData = [];
            $highestRow2 = $worksheet2->getHighestRow();
            
            for ($row = 2; $row <= $highestRow2; $row++) {
                $name = $this->cleanValue($worksheet2->getCell($columns['name2'] . $row)->getValue());
                $ic = $this->cleanValue($worksheet2->getCell($columns['ic2'] . $row)->getValue());
                
                // Get position data if the column exists
                $position = $columns['position2'] ? 
                    $worksheet2->getCell($columns['position2'] . $row)->getValue() : 
                    '';
                
                if ($name && $ic) {
                    $lookupData[$name] = [
                        'ic' => $ic,
                        'position' => $position
                    ];
                    \Log::info("Found mapping: $name -> $ic (Position: $position)");
                }
            }

            // Process the first worksheet
            $stats = $this->processWorksheet($worksheet1, $lookupData, $columns);

            // Add new records from second file that don't exist in first file
            $stats['added'] = $this->addNewRecords($worksheet1, $lookupData, $columns, $stats['existingNames']);

            $writer = new Xlsx($spreadsheet1);
            $fileName = 'updated_excel_' . time() . '.xlsx';
            $savePath = storage_path('app/public/' . $fileName);
            $writer->save($savePath);

            $message = sprintf(
                "Updated %d records, added %d new records, removed %d records",
                $stats['updated'],
                $stats['added'],
                $stats['removed']
            );

            return back()->with([
                'success' => $message,
                'download_file' => $fileName
            ]);

        } catch (\Exception $e) {
            \Log::error('Excel processing error: ' . $e->getMessage());
            \Log::error($e->getTraceAsString());
            return back()->with('error', 'Error processing files: ' . $e->getMessage());
        }
    }

    private function processWorksheet($worksheet, $lookupData, $columns)
    {
        $stats = [
            'updated' => 0,
            'removed' => 0,
            'existingNames' => []
        ];

        $highestRow = $worksheet->getHighestRow();
        $rowsToDelete = [];
        
        // First pass: identify rows to keep or delete and update IC numbers
        for ($row = 2; $row <= $highestRow; $row++) {
            $name = $this->cleanValue($worksheet->getCell($columns['name1'] . $row)->getValue());
            
            if (isset($lookupData[$name])) {
                // Update IC number
                $worksheet->setCellValue($columns['ic1'] . $row, $lookupData[$name]['ic']);
                
                // Update designation/position if column exists
                if ($columns['designation1']) {
                    $worksheet->setCellValue($columns['designation1'] . $row, $lookupData[$name]['position']);
                }
                
                $stats['updated']++;
                $stats['existingNames'][] = $name;
            } else {
                // Mark row for deletion if name doesn't exist in second file
                $rowsToDelete[] = $row;
            }
        }

        // Second pass: delete rows (in reverse order to maintain row indices)
        rsort($rowsToDelete);
        foreach ($rowsToDelete as $row) {
            $worksheet->removeRow($row, 1);
            $stats['removed']++;
        }

        return $stats;
    }

    private function addNewRecords($worksheet, $lookupData, $columns, $existingNames)
    {
        $addedCount = 0;
        $nextRow = $worksheet->getHighestRow() + 1;

        foreach ($lookupData as $name => $data) {
            if (!in_array($name, $existingNames)) {
                // Add new row with name, IC, and position
                $worksheet->setCellValue($columns['name1'] . $nextRow, $name);
                $worksheet->setCellValue($columns['ic1'] . $nextRow, $data['ic']);
                
                // Add position if designation column exists
                if ($columns['designation1']) {
                    $worksheet->setCellValue($columns['designation1'] . $nextRow, $data['position']);
                }
                
                $nextRow++;
                $addedCount++;
                \Log::info("Added new record: $name -> {$data['ic']} (Position: {$data['position']})");
            }
        }

        return $addedCount;
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
            'designation1' => null,
            'name2' => null,
            'ic2' => null,
            'position2' => null,
            'error' => false,
            'message' => ''
        ];

        // Array of possible header variations
        $nameVariations = ['name', 'nama', 'full name', 'employee name'];
        $icVariations = ['i.c. no.', 'ic no', 'ic number', 'ic', 'i.c no', 'i.c. number'];
        $positionVariations = ['position', 'designation', 'job title', 'role', 'jawatan'];

        // Find columns in first worksheet
        $highestColumn1 = $worksheet1->getHighestColumn();
        foreach ($this->getColumnHeaders($worksheet1) as $col => $header) {
            $cleanedHeader = $this->cleanValue($header['raw_value']);
            if (in_array($cleanedHeader, $nameVariations)) {
                $result['name1'] = $col;
            }
            if (in_array($cleanedHeader, $icVariations)) {
                $result['ic1'] = $col;
            }
            if (in_array($cleanedHeader, $positionVariations)) {
                $result['designation1'] = $col;
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
            if (in_array($cleanedHeader, $positionVariations)) {
                $result['position2'] = $col;
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

        if (!$result['designation1']) {
            // If designation column doesn't exist in first file, create it after IC
            $newCol = $result['ic1'];
            while ($newCol <= $highestColumn1) {
                $newCol++;
            }
            $result['designation1'] = $newCol;
            
            // Add the header
            $worksheet1->setCellValue($newCol . '1', 'Designation');
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