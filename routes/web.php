<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\ExcelController;

Route::get('/', [ExcelController::class, 'index'])->name('excel.index');
Route::post('/compare', [ExcelController::class, 'compare'])->name('excel.compare');