@echo off
REM ============================================================
REM  KBT Universal Python Tools — Build All .exe Files
REM  Run this from the ExeTest\python_exe folder
REM ============================================================
REM
REM  PREREQUISITES:
REM    1. Python 3.10+ installed
REM    2. Run: pip install -r requirements.txt
REM
REM  This script builds all 22 Python tools into standalone .exe
REM  files. Each .exe can be double-clicked or run from command
REM  line — no Python installation needed on the target machine.
REM
REM  Output: dist\ folder with all .exe files
REM ============================================================

echo.
echo ============================================================
echo  KBT Universal Python Tools — Building .exe Files
echo ============================================================
echo.

REM Set source paths
set SRC=..\..\UniversalToolsForAllFiles\python
set SRC_NEW=..\..\UniversalToolsForAllFiles\python\NewTools

REM Clean previous builds
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [1/22] Building clean_data.exe...
pyinstaller --onefile --name "KBT_CleanData" --icon=NONE "%SRC%\clean_data.py"
if errorlevel 1 echo    *** FAILED: clean_data ***

echo [2/22] Building compare_files.exe...
pyinstaller --onefile --name "KBT_CompareFiles" --icon=NONE "%SRC%\compare_files.py"
if errorlevel 1 echo    *** FAILED: compare_files ***

echo [3/22] Building consolidate_files.exe...
pyinstaller --onefile --name "KBT_ConsolidateFiles" --icon=NONE "%SRC%\consolidate_files.py"
if errorlevel 1 echo    *** FAILED: consolidate_files ***

echo [4/22] Building consolidate_budget.exe...
pyinstaller --onefile --name "KBT_ConsolidateBudget" --icon=NONE "%SRC%\consolidate_budget.py"
if errorlevel 1 echo    *** FAILED: consolidate_budget ***

echo [5/22] Building variance_analysis.exe...
pyinstaller --onefile --name "KBT_VarianceAnalysis" --icon=NONE "%SRC%\variance_analysis.py"
if errorlevel 1 echo    *** FAILED: variance_analysis ***

echo [6/22] Building variance_decomposition.exe...
pyinstaller --onefile --name "KBT_VarianceDecomposition" --icon=NONE "%SRC%\variance_decomposition.py"
if errorlevel 1 echo    *** FAILED: variance_decomposition ***

echo [7/22] Building aging_report.exe...
pyinstaller --onefile --name "KBT_AgingReport" --icon=NONE "%SRC%\aging_report.py"
if errorlevel 1 echo    *** FAILED: aging_report ***

echo [8/22] Building bank_reconciler.exe...
pyinstaller --onefile --name "KBT_BankReconciler" --icon=NONE "%SRC%\bank_reconciler.py"
if errorlevel 1 echo    *** FAILED: bank_reconciler ***

echo [9/22] Building gl_reconciliation.exe...
pyinstaller --onefile --name "KBT_GLReconciliation" --icon=NONE "%SRC%\gl_reconciliation.py"
if errorlevel 1 echo    *** FAILED: gl_reconciliation ***

echo [10/22] Building reconciliation_exceptions.exe...
pyinstaller --onefile --name "KBT_ReconExceptions" --icon=NONE "%SRC%\reconciliation_exceptions.py"
if errorlevel 1 echo    *** FAILED: reconciliation_exceptions ***

echo [11/22] Building fuzzy_lookup.exe...
pyinstaller --onefile --name "KBT_FuzzyLookup" --icon=NONE "%SRC%\fuzzy_lookup.py"
if errorlevel 1 echo    *** FAILED: fuzzy_lookup ***

echo [12/22] Building master_data_mapper.exe...
pyinstaller --onefile --name "KBT_MasterDataMapper" --icon=NONE "%SRC%\master_data_mapper.py"
if errorlevel 1 echo    *** FAILED: master_data_mapper ***

echo [13/22] Building forecast_rollforward.exe...
pyinstaller --onefile --name "KBT_ForecastRollforward" --icon=NONE "%SRC%\forecast_rollforward.py"
if errorlevel 1 echo    *** FAILED: forecast_rollforward ***

echo [14/22] Building batch_process.exe...
pyinstaller --onefile --name "KBT_BatchProcess" --icon=NONE "%SRC%\batch_process.py"
if errorlevel 1 echo    *** FAILED: batch_process ***

echo [15/22] Building unpivot_data.exe...
pyinstaller --onefile --name "KBT_UnpivotData" --icon=NONE "%SRC%\unpivot_data.py"
if errorlevel 1 echo    *** FAILED: unpivot_data ***

echo [16/22] Building regex_extractor.exe...
pyinstaller --onefile --name "KBT_RegexExtractor" --icon=NONE "%SRC%\regex_extractor.py"
if errorlevel 1 echo    *** FAILED: regex_extractor ***

echo [17/22] Building pdf_extractor.exe...
pyinstaller --onefile --name "KBT_PDFExtractor" --icon=NONE "%SRC%\pdf_extractor.py"
if errorlevel 1 echo    *** FAILED: pdf_extractor ***

echo [18/22] Building word_report.exe...
pyinstaller --onefile --name "KBT_WordReport" --icon=NONE "%SRC%\word_report.py"
if errorlevel 1 echo    *** FAILED: word_report ***

echo [19/22] Building sql_query_tool.exe...
pyinstaller --onefile --name "KBT_SQLQueryTool" --icon=NONE "%SRC_NEW%\sql_query_tool.py"
if errorlevel 1 echo    *** FAILED: sql_query_tool ***

echo [20/22] Building multi_file_consolidator.exe...
pyinstaller --onefile --name "KBT_MultiFileConsolidator" --icon=NONE "%SRC_NEW%\multi_file_consolidator.py"
if errorlevel 1 echo    *** FAILED: multi_file_consolidator ***

echo [21/22] Building two_file_reconciler.exe...
pyinstaller --onefile --name "KBT_TwoFileReconciler" --icon=NONE "%SRC_NEW%\two_file_reconciler.py"
if errorlevel 1 echo    *** FAILED: two_file_reconciler ***

echo [22/22] Building date_format_unifier.exe...
pyinstaller --onefile --name "KBT_DateFormatUnifier" --icon=NONE "%SRC_NEW%\date_format_unifier.py"
if errorlevel 1 echo    *** FAILED: date_format_unifier ***

echo.
echo ============================================================
echo  BUILD COMPLETE
echo ============================================================
echo.
echo  Check the dist\ folder for all .exe files.
echo  Each file is standalone — no Python needed to run it.
echo.
echo  To test any tool, open Command Prompt and run:
echo    dist\KBT_CleanData.exe --help
echo.
dir /b dist\*.exe 2>nul
echo.
pause
