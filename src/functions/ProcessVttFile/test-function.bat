@echo off
echo Testing Azure Function VTT Processing
echo =====================================

echo.
echo Test 1: Processing Vikran file...
az rest --method GET --url "http://localhost:7071/api/ProcessVttFile?name=Vikran-xrmtool2.vtt"

echo.
echo Test 2: Processing small file...
az rest --method GET --url "http://localhost:7071/api/ProcessVttFile?name=field app-notesai-na.vtt"

echo.
echo Testing complete!
pause