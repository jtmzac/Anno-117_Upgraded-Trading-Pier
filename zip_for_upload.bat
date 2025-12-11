for %%I in (.) do set CurrDirName=%%~nxI

SET p1="%CurrDirName%\%CurrDirName%.zip"
SET p2="%CurrDirName%\modinfo.json"
SET p3="-ir!%CurrDirName%\data\*"


del "%CurrDirName%.zip"


cd ..

"C:\Program Files\7-Zip\7z.exe" a %p1% %p2% %p3%

