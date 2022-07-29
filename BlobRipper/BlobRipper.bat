:: If segments are too numerous for Python to handle
:: (for %i in (*.ts) do @echo file '%i') > mylist.txt
:: ffmpeg -f concat -safe 0 -i mylist.txt -c copy all.ts

set /p vidURL=Input Video Path:
set /p audURL=Input Audio Path:

set vidTS=vidall.ts
set vidMP4=vidall.mp4
set audTS=audall.ts
set audMP4=audall.mp4

copy /b %vidURL% %vidTS%
copy /b %audURL% %audTS%

ffmpeg -i %vidTS% -c copy %vidMP4%
echo Video TS converted to MP4
ffmpeg -i %audTS% -c copy %audMP4%
echo Audio TS converted to MP4

ffmpeg -i %vidMP4% -i %audMP4% -c copy output.mp4

cmd /k