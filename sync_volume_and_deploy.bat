@echo off
setlocal

rem Sync Fly volume data into repo, commit/push, then fly deploy.
rem Requirements: flyctl logged in; git configured.

set "APP=akare-emisyon"
set "ROOT=%~dp0"
set "DATA_DL=%ROOT%data-downloaded\data"

echo [1/5] Fetching latest /data from Fly volume...
if not exist "%ROOT%fetch_fly_data.bat" (
  echo ERROR: fetch_fly_data.bat not found at "%ROOT%fetch_fly_data.bat"
  exit /b 1
)
call "%ROOT%fetch_fly_data.bat"
if errorlevel 1 (
  echo Fetch failed. Aborting.
  exit /b 1
)

echo [2/5] Copying data files into repo...
if not exist "%DATA_DL%" (
  echo Downloaded data folder not found: %DATA_DL%
  exit /b 1
)

pushd "%ROOT%"
set FILES=firma_kayit.json teklif.json firma_olcum.json saha_olc.json parameters.json baca_bilgileri.json parametre_olcum.json parametre_sahabil.json asgari_fiyatlar.json forms.json users.json used_teklif_numbers.json par_saha_header_groups.json
for %%F in (%FILES%) do (
  if exist "%DATA_DL%\%%F" (
    copy /Y "%DATA_DL%\%%F" "%%F" >nul
    echo   synced %%F
  ) else (
    echo   missing in download: %%F
  )
)

echo [3/5] Git add/commit/push...
git add %FILES%
git commit -m "Sync volume data before deploy" || echo (no changes to commit)
git push origin main
if errorlevel 1 (
  echo Git push failed. Aborting.
  popd
  exit /b 1
)

echo [4/5] Fly deploy...
flyctl deploy
if errorlevel 1 (
  echo Fly deploy failed.
  popd
  exit /b 1
)

echo [5/5] Done.
popd
endlocal
