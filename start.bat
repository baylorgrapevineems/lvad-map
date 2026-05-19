@echo off
cd /d "%~dp0"
echo.
echo  LVAD Patient Monitor - Baylor Scott ^& White DFW
echo  ================================================
echo.
echo  Map:   http://localhost:3001
echo  Admin: http://localhost:3001/admin.html
echo.
echo  WARNING: This app displays sensitive patient information (PHI).
echo  Ensure it is only accessible on an internal or VPN-protected network.
echo.
if not exist node_modules (
  echo  Installing dependencies...
  npm install
  echo.
)
npm start
pause
