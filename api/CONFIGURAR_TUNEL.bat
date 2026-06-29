@echo off
echo === Cloudflare Tunnel (gratis, URL publica HTTPS) ===
echo.
echo 1. Descargue cloudflared desde: https://developers.cloudflare.com/cloudflare-one/connections/connect-apps/install-and-setup/installation/
echo 2. Ejecute: cloudflared tunnel login
echo 3. Ejecute: cloudflared tunnel create olnatura-auth
echo 4. Copie config.yml.example a config.yml y ajuste tunnel id / hostname
echo 5. En otra ventana, inicie la API con INICIAR_API.bat
echo 6. Ejecute: cloudflared tunnel --config config.yml run
echo.
echo Prueba rapida (URL temporal, cambia al reiniciar):
echo   cloudflared tunnel --url http://127.0.0.1:8000
echo.
pause
