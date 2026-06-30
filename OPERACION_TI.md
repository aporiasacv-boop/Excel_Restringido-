# PC de TI — Olnatura Excel + convivencia con Órdenes de venta

En la misma PC corren **dos APIs** y **un solo ngrok** con dos túneles.

| Servicio | Carpeta | Puerto | Túnel ngrok |
|----------|---------|--------|-------------|
| Órdenes de venta (Java) | `prueba_ov\dynamics-integration` | **8080** | `dynamics` |
| Login Excel NIKZON (Python) | `Excel_Restringido-\api` | **8011** | `olnatura` |

---

## 1. Git pull (dos terminales, dos carpetas)

**Terminal A — Órdenes de venta**

```powershell
cd C:\Users\IDI\Desktop\Excels\prueba_ov
git pull origin main
```

**Terminal B — Excel restringido**

```powershell
cd C:\Users\IDI\Desktop\Excels\Excel_Restringido-
git pull origin main
```

(Si en TI clonaron solo `aporiasacv-boop`, use `origin`. Si usan `olnaturaqr-suite`, use `empresa`.)

---

## 2. Una sola vez — fusionar ngrok

En la carpeta **prueba_ov** (no importa cuál terminal, solo esta ruta):

```powershell
cd C:\Users\IDI\Desktop\Excels\prueba_ov
powershell -ExecutionPolicy Bypass -File .\instalar-ngrok-coexistencia.ps1
```

Si pide token: `ngrok config add-authtoken SU_TOKEN` y repetir el script.

---

## 3. Cada reinicio de PC — tres ventanas abiertas

### Ventana 1 — Java (Órdenes de venta)

```powershell
cd C:\Users\IDI\Desktop\Excels\prueba_ov\dynamics-integration
.\run.ps1
```

Espere: `Started DynamicsIntegrationApplication`.

### Ventana 2 — API Olnatura

```powershell
cd C:\Users\IDI\Desktop\Excels\Excel_Restringido-\api
copy .env.example .env
```

Edite `api\.env` y confirme:

```
API_PORT=8011
```

Luego:

```powershell
.\INICIAR_API.bat
```

Debe mostrar: `Iniciando Olnatura Auth en 127.0.0.1:8011`.

### Ventana 3 — ngrok (ambos túneles)

```powershell
cd C:\Users\IDI\Desktop\Excels\prueba_ov
.\INICIAR_NGROK.bat
```

Copie las **dos** URLs `https://....ngrok-free.dev`:
- **dynamics** → pegar en Excel Órdenes, hoja Resultado, celda **B1**
- **olnatura** → usar en los Excel NIKZON (paso 4)

---

## 4. URL en los Excel NIKZON

Con ngrok corriendo, en carpeta **Excel_Restringido-**:

```powershell
cd C:\Users\IDI\Desktop\Excels\Excel_Restringido-
powershell -ExecutionPolicy Bypass -File .\api\actualizar_url_ngrok.ps1
```

Cierre Excel si estaba abierto. Luego:

```powershell
.\INSTALAR.bat
```

Salida en `LISTOS\*.xlsm` → subir esos archivos a OneDrive.

---

## 5. Usuarios (solo en TI)

```powershell
cd C:\Users\IDI\Desktop\Excels\Excel_Restringido-\api
py -3.12 crear_usuario.py maria SuClave123!
py -3.12 listar_usuarios.py
```

Admin por defecto: `Admin` / `Admin123!`

---

## 6. Probar en esta PC (antes de producción)

1. Abrir un `.xlsm` desde **`LISTOS\`** (no desde OneDrive en TI).
2. **Habilitar contenido** / macros.
3. Login con usuario creado.
4. Si falla: comprobar que la URL en el error no sea `127.0.0.1` (significa VBA viejo → repetir `INSTALAR.bat`).

Verificación rápida API:

```powershell
curl http://127.0.0.1:8011/health
```

---

## 7. Producción

Los operadores abren el Excel desde **OneDrive** en **Excel de escritorio**, habilitan macros e inician sesión con su usuario. La PC de producción **no** necesita Python ni ngrok; solo internet hacia la URL `olnatura` de ngrok.
