# Arquitectura — Olnatura Excel Platform

## Visión general

Olnatura es una aplicación empresarial interna donde **Excel actúa exclusivamente como interfaz de usuario**. Toda la información persistente reside en **PostgreSQL**. El código VBA se versiona en **Git** y se despliega mediante `git pull` sin modificaciones manuales post-despliegue.

Esta fase establece únicamente la **infraestructura arquitectónica**. No existe funcionalidad visible para el usuario final.

---

## Decisiones arquitectónicas

### 1. Excel como shell, no como motor de datos

Excel no almacena lógica de negocio ni datos maestros. Su rol se limita a presentar información y capturar interacción del usuario. Esta separación evita que el libro se convierta en un sistema frágil dependiente de fórmulas, macros dispersas o datos embebidos.

### 2. Arquitectura por capas estricta

Cada capa tiene una responsabilidad única y depende solo de las capas inferiores. Ningún módulo puede saltarse la jerarquía:

```
UI  →  Services  →  Repositories  →  Database  →  PostgreSQL
         ↓              ↓               ↓
      Logging        Logging         Config
         ↓
      Security (transversal)
```

Esta restricción facilita pruebas, mantenimiento y sustitución de componentes sin efecto dominó.

### 3. Punto único de arranque

Todo el sistema inicia desde `InitializeApplication()` en `modBootstrap`. Esto garantiza un orden de inicialización predecible y evita estados parciales donde algunos módulos están listos y otros no.

El orquestador interno es `clsApplicationContext`, que actúa como contenedor de dependencias del sistema.

### 4. Configuración externa centralizada

Toda configuración vive en `config.ini` en la raíz del proyecto. Ningún valor configurable se escribe en código VBA. `clsConfigManager` es el único módulo autorizado a leer ese archivo.

Esta decisión permite cambiar entornos (desarrollo, producción) sin recompilar ni editar macros.

### 5. DatabaseManager como única puerta a PostgreSQL

Ningún repositorio, servicio o formulario abrirá conexiones ADODB directamente. Toda comunicación con la base de datos pasa por `clsDatabaseManager`. Esto centraliza manejo de conexiones, transacciones, reintentos y logging de errores de persistencia.

### 6. Código fuente en Git, no dentro del .xlsm

Los módulos VBA se almacenan como archivos `.bas` y `.cls` en `src/`. El libro Excel (.xlsm) es el contenedor de ejecución, pero la fuente de verdad es el repositorio Git. El flujo de despliegue es:

```
Desarrollo → git push → Producción → git pull → Abrir Excel → Funcionar
```

### 7. Clases base para extensibilidad

`clsRepositoryBase` y `clsServiceBase` definen el contrato de las capas de datos y negocio. Las implementaciones futuras seguirán estos patrones, evitando divergencia arquitectónica conforme crezca el equipo.

---

## Estructura del proyecto

```
/
├── ARCHITECTURE.md          Documentación arquitectónica
├── config.ini               Configuración externa del sistema
├── .gitignore               Exclusiones de Git
└── src/
    ├── Core/                Arranque y orquestación
    ├── Configuration/       Lectura de config.ini
    ├── Database/            Acceso a PostgreSQL
    ├── Repositories/        Acceso a datos por entidad
    ├── Services/            Lógica de negocio
    ├── Security/            Autenticación y sesión
    ├── Logging/             Registro de eventos
    ├── UI/                  Presentación (formularios, helpers)
    └── Utils/               Utilidades generales
```

---

## Responsabilidad de cada módulo

| Módulo | Archivo | Responsabilidad |
|--------|---------|-----------------|
| **Bootstrap** | `Core/modBootstrap.bas` | Punto de entrada público. Expone `InitializeApplication` y `ShutdownApplication`. |
| **ApplicationContext** | `Core/clsApplicationContext.cls` | Orquesta el ciclo de vida. Crea y conecta todos los componentes del sistema. |
| **ConfigManager** | `Configuration/clsConfigManager.cls` | Lee y expone `config.ini`. Único acceso a configuración. |
| **DatabaseManager** | `Database/clsDatabaseManager.cls` | Gestiona conexión y ejecución de comandos SQL. Único acceso a PostgreSQL. |
| **RepositoryBase** | `Repositories/clsRepositoryBase.cls` | Plantilla para repositorios por entidad. Delega a DatabaseManager. |
| **ServiceBase** | `Services/clsServiceBase.cls` | Plantilla para servicios de negocio. Orquesta repositorios. |
| **AuthenticationManager** | `Security/clsAuthenticationManager.cls` | Autenticación, sesión y autorización (fase futura). |
| **Logger** | `Logging/clsLogger.cls` | Registro centralizado de eventos del sistema. |
| **LogLevel** | `Logging/modLogLevel.bas` | Enumeración de niveles de log. |
| **UIHelpers** | `UI/modUIHelpers.bas` | Utilidades de presentación compartidas. |
| **Utils** | `Utils/modUtils.bas` | Funciones generales sin lógica de negocio. |

---

## Dependencias entre capas

```
modBootstrap
    └── clsApplicationContext
            ├── clsConfigManager
            │       └── modUtils (rutas, archivos)
            ├── clsLogger
            │       └── clsConfigManager
            ├── clsDatabaseManager
            │       ├── clsConfigManager
            │       └── clsLogger
            └── clsAuthenticationManager
                    ├── clsDatabaseManager
                    └── clsLogger

[Futuro]
clsServiceBase → clsRepositoryBase → clsDatabaseManager
modUIHelpers   → clsServiceBase (nunca a repositorios ni database)
```

### Reglas de dependencia

1. **UI** solo puede llamar a **Services** y **UIHelpers**.
2. **Services** solo puede llamar a **Repositories** y **Logger**.
3. **Repositories** solo puede llamar a **DatabaseManager**.
4. **DatabaseManager** solo puede llamar a **ConfigManager** y **Logger**.
5. **Security** es transversal: Services y UI consultan autenticación, pero Security no conoce UI.
6. **Utils** no depende de ninguna otra capa del dominio.

---

## Flujo de ejecución

### Arranque

```
1. Excel abre el libro (.xlsm)
2. Auto_Open o acción del usuario invoca InitializeApplication()
3. ApplicationContext.Initialize() ejecuta en orden:
   a. ConfigManager.Load()      → valida existencia de config.ini
   b. Logger.Initialize()       → prepara sistema de registro
   c. DatabaseManager.Initialize() → recibe config y logger
   d. AuthenticationManager.Initialize() → recibe database y logger
4. Sistema listo (sin funcionalidad visible en esta fase)
```

### Apagado

```
1. ShutdownApplication() invoca ApplicationContext.Shutdown()
2. Se cierran componentes en orden inverso:
   Authentication → Database → Logger → Config
3. Referencias liberadas, estado limpio
```

---

## Configuración (`config.ini`)

El archivo se ubica junto al libro Excel en la raíz del proyecto. Secciones preparadas:

| Sección | Propósito |
|---------|-----------|
| `[Application]` | Nombre, versión y entorno de ejecución |
| `[Database]` | Parámetros de conexión PostgreSQL (fase futura) |
| `[Logging]` | Nivel y destino de logs (fase futura) |
| `[Security]` | Tiempo de sesión y parámetros de auth (fase futura) |

`clsConfigManager` valida la existencia del archivo en `Load()`. La lectura completa de valores se implementará en la siguiente fase.

---

## Cómo crecerá el proyecto

### Fase 2 — Configuración y Logging operativos
- Parser completo de `config.ini`
- Escritura de logs a archivo según nivel configurado
- Registro de eventos de arranque y error

### Fase 3 — Conexión PostgreSQL
- Implementar `Connect()` / `Disconnect()` en DatabaseManager
- Driver ADODB + ODBC para PostgreSQL
- Manejo de errores de conexión y reintentos

### Fase 4 — Autenticación
- Login contra tabla de usuarios en PostgreSQL
- Gestión de sesión con timeout configurable
- Integración con Security en el flujo de arranque

### Fase 5 — CRUD y Repositorios
- Repositorios concretos por entidad (ej. `clsUserRepository`)
- Operaciones Create, Read, Update, Delete
- Sin lógica de negocio en repositorios

### Fase 6 — Servicios de negocio
- Servicios concretos que orquestan repositorios
- Validaciones y reglas de dominio
- Punto de entrada para la capa UI

### Fase 7 — Interfaz de usuario
- Formularios Excel (UserForms)
- Formularios delegan a Services, nunca a Database
- Reportes y visualizaciones

### Fase 8 — Seguridad avanzada
- Roles y permisos
- Auditoría de operaciones críticas
- Protección de módulos VBA

---

## Estándares de código

- `Option Explicit` en todos los módulos
- Variables tipadas con nombres descriptivos
- Sin variables globales innecesarias
- Comentarios solo cuando aportan contexto no evidente
- Clases con `VB_PredeclaredId = True` solo para singletons intencionales (`ApplicationContext`)
- Prefijos: `cls` para clases, `mod` para módulos estándar

---

## Importación de módulos al libro Excel

Para integrar el código fuente en el libro de trabajo:

1. Abrir el Editor VBA (`Alt + F11`)
2. Importar cada archivo de `src/` mediante *Archivo → Importar archivo...*
3. Mantener la estructura lógica de carpetas usando nombres de módulo descriptivos
4. Colocar `config.ini` en la misma carpeta que el `.xlsm`
5. En `ThisWorkbook`, invocar `InitializeApplication` desde `Workbook_Open` (fase futura)

El repositorio Git es la fuente de verdad. Tras `git pull`, reimportar módulos modificados o usar una herramienta de sincronización VBA.

---

## Lo que NO existe en esta fase

- Conexión a PostgreSQL
- Autenticación de usuarios
- Formularios funcionales
- CRUD de entidades
- Reportes
- Funcionalidad visible para el usuario

El sistema arranca, valida configuración y queda en estado listo para las siguientes fases.
