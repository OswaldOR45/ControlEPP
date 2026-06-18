# 🦺 Control de EPP — Registro y Trazabilidad de Incumplimientos de Protección Personal

Sistema web para registrar, en piso de planta, cuándo un trabajador **incumple con su Equipo de Protección Personal (EPP)** y convertir esas faltas en decisiones de seguridad e higiene. Corre **sin servidores propios ni costos de infraestructura**: toda la solución vive dentro de Google Workspace.

- **App web (mobile-first)** — captura guiada de incumplimientos con firma electrónica del trabajador, servida con `HtmlService`.
- **Motor de reglas (backend Apps Script)** — ventana móvil de 30 días, escalamiento automático al tercer incumplimiento y reportes diario / mensual / de análisis sobre un histórico permanente.

![Status](https://img.shields.io/badge/status-en%20producción-success)
![Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-V8-4285F4)
![Frontend](https://img.shields.io/badge/Frontend-HTML%20%7C%20CSS%20%7C%20Vanilla%20JS-E34F26)
![Datos](https://img.shields.io/badge/datos-Google%20Sheets-0F9D58)
![Storage](https://img.shields.io/badge/firmas-Google%20Drive-FBBC04)

---

## 📌 El problema

El control de EPP se llevaba en papel u hojas sueltas, sin un formato común ni trazabilidad de quién incumplió, qué falló y quién firmó de enterado. Eso impedía responder preguntas clave: ¿cuántas veces ha reincidido este trabajador este mes?, ¿qué área concentra más faltas?, ¿qué EPP es el que más se incumple?, ¿cuándo un caso debe escalar al supervisor?

**Control de EPP** estandariza la captura en piso (rápida, a prueba de guantes), acumula las faltas dentro de una ventana móvil, escala solo los casos que cruzan el límite y deja un histórico limpio y firmado que alimenta los reportes de seguridad e higiene.

---

## ✨ Características principales

### 👷 Captura en piso (registro de faltas)
- **Mobile-first y a prueba de guantes**: flujo por pasos (área → trabajador → EPP → firma) con áreas táctiles grandes.
- **Catálogo configurable** de 9 tipos de EPP (casco, lentes, calzado, cofia, uniforme, barba, limpieza personal, uñas, pertenencias no autorizadas).
- **Alerta en vivo**: al seleccionar al trabajador se consulta su acumulado en la ventana activa y se avisa si ya está en el límite, pidiendo confirmación antes de registrar otra.
- **Firma electrónica** por `canvas` (mouse o touch), con opción de *omitir* cuando aplica.

### ⏳ Ventana móvil y escalamiento
- Cada falta **caduca sola** dentro de un periodo de 30 días: el expediente del trabajador se reinicia de forma predecible.
- Al alcanzar **3 faltas** en la ventana, el caso se escala automáticamente a una hoja de histórico con estado *Pendiente* para el supervisor.
- El escalamiento incluye el **detalle de las 3 faltas** (fecha + EPP) que lo dispararon.

### 📊 Reportes y análisis
- **Reporte diario**: incumplimientos del día con el acumulado en ventana por persona.
- **Reporte mensual**: agregado por trabajador y por área (mes calendario), con detección de trabajadores en alerta.
- **Análisis** (ventana configurable): áreas más problemáticas, EPP más incumplido, **top 10 de reincidentes** y bitácora de escalamientos.

### 🛠️ Administración
- Alta de personal por área (padrón).
- Inicialización idempotente de hojas y de la carpeta privada de firmas.
- **Migración idempotente** del esquema antiguo (3 faltas por fila) al modelo de eventos, sin duplicar datos al re-ejecutarla.

---

## 👥 Flujo de trabajo

```
SUPERVISOR (captura en piso)                 TRABAJADOR              SEGURIDAD E HIGIENE
   │                                             │                         │
   │ elige área y trabajador                     │                         │
   │ marca EPP incumplido ───── consulta ────────┤                         │
   │   acumulado en ventana de 30 días           │                         │
   │ ───────────────────────── pide firma ──────▶│ firma en canvas         │
   │ registra la falta ◀─────────────────────────│                         │
   │                                             │                         │
   │ si acumulado == 3  → escala automáticamente ─────────────────────────▶│ revisa histórico
   ▼                                             ▼                         ▼
              Google Sheets (_Registro append-only)  ──▶  Reportes + Análisis + Firmas (Drive)
```

El padrón (personal por área) y el registro de faltas viven en hojas separadas: el alta de trabajadores nunca se mezcla con las transacciones de incumplimiento.

---

## 🧱 Stack tecnológico

| Componente | Tecnología |
|-----------|-----------|
| Backend / reglas | Google Apps Script (runtime V8) |
| Base de datos | Google Sheets (`SpreadsheetApp`) |
| Almacenamiento de firmas | Google Drive (`DriveApp`, carpeta privada) |
| Servido de la app | `HtmlService` (Web App) |
| Frontend | HTML5, CSS3, JavaScript (vanilla, sin frameworks) |
| Puente cliente-servidor | `google.script.run` |
| UI | DM Sans · DM Mono, diseño responsive *mobile-first* |

---

## 🏗️ Arquitectura

Arquitectura *serverless* **por capas** dentro de Google Workspace: el cliente solo presenta y captura, la lógica de negocio vive en el servidor y el acceso a datos queda aislado en helpers de hoja.

```
Index.html  (cliente, una sola página por secciones)
   │  google.script.run
   ▼
Código.gs
   ├── doGet()                     Entrega la app vía HtmlService
   │
   ├── Capa de servicios           Funciones expuestas al cliente
   │     ├── getInitialData()              áreas + EPP + config en 1 viaje
   │     ├── getPersonal(area)             padrón del área
   │     ├── registrarIncumplimiento()     alta de falta + firma + escalamiento
   │     ├── getReporteDiario / Mensual()  reportes
   │     └── getAnalisis(dias)             reincidentes, áreas y EPP top
   │
   ├── Reglas de negocio
   │     └── _getEstadoFaltasFijas()       ventana de 30 días por bloques
   │
   └── Capa de datos (helpers)
         ├── _leerRegistro() / _getRegistroSheet()    log de eventos
         ├── _getHistoricoSheet()                     escalamientos
         └── _guardarFirma() / _getFirmasFolder()     PNG privado en Drive

Google Sheets                        Google Drive
   ├── 16 hojas de área (padrón)        └── "Firmas EPP (privado)"
   ├── _Registro  (append-only)
   └── _Historico (escalamientos)
```

**Decisiones de diseño destacadas:**
- **Modelo de eventos (`append-only`)**: cada falta es un renglón inmutable en `_Registro` con su propia fecha. Da histórico permanente, auditoría, caducidad natural de faltas y reportes derivables sin perder información.
- **Ventana por bloques fijos** (`_getEstadoFaltasFijas`): las faltas se agrupan en periodos de 30 días contados *desde la primera falta* del trabajador, de modo que el expediente se reinicia de forma predecible —no es una ventana puramente deslizante.
- **Privacidad por diseño**: las firmas se guardan en una carpeta de Drive que **nunca** se hace pública (se evita deliberadamente `setSharing(ANYONE_WITH_LINK)`); el acceso de supervisores se da compartiendo la carpeta una sola vez.
- **Carga inicial de un solo viaje** (`getInitialData`): áreas, catálogo de EPP y configuración llegan en una única llamada para minimizar el costo de invocar Apps Script.
- **Autoinicialización y migración idempotente**: hojas y carpeta se crean *on-demand*, y `migrarDatos()` puede correrse varias veces sin duplicar registros.
- **Frontend sin dependencias**: cero build, cero framework, cero peso de bundle.

---

## ⏳ Lógica destacada: ventana de cumplimiento de 30 días

```
faltas del trabajador, ordenadas por fecha
   │
   ▼
toma la 1ª falta como inicio de una ventana de 30 días
   ├── agrupa todas las faltas que caen dentro de esos 30 días
   ├── si HOY cae dentro de la ventana → esta es la ventana ACTIVA
   │        → acumulado = nº de faltas en ella
   │        → diasRestantes = días hasta que el expediente se reinicia
   └── si la ventana ya venció → arranca la siguiente desde la falta posterior
```

Sobre este conteo se decide todo: el aviso en vivo al capturar, el bloqueo con confirmación al superar el límite y el escalamiento automático al histórico cuando el acumulado llega a **3**.

---

## 🚀 Instalación y despliegue

> Requisitos: una cuenta de Google con acceso a Google Sheets y Google Drive.

1. **Crear la hoja de cálculo.** Crea un Google Sheet y copia su **ID** (la cadena entre `/d/` y `/edit` de la URL).

2. **Crear el proyecto de Apps Script.** Desde la hoja: `Extensiones → Apps Script`.

3. **Agregar el código.**
   - Pega el contenido de `Código.gs` en el archivo de script.
   - Crea un archivo HTML llamado **`Index`** y pega el contenido de `Index.html`.

4. **Configurar el ID** y los catálogos (ver abajo).

5. **Inicializar.** Ejecuta `initSheets()` una vez desde el editor para crear las 16 hojas de área, `_Registro`, `_Historico` y la carpeta de firmas. Autoriza los permisos cuando se soliciten.

6. **Desplegar como aplicación web.** `Implementar → Nueva implementación → Aplicación web` · Ejecutar como: *Yo* · Acceso: según tu política. Copia la URL pública.

7. *(Opcional)* **Migrar datos antiguos.** Si vienes del esquema previo, ejecuta `migrarDatos()` una sola vez.

Constantes de configuración al inicio de `Código.gs`:

```javascript
const SPREADSHEET_ID = "TU_ID_DE_HOJA";   // ← obligatorio
const VENTANA_DIAS   = 30;                // ventana móvil para contar faltas
const LIMITE_FALTAS  = 3;                 // faltas que disparan el escalamiento
```

También puedes ajustar los catálogos `AREAS` y `EPP_ITEMS` según tu planta.

---

## 🗃️ Modelo de datos

Google Sheets funciona como base de datos, separando el **padrón** del **registro de eventos**.

- **16 hojas de área** (padrón): el nombre de cada trabajador vive en la columna `B` a partir de la fila 3.
- **`_Registro`** (log *append-only*, una falta por renglón): `FECHA · AREA · NOMBRE · EPP · FIRMA` (URL en Drive).
- **`_Historico`** (escalamientos, generado automáticamente): `FECHA ESCALAMIENTO · AREA · NOMBRE · FALTAS EN VENTANA · DETALLE · ESTADO`.
- **Firmas**: PNG en la carpeta privada `Firmas EPP (privado)` de Drive; en `_Registro` solo se guarda la URL.

---

## 🛣️ Roadmap

- [ ] Autenticación de supervisores y registro de quién captura cada falta.
- [ ] Exportación de reportes a PDF y envío automático al área de seguridad.
- [ ] Dashboard con gráficas de tendencia por área y por tipo de EPP.
- [ ] Notificaciones (correo / WhatsApp) al generarse un escalamiento.
- [ ] Capturas de pantalla / GIF del flujo en este README.
- [ ] Pruebas automatizadas de la lógica de ventana y escalamiento.

---

## 🔒 Seguridad

El `SPREADSHEET_ID` identifica la base de datos completa y debe mantenerse fuera de repositorios públicos (en el repo va vacío de forma intencional). Las firmas de los trabajadores se guardan en una carpeta de Drive **privada**, sin enlace público: el acceso se concede compartiendo la carpeta una sola vez con quien corresponda.

---

## 👤 Autor

**Oswaldo Reynoso Robles** — diseño e implementación completa: modelo de eventos *append-only*, lógica de ventana móvil y escalamiento, captura mobile-first con firma electrónica y reportes de análisis. Proyecto desarrollado como practicante de IT para digitalizar el control de EPP de la planta.
