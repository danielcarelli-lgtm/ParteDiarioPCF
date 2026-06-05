Markdown
# ParteDiarioPCF ⏱️

Un componente de interfaz de usuario personalizado (PCF - Power Apps Component Framework) diseñado para Dynamics 365 y Power Platform. Este control transforma la gestión de horas (entidad `msdyn_timeentry`) ofreciendo una línea de tiempo visual e interactiva para los partes de trabajo diarios en el entorno de **Field Service**.

## 🚀 Características Principales

- **Línea de tiempo interactiva:** Visualiza las imputaciones de tiempo en un formato de línea de tiempo con soporte para orientación vertical y horizontal.
- **Drag & Drop:** Mueve o ajusta la duración de las entradas de tiempo simplemente arrastrando los bloques con el ratón.
- **Acciones Rápidas:** Botones integrados para crear automáticamente bloques comunes como Almuerzo o Aprovisionamiento, además de una función inteligente para rellenar huecos vacíos.
- **Ajuste Inteligente (Doble Click):** Expande automáticamente un bloque de tiempo para cubrir los espacios vacíos adyacentes.
- **Validación Visual:** Detección automática de horas imputadas fuera de la jornada laboral y alertas visuales de solapamiento de horas (bloques cruzados).
- **Cálculo en Tiempo Real:** Muestra el total de horas imputadas frente a las horas objetivo de la jornada, indicando el tiempo faltante o extra.

# Manual de Uso - Panel Interactivo de Field Service ⏱️

Bienvenido al nuevo panel interactivo para la gestión del Parte Diario. Este componente sustituye la clásica cuadrícula de registros por una línea de tiempo visual, diseñada para agilizar y simplificar la imputación de horas de trabajo, descansos, aprovisionamientos y festivos.

---

## 1. Panel de Métricas (Totales)
Ubicado en la parte superior derecha, te muestra el balance de tu día en tiempo real:
- **⏳ Imputado:** Suma total de horas y minutos registrados.
- **Jornada:** Horas totales de tu turno o jornada laboral teórica.
- **Faltan / Extra:** Indica en **rojo** si faltan horas por completar o en **verde** si has superado la jornada o la has cumplido exactamente.

---

## 2. Botones de Acción (Barra Superior)
- **↔️ / ↕️ Vista:** Cambia la interfaz entre formato vertical u horizontal.
- **🔍 Zoom:** Cambia la escala para ver 24h o solo tu jornada laboral.
- **🔄 Refrescar:** Sincroniza los datos con el servidor.
- **🍔 Crear Almuerzo:** Registra automáticamente un tiempo de comida (por defecto 1 hora).
- **📦 Crear Aprovisionamiento:** Registra tareas de carga/descarga (por defecto 30 min).
- **Completar Huecos:** Rellena automáticamente los espacios vacíos de tu jornada con bloques de trabajo.
- **🎉 Festivo:** Rellena los espacios vacíos de tu jornada marcándolos como "Festivo".
- **🚀 Enviar Parte:** Bloquea el parte para su revisión (estado "Enviado"). **Nota:** Una vez pulsado, el registro se vuelve de solo lectura.

---

## 3. Interacción con los Bloques de Tiempo (Slots)
Puedes manipular los bloques de forma muy intuitiva:

- **Creación por Arrastre (Draw to Create):** Haz clic en cualquier espacio vacío de la línea de tiempo y arrastra el ratón. Verás un bloque "fantasma" que se estira; al soltar, se creará automáticamente una nueva entrada de tiempo con esa duración exacta.
- **Mover (Drag & Drop):** Haz clic y arrastra el centro de cualquier bloque para cambiar su hora.
- **Redimensionar:** Haz clic y arrastra los bordes del bloque para ajustar su duración de 5 en 5 minutos.
- **Ajuste Inteligente (Doble Clic):** Haz doble clic sobre un bloque existente para que se estire automáticamente hasta tocar el bloque anterior o posterior, o los límites de tu jornada.
- **✏️ Editar:** Pulsa el icono del lápiz para abrir el modal y ajustar manualmente hora, duración y descripción.
- **❌ Eliminar:** Pulsa la "X" para borrar la entrada (siempre que el parte esté en borrador).

---

## 4. Leyenda Visual y Colores
El panel identifica rápidamente el tipo de tarea realizada:

- 🛠️ **Trabajo (Azul):** Horas de trabajo productivo.
- 📦 **Aprovisionamiento (Naranja):** Gestión de almacén.
- 🚗 **Viaje (Verde):** Tiempos de desplazamiento.
- 🍔 / ☕ **Descanso (Rojo):** Paradas.
- 🌴 **Vacaciones (Granate):** Días libres (Solo lectura).
- 🎉 **Festivo (Morado):** Días festivos registrados.
- **Extra (Azul Oscuro):** Horas extraordinarias.

> **Nota sobre Festivos y Descansos:** A diferencia de las vacaciones, los bloques de **Festivos** pueden editarse, redimensionarse o eliminarse mientras el parte esté en estado "Borrador".

---

## 5. Avisos Visuales
- **⚠️ Solapamiento (Fondo rojo a rayas):** Ocurre si dos bloques se pisan. Debes corregirlos para que no coincidan.
- **Fuera de Horario (Fondo transparente a rayas):** Indica que el bloque está registrado fuera de tu horario oficial (antes del inicio o después del fin de jornada).

---
## 🛠️ Instalación y Despliegue (Perfil Técnico)

### Requisitos previos
- [Node.js](https://nodejs.org/)
- [Power Platform CLI (pac)](https://docs.microsoft.com/en-us/powerapps/developer/data-platform/powerapps-cli)

### Construcción
1. Clona este repositorio.
2. Abre una terminal en la carpeta raíz del proyecto (`ParteDiarioPCF`).
3. Instala las dependencias:
```bash
   npm install
Compila el código para verificar que no hay errores (opcional):

Bash
   npm run build
Despliegue en Dataverse
Para publicar el componente directamente en el entorno de desarrollo:

Bash
# 1. Autenticación en el entorno (si es necesario)
pac auth create --url https://<entorno>[.crm4.dynamics.com/](https://.crm4.dynamics.com/)

# 2. Empaquetar y subir el componente usando el prefijo del publicador (ej: sec)
pac pcf push --publisher-prefix sec
---
###Configuración en el Formulario
Al añadir este componente a un formulario en Power Apps, se deben mapear las siguientes propiedades (definidas en el Manifest):

Campo Base / Estado (sec_estadoparte): Campo statuscode del parte diario.

Fecha del Parte (sec_fecha): Fecha a filtrar.

Hora Inicio / Fin Jornada (sec_horainicio, sec_horafin): Límites de la jornada laboral para los cálculos.

Recurso (sec_recursoid): El Recurso Reservable (Bookable Resource) asociado.

Orientación: Controla si inicia en vista Horizontal o Vertical.

Desarrollado para operaciones de Field Service.
