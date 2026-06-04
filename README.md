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

---
# Manual de Uso - Panel Interactivo de Field Service ⏱️

Bienvenido al nuevo panel interactivo para la gestión del Parte Diario. Este componente sustituye la clásica cuadrícula de registros por una línea de tiempo visual, diseñada para agilizar y simplificar la imputación de horas de trabajo, descansos y aprovisionamientos.

A continuación, se detalla el funcionamiento de cada elemento de la interfaz.

---

## 1. Panel de Métricas (Totales)
Ubicado en la parte superior derecha de la pantalla, este panel te muestra el balance de tu día en tiempo real:
- **⏳ Imputado:** Es la suma total de horas y minutos que ya tienes registrados en tu parte de hoy.
- **Jornada:** Indica las horas totales que dura tu turno o jornada laboral teórica de ese día.
- **Faltan / Extra:** El sistema calcula automáticamente la diferencia. Se mostrará en **rojo** si aún te faltan horas para completar la jornada, o en **verde** si has completado la jornada o tienes horas extras.

---

## 2. Botones de Acción (Barra Superior)
En la parte superior encontrarás una barra de herramientas para realizar acciones rápidas:

- **↔️ / ↕️ Vista:** Cambia la interfaz entre un formato de línea de tiempo vertical u horizontal, según lo que te resulte más cómodo.
- **🔍 Zoom:** Permite cambiar la escala de visualización. Puedes ver las 24 horas del día completas o hacer "zoom" para encuadrar y centrarte únicamente en las horas de tu jornada laboral.
- **🔄 Refrescar:** Recarga los datos para asegurar que estás viendo la última información guardada.
- **🍔 Crear Almuerzo:** Abre una ventana rápida para registrar tu tiempo de comida. Por defecto sugiere las 14:00h y una duración de 1 hora (modificable).
- **📦 Crear Aprovisionamiento:** Abre una ventana para registrar tareas de almacén, carga o descarga. Sugiere por defecto las 08:00h y 30 minutos de duración.
- **Completar Huecos:** Examina tu línea de tiempo y crea automáticamente bloques de horas de trabajo en todos los espacios vacíos dentro de tu jornada. Úsalo para cuadrar tu día al 100% con un solo clic.
- **🚀 Enviar Parte:** Una vez que tus horas cuadran y el parte está listo, pulsa este botón para enviarlo. 
  > **⚠️ Atención:** Al enviar el parte, el estado cambiará a "Enviado" y el panel se bloqueará por completo (modo lectura), por lo que no podrás hacer más modificaciones.

---

## 3. Interacción con los Bloques de Tiempo (Slots)
Cada franja de color en tu línea de tiempo representa un bloque de horas. Mientras el parte esté en estado "Borrador", puedes manipularlos fácilmente con el ratón:

- **Mover (Drag & Drop):** Haz clic y mantén pulsado en el centro de un bloque para arrastrarlo a una nueva hora del día.
- **Redimensionar:** Pasa el ratón por el borde superior o inferior de un bloque. Haz clic y arrastra para acortar o alargar la duración (el ajuste va de 5 en 5 minutos).
- **Ajuste Inteligente (Doble Clic):** Si haces doble clic sobre un bloque existente, este se estirará automáticamente hacia arriba y hacia abajo para ocupar todo el tiempo libre disponible, deteniéndose justo donde empieza otro bloque o en el límite de tu jornada.
- **✏️ Editar:** Haz clic en el pequeño icono de lápiz dentro del bloque. Se abrirá una ventana para ajustar manualmente la hora exacta, la duración y la descripción de esa tarea.
- **❌ Eliminar:** Haz clic en la "X" del bloque para borrarlo permanentemente (el sistema te pedirá confirmación).

---

## 4. Leyenda Visual y Colores
El panel utiliza un código de colores e iconos para que identifiques de un vistazo en qué has invertido tu tiempo:

- 🛠️ **Trabajo (Azul):** Horas estándar de trabajo u órdenes de trabajo productivas.
- 📦 **Aprovisionamiento (Naranja):** Tareas relacionadas con la gestión de almacén o carga de material.
- 🚗 **Viaje (Verde):** Tiempos de desplazamiento entre ubicaciones.
- 🍔 / ☕ **Descanso o Almuerzo (Rojo):** Tiempos de parada.
- 🌴 **Vacaciones o Ausencias (Granate):** Días libres o permisos. 
  > *Nota: Estos bloques están protegidos por el sistema y son de solo lectura. No se pueden modificar ni mover desde este panel.*
- **Extra (Azul Oscuro):** Horas catalogadas oficialmente como extraordinarias.

---

## 5. Avisos Visuales Especiales
El panel te avisa de posibles errores en la imputación mediante tramas especiales en los bloques:

- **⚠️ Solapamiento (Fondo rojo a rayas):** Si dos bloques de tiempo se pisan en la misma hora, se pintarán con un fondo rayado rojo y un símbolo de advertencia. Debes corregir las horas para que no coincidan.
- **Fuera de Horario (Fondo transparente a rayas):** Si registras un bloque de tiempo antes de tu hora de inicio oficial o después de tu hora de fin, aparecerá con un sombreado rayado para indicarte que está fuera del horario estándar establecido.
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
