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

## 🛠️ Instalación y Despliegue

### Requisitos previos
- [Node.js](https://nodejs.org/)
- [Power Platform CLI (pac)](https://docs.microsoft.com/en-us/powerapps/developer/data-platform/powerapps-cli)

### Construcción
1. Clona este repositorio.
2. Abre una terminal en la carpeta raíz del proyecto (`ParteDiarioPCF`).
3. Instala las dependencias:
```bash
   npm install

   Despliegue en Dataverse
Para publicar el componente directamente en el entorno de desarrollo:

Bash
# 1. Autenticación en el entorno (si es necesario)
pac auth create --url https://<entorno>[.crm4.dynamics.com/](https://.crm4.dynamics.com/)

# 2. Empaquetar y subir el componente usando el prefijo del publicador (ej: sec)
pac pcf push --publisher-prefix sec
Configuración en el Formulario
Al añadir este componente a un formulario en Power Apps, se deben mapear las siguientes propiedades (definidas en el Manifest):

Campo Base / Estado (sec_estadoparte): Campo statuscode del parte diario.

Fecha del Parte (sec_fecha): Fecha a filtrar.

Hora Inicio / Fin Jornada (sec_horainicio, sec_horafin): Límites de la jornada laboral para los cálculos.

Recurso (sec_recursoid): El Recurso Reservable (Bookable Resource) asociado.

Orientación: Controla si inicia en vista Horizontal o Vertical.

📖 Manual de Uso
Este componente sustituye la clásica cuadrícula de registros por un panel interactivo. A continuación, se detalla el funcionamiento de cada elemento de la interfaz.

1. Panel de Métricas (Totales)
Ubicado en la parte superior derecha, muestra el balance de horas del día en tiempo real:

⏳ Imputado: Suma total de horas y minutos registrados en el parte.

Jornada: Las horas que dura el turno o jornada laboral de ese día.

Faltan / Extra: Calcula automáticamente la diferencia, mostrándose en rojo si faltan horas para completar la jornada, o en verde si hay horas extras o la jornada está completa.

2. Botones de Acción (Barra Superior)
↔️/↕️ Vista: Alterna la interfaz entre un formato de línea de tiempo vertical u horizontal según la preferencia del usuario.

🔍 Zoom: Permite cambiar la escala de visualización. Puedes ver las 24 horas del día completas o hacer "zoom" para encuadrar únicamente las horas de tu jornada laboral.

🔄 Refrescar: Recarga los datos directamente de Dataverse por si ha habido cambios externos.

🍔 Crear Almuerzo: Abre una ventana rápida para registrar un tiempo de comida. Por defecto sugiere las 14:00h y una duración de 1 hora, pero es modificable.

📦 Crear Aprovisionamiento: Abre una ventana para registrar tareas de carga/descarga de almacén. Sugiere por defecto las 08:00h y 30 minutos de duración.

Completar Huecos: Examina tu línea de tiempo y crea automáticamente bloques de horas en todos aquellos espacios vacíos dentro de tu jornada laboral para que el total cuadre al 100%.

🚀 Enviar Parte: Una vez que las horas están completas, este botón permite enviar el parte para su revisión. Atención: Al pulsarlo, el estado del registro cambiará a "Enviado" y el componente se bloqueará, pasando a modo de solo lectura.

3. Interacción con los Bloques de Tiempo (Slots)
Cada franja de color representa un bloque de tiempo registrado. Si el parte está en borrador, puedes interactuar con ellos de las siguientes formas:

Mover (Drag & Drop): Haz clic y mantén pulsado en el centro del bloque para arrastrarlo a una nueva hora.

Redimensionar (Ajustar duración): Pasa el ratón por el borde superior/inferior (o izquierdo/derecho en modo horizontal) de un bloque. Haz clic y arrastra el borde para acortar o extender la duración del bloque en intervalos de 5 minutos.

Doble Click (Expansión inteligente): Si haces doble click sobre un bloque, este se estirará automáticamente hacia arriba y hacia abajo para ocupar todo el tiempo libre disponible hasta chocar con el inicio/fin de jornada u otro bloque de tiempo existente.

✏️ Editar: Cada bloque tiene un pequeño icono de lápiz. Al pulsarlo, se abre un modal que te permite ajustar la hora exacta de inicio, elegir la duración en un desplegable y modificar la descripción de la tarea.

❌ Eliminar (x): Borra permanentemente el bloque de tiempo de la base de datos (pedirá confirmación previa).

4. Leyenda Visual y Colores
El componente utiliza un sistema de colores e iconos para identificar de un vistazo el tipo de tarea realizada:

🛠️ Trabajo (Azul): Horas estándar de trabajo productivo.

📦 Aprovisionamiento (Naranja): Tareas relacionadas con gestión de almacén o carga de material.

🚗 Viaje (Verde): Tiempos de desplazamiento.

🍔/☕ Descanso/Almuerzo (Rojo): Tiempos de parada.

🌴 Vacaciones/Ausencia (Granate): Días libres o permisos. (Nota: Estos bloques están protegidos y son de solo lectura, no se pueden modificar ni mover desde este panel).

Extra (Azul Oscuro): Horas catalogadas como extraordinarias.

Avisos visuales especiales:
⚠️ Solapamiento (Fondo rojo a rayas): Si dos bloques de tiempo comparten la misma hora (se pisan), se pintarán con un fondo rayado rojo y un icono de advertencia para indicar que debes corregir las horas.

Fuera de Horario (Fondo transparente a rayas): Si un bloque se registra antes de la hora de inicio de jornada o después de la hora de fin, tendrá un sombreado rayado para indicar que está ocurriendo fuera del horario oficial.

Desarrollado para operaciones de Field Service.