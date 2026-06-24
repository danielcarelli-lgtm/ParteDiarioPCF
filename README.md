# ParteDiarioPCF ⏱️ (v1.0.63)

Un componente de interfaz de usuario personalizado (PCF - Power Apps Component Framework) diseñado para Dynamics 365 y Power Platform. Este control transforma la gestión de horas (entidad `msdyn_timeentry`) ofreciendo una línea de tiempo visual e interactiva para los partes de trabajo diarios en el entorno de **Field Service**.

## 🚀 Características Principales

- **Línea de tiempo interactiva:** Visualiza las imputaciones de tiempo en un formato de línea de tiempo con soporte para orientación vertical y horizontal.
- **Drag & Drop:** Mueve o ajusta la duración de las entradas de tiempo simplemente arrastrando los bloques con el ratón.
- **Acciones Rápidas:** Botones integrados para crear automáticamente bloques comunes como Almuerzo y un menú unificado **"Crear entrada"** que permite elegir categorías (Control de stock, Visita sin OT, Acopio Material, Formación, Otros motivos laborales) estandarizadas como tipo "Trabajo".
- **Ajuste Inteligente (Doble Click):** Expande automáticamente un bloque de tiempo para cubrir los espacios vacíos adyacentes.
- **Validación Visual:** Detección automática de horas imputadas fuera de la jornada laboral y alertas visuales de solapamiento de horas (bloques cruzados).
- **Cálculo en Tiempo Real:** Muestra el total de horas imputadas frente a las horas objetivo de la jornada, indicando el tiempo faltante o extra.

---

# Manual de Uso - Panel Interactivo de Field Service ⏱️

Bienvenido al nuevo panel interactivo para la gestión del Parte Diario. Este componente sustituye la clásica cuadrícula de registros por una vista de calendario diaria mejorada.

### Acciones Disponibles:
1.  **Vista:** Puedes alternar entre vista vertical y horizontal, así como aplicar zoom para centrarte en tu jornada.
2.  **Crear Entrada:** Al presionar "Crear entrada", se desplegará un menú para clasificar tu tiempo bajo categorías específicas (Control de stock, Visita sin OT, Acopio Material, Formación, Otros motivos laborales).
3.  **Registro de Almuerzo:** Botón dedicado para registrar pausas de almuerzo rápidamente.
4.  **Completar Huecos:** Herramienta inteligente que rellena automáticamente los espacios libres de tu jornada con entradas de trabajo o festivos.
5.  **Enviar Parte:** Botón para consolidar el parte cuando la jornada esté completa (bloquea ediciones futuras).

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