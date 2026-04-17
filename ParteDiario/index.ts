import { IInputs, IOutputs } from "./generated/ManifestTypes";

interface ITimeEntry extends ComponentFramework.WebApi.Entity {
    msdyn_timeentryid?: string;
    msdyn_start?: string;
    msdyn_end?: string;
    msdyn_type?: number;
    "_msdyn_workorder_value"?: string;
    "_msdyn_bookableresourcebooking_value"?: string;
    "msdyn_type@OData.Community.Display.V1.FormattedValue"?: string;
    "_msdyn_workorder_value@OData.Community.Display.V1.FormattedValue"?: string;
    "_msdyn_bookableresourcebooking_value@OData.Community.Display.V1.FormattedValue"?: string;
}

export class ParteDiario implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _version = "v1.0.20";
    private _palette = ["#0078d4", "#107c10", "#d83b01", "#5c2d91", "#008272", "#a80000", "#e3008c", "#ff8c00"];

    // Variables de estado para Drag & Drop
    private _timelineEl: HTMLDivElement;
    private _liveTooltip: HTMLDivElement;
    private _isVertical = false;
    private _isDragging = false;
    private _dragType: 'move' | 'left' | 'right' | null = null;
    private _dragTarget: HTMLElement | null = null;
    private _dragData = { id: "", originalStart: 0, originalEnd: 0, offsetDecimal: 0, newStart: 0, newEnd: 0, origStartDate: "", origEndDate: "" };

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._context = context;
        this._container = container;

        // Tooltip flotante para mostrar la hora en tiempo real al arrastrar
        this._liveTooltip = document.createElement("div");
        this._liveTooltip.className = "pd-live-tooltip";
        document.body.appendChild(this._liveTooltip);

        // Eventos globales para manejar el final del arrastre fuera del componente
        document.addEventListener("pointermove", this.onPointerMove.bind(this));
        document.addEventListener("pointerup", this.onPointerUp.bind(this));
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        this.renderTimeline();
    }

    private renderTimeline(): void {
        this._container.innerHTML = "";
        const context = this._context;

        try {
            const timelineWrapper = document.createElement("div");
            timelineWrapper.className = "pd-timeline-wrapper";

            this._timelineEl = document.createElement("div");
            this._timelineEl.className = "pd-timeline";
            timelineWrapper.appendChild(this._timelineEl);
            
            const params = context.parameters;
            const orientationRaw = params.orientacion?.raw as boolean | string | undefined;
            this._isVertical = orientationRaw === undefined || orientationRaw === null || orientationRaw === true || orientationRaw === 'vertical' || orientationRaw === 'Vertical';

            if (this._isVertical) {
                timelineWrapper.className = "pd-timeline-wrapper pd-vertical";
            }

            const axisDiv = document.createElement("div");
            axisDiv.className = this._isVertical ? "pd-axis pd-axis-vertical" : "pd-axis";
            for (let i = 0; i <= 24; i += 2) {
                const tick = document.createElement("div");
                tick.className = this._isVertical ? "pd-tick pd-tick-vertical" : "pd-tick";
                if (this._isVertical) {
                    tick.style.top = `${(i / 24) * 100}%`;
                } else {
                    tick.style.left = `${(i / 24) * 100}%`;
                }
                tick.innerText = `${i}:00`;
                axisDiv.appendChild(tick);
            }
            timelineWrapper.appendChild(axisDiv);
            this._container.appendChild(timelineWrapper);

            const versionEl = document.createElement("div");
            versionEl.className = "pd-version-tag";
            versionEl.innerText = this._version;
            timelineWrapper.appendChild(versionEl);

            const fechaRaw = params.sec_fecha?.raw;
            const inicioRaw = params.sec_horainicio?.raw;
            const finRaw = params.sec_horafin?.raw;
            const recursoRaw = params.sec_recursoid?.raw;

            let recursoId = "";
            if (Array.isArray(recursoRaw) && recursoRaw.length > 0 && recursoRaw[0] && recursoRaw[0].id) {
                recursoId = recursoRaw[0].id;
            } else if (typeof recursoRaw === "string") {
                recursoId = recursoRaw;
            }

            if (!fechaRaw || !inicioRaw || !finRaw || !recursoId) {
                const infoMsg = document.createElement("div");
                infoMsg.innerText = "Cargando componentes del formulario...";
                infoMsg.style.fontSize = "12px";
                infoMsg.style.color = "#666";
                infoMsg.style.marginTop = "10px";
                timelineWrapper.appendChild(infoMsg);
                return;
            }

            recursoId = recursoId.replace(/[{}]/g, "").toLowerCase();

            const hInicio = new Date(inicioRaw as string | number | Date);
            const hFin = new Date(finRaw as string | number | Date);

            const inicioDecimal = hInicio.getHours() + (hInicio.getMinutes() / 60);
            const finDecimal = hFin.getHours() + (hFin.getMinutes() / 60);

            const jornadaDiv = document.createElement("div");
            jornadaDiv.className = "pd-jornada";
            if (this._isVertical) {
                jornadaDiv.style.top = `${(inicioDecimal / 24) * 100}%`;
                jornadaDiv.style.height = `${((finDecimal - inicioDecimal) / 24) * 100}%`;
                jornadaDiv.style.width = `100%`;
            } else {
                jornadaDiv.style.left = `${(inicioDecimal / 24) * 100}%`;
                jornadaDiv.style.width = `${((finDecimal - inicioDecimal) / 24) * 100}%`;
            }
            jornadaDiv.title = `Jornada Laboral: ${this.formatTimeObj(hInicio)} - ${this.formatTimeObj(hFin)}`;
            this._timelineEl.appendChild(jornadaDiv);

            this.fetchTimeEntries(context, new Date(fechaRaw as string | number | Date), recursoId).then((entries) => {
                if (!entries || entries.length === 0) return;

                const typeColorMap = new Map<string, string>();
                let colorIndex = 0;

                entries.forEach((entry: ITimeEntry) => {
                    if (entry.msdyn_start && entry.msdyn_end && entry.msdyn_timeentryid) {
                        const startEntry = new Date(entry.msdyn_start);
                        const endEntry = new Date(entry.msdyn_end);

                        const startDec = startEntry.getHours() + (startEntry.getMinutes() / 60);
                        const endDec = endEntry.getHours() + (endEntry.getMinutes() / 60);

                        const isOutOfHours = (startDec < inicioDecimal) || (endDec > finDecimal);

                        const otName = entry["_msdyn_workorder_value@OData.Community.Display.V1.FormattedValue"] || "Ninguna";
                        const reservaName = entry["_msdyn_bookableresourcebooking_value@OData.Community.Display.V1.FormattedValue"] || "Ninguna";
                        const typeName = entry["msdyn_type@OData.Community.Display.V1.FormattedValue"] || (entry.msdyn_type ? `Tipo ${entry.msdyn_type}` : "General");

                        let entryColor = typeColorMap.get(typeName);
                        if (!entryColor) {
                            entryColor = this._palette[colorIndex % this._palette.length];
                            typeColorMap.set(typeName, entryColor);
                            colorIndex++;
                        }

                        const entryDiv = document.createElement("div");
                        entryDiv.className = isOutOfHours ? "pd-entry pd-out-of-hours" : "pd-entry pd-in-hours";
                        if (this._isVertical) {
                            entryDiv.style.top = `${(startDec / 24) * 100}%`;
                            entryDiv.style.height = `${((endDec - startDec) / 24) * 100}%`;
                            entryDiv.style.width = `100%`;
                        } else {
                            entryDiv.style.left = `${(startDec / 24) * 100}%`;
                            entryDiv.style.width = `${((endDec - startDec) / 24) * 100}%`;
                        }
                        entryDiv.style.backgroundColor = entryColor;
                        
                        // Guardar datos en el elemento para el drag and drop
                        entryDiv.dataset.id = entry.msdyn_timeentryid;
                        entryDiv.dataset.startDec = startDec.toString();
                        entryDiv.dataset.endDec = endDec.toString();
                        entryDiv.dataset.origStart = entry.msdyn_start;
                        entryDiv.dataset.origEnd = entry.msdyn_end;

                        const badgeText = `OT:${otName}-RES:${reservaName}`;
                        entryDiv.title = `[${this.formatTimeObj(startEntry)} - ${this.formatTimeObj(endEntry)}] ${typeName}\n${badgeText}`;

                        const badge = document.createElement("div");
                        badge.className = "pd-entry-badge";
                        badge.innerText = badgeText;
                        entryDiv.appendChild(badge);

                        // Tiradores de redimensión
                        const resizerLeft = document.createElement("div");
                        resizerLeft.className = "pd-resizer pd-resizer-left";
                        resizerLeft.dataset.action = "left";
                        
                        const resizerRight = document.createElement("div");
                        resizerRight.className = "pd-resizer pd-resizer-right";
                        resizerRight.dataset.action = "right";

                        entryDiv.appendChild(resizerLeft);
                        entryDiv.appendChild(resizerRight);

                        // Evento de inicio de arrastre
                        entryDiv.addEventListener("pointerdown", this.onPointerDown.bind(this));

                        this._timelineEl.appendChild(entryDiv);
                    }
                });

                this.arrangeEntryBadges();

                if (typeColorMap.size > 0) {
                    const legendDiv = document.createElement("div");
                    legendDiv.className = "pd-legend";
                    typeColorMap.forEach((color, typeName) => {
                        const item = document.createElement("div");
                        item.className = "pd-legend-item";
                        const colorBox = document.createElement("div");
                        colorBox.className = "pd-legend-color";
                        colorBox.style.backgroundColor = color;
                        const label = document.createElement("span");
                        label.innerText = typeName;
                        item.appendChild(colorBox);
                        item.appendChild(label);
                        legendDiv.appendChild(item);
                    });
                    timelineWrapper.appendChild(legendDiv);
                }

                return undefined;
            }).catch(err => console.error("Error WebAPI:", err));

        } catch (error: unknown) {
            console.error("Error en updateView:", error);
        }
    }

    // --- LÓGICA DE DRAG & DROP Y RESIZE ---

    private onPointerDown(e: PointerEvent): void {
        const target = e.target as HTMLElement;
        const entryDiv = target.closest('.pd-entry') as HTMLElement;
        if (!entryDiv || !this._timelineEl) return;

        e.preventDefault(); // Evita selecciones de texto raras
        this._isDragging = true;
        this._dragTarget = entryDiv;
        
        // Identificar si estamos arrastrando el centro (mover) o los bordes (redimensionar)
        if (target.classList.contains('pd-resizer-left')) {
            this._dragType = 'left';
        } else if (target.classList.contains('pd-resizer-right')) {
            this._dragType = 'right';
        } else {
            this._dragType = 'move';
        }

        // Leer los datos iniciales
        this._dragData.id = entryDiv.dataset.id || "";
        this._dragData.originalStart = parseFloat(entryDiv.dataset.startDec || "0");
        this._dragData.originalEnd = parseFloat(entryDiv.dataset.endDec || "0");
        this._dragData.newStart = this._dragData.originalStart;
        this._dragData.newEnd = this._dragData.originalEnd;
        this._dragData.origStartDate = entryDiv.dataset.origStart || "";
        this._dragData.origEndDate = entryDiv.dataset.origEnd || "";

        // Calcular el offset inicial para el 'move'
        const rect = this._timelineEl.getBoundingClientRect();
        const pointerDec = this._isVertical
            ? (e.clientY - rect.top) / rect.height * 24
            : (e.clientX - rect.left) / rect.width * 24;
        this._dragData.offsetDecimal = pointerDec - this._dragData.originalStart;

        entryDiv.classList.add("pd-dragging");
        this._liveTooltip.style.opacity = "1";
        this.updateLiveTooltip(e.clientX, e.clientY);
    }

    private onPointerMove(e: PointerEvent): void {
        if (!this._isDragging || !this._dragTarget || !this._timelineEl) return;

        const rect = this._timelineEl.getBoundingClientRect();
        const pointerDec = this._isVertical
            ? Math.max(0, Math.min(24, ((e.clientY - rect.top) / rect.height) * 24))
            : Math.max(0, Math.min(24, ((e.clientX - rect.left) / rect.width) * 24));

        // Calcular hora decimal basada en posición y hacer snap de 5 minutos
        const snappedDec = this.snapTo5Mins(pointerDec);

        const duration = this._dragData.originalEnd - this._dragData.originalStart;

        if (this._dragType === 'move') {
            let newStart = this.snapTo5Mins(pointerDec - this._dragData.offsetDecimal);
            
            // Límites de colisión
            if (newStart < 0) newStart = 0;
            if (newStart + duration > 24) newStart = 24 - duration;
            
            this._dragData.newStart = newStart;
            this._dragData.newEnd = newStart + duration;

        } else if (this._dragType === 'left') {
            // No dejar que el inicio sobrepase al final (con 5 mins de margen mínimo)
            this._dragData.newStart = Math.min(snappedDec, this._dragData.newEnd - (5/60));
        } else if (this._dragType === 'right') {
            // No dejar que el final sea menor al inicio
            this._dragData.newEnd = Math.max(snappedDec, this._dragData.newStart + (5/60));
        }

        // Actualizar visualmente la barra
        if (this._isVertical) {
            this._dragTarget.style.top = `${(this._dragData.newStart / 24) * 100}%`;
            this._dragTarget.style.height = `${((this._dragData.newEnd - this._dragData.newStart) / 24) * 100}%`;
        } else {
            this._dragTarget.style.left = `${(this._dragData.newStart / 24) * 100}%`;
            this._dragTarget.style.width = `${((this._dragData.newEnd - this._dragData.newStart) / 24) * 100}%`;
        }

        this.updateLiveTooltip(e.clientX, e.clientY);
    }

    private async onPointerUp(e: PointerEvent): Promise<void> {
        if (!this._isDragging) return;
        this._isDragging = false;
        this._liveTooltip.style.opacity = "0";

        if (this._dragTarget) {
            this._dragTarget.classList.remove("pd-dragging");
            
            // Comprobar si realmente ha habido un cambio
            const isChanged = (this._dragData.originalStart !== this._dragData.newStart) || 
                              (this._dragData.originalEnd !== this._dragData.newEnd);

            if (isChanged && this._dragData.id) {
                // Actualizar estilo a "cargando" para dar feedback
                this._dragTarget.style.opacity = "0.5";
                
                try {
                    // Generar las nuevas fechas ISO manteniendo el día original
                    const updatedStartISO = this.applyTimeToDateStr(this._dragData.origStartDate, this._dragData.newStart);
                    const updatedEndISO = this.applyTimeToDateStr(this._dragData.origEndDate, this._dragData.newEnd);

                    const payload = {
                        msdyn_start: updatedStartISO,
                        msdyn_end: updatedEndISO
                    };

                    await this._context.webAPI.updateRecord("msdyn_timeentry", this._dragData.id, payload);
                    
                    // Refrescar toda la vista tras guardar
                    this.renderTimeline();

                } catch (error) {
                    console.error("Error al actualizar la entrada de tiempo", error);
                    this.renderTimeline(); // Revertir visualmente si hay fallo
                }
            }
        }
        this._dragTarget = null;
        this._dragType = null;
    }

    // --- FUNCIONES AUXILIARES ---

    private updateLiveTooltip(x: number, y: number): void {
        const startStr = this.formatDecimalTime(this._dragData.newStart);
        const endStr = this.formatDecimalTime(this._dragData.newEnd);
        this._liveTooltip.innerText = `${startStr} - ${endStr}`;
        this._liveTooltip.style.left = `${x}px`;
        this._liveTooltip.style.top = `${y}px`;
    }

    private snapTo5Mins(decimalHours: number): number {
        const totalMinutes = Math.round(decimalHours * 60);
        const snappedMinutes = Math.round(totalMinutes / 5) * 5;
        return snappedMinutes / 60;
    }

    private applyTimeToDateStr(origIsoString: string, newDecimalHours: number): string {
        const dateObj = new Date(origIsoString);
        const hours = Math.floor(newDecimalHours);
        const mins = Math.round((newDecimalHours - hours) * 60);
        // Modificamos la hora local del objeto Date (mantiene el día correcto en nuestra zona)
        dateObj.setHours(hours, mins, 0, 0);
        // Devolvemos el ISO que Dataverse espera
        return dateObj.toISOString();
    }

    private formatDecimalTime(decimalHours: number): string {
        const h = Math.floor(decimalHours);
        const m = Math.round((decimalHours - h) * 60);
        return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
    }

    private formatTimeObj(date: Date): string {
        if (isNaN(date.getTime())) return "--:--";
        return `${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
    }

    private async fetchTimeEntries(context: ComponentFramework.Context<IInputs>, fecha: Date, recursoId: string): Promise<ITimeEntry[]> {
        const y = fecha.getFullYear();
        const m = String(fecha.getMonth() + 1).padStart(2, '0');
        const d = String(fecha.getDate()).padStart(2, '0');

        const fechaInicio = `${y}-${m}-${d}T00:00:00Z`;
        const fechaFin = `${y}-${m}-${d}T23:59:59Z`;

        const filter = `_msdyn_bookableresource_value eq ${recursoId} and msdyn_start ge ${fechaInicio} and msdyn_start le ${fechaFin}`;
        
        // Incluimos el msdyn_timeentryid para poder actualizarlo luego
        const query = `?$filter=${filter}&$select=msdyn_timeentryid,msdyn_start,msdyn_end,msdyn_type,_msdyn_workorder_value,_msdyn_bookableresourcebooking_value`;

        const response = await context.webAPI.retrieveMultipleRecords("msdyn_timeentry", query);
        return response.entities as ITimeEntry[];
    }

    private arrangeEntryBadges(): void {
        if (!this._timelineEl) return;

        const badges = Array.from(this._timelineEl.querySelectorAll<HTMLDivElement>(".pd-entry-badge"));
        const rowHeight = 18;
        const baseTop = -24;

        if (this._isVertical) {
            const columns: { top: number; bottom: number }[][] = [];
            const gap = 8;

            badges.forEach((badge) => {
                const entry = badge.parentElement as HTMLElement | null;
                if (!entry) return;

                const entryTop = entry.offsetTop;
                const entryHeight = entry.offsetHeight;
                const badgeHeight = badge.offsetHeight;
                const badgeTop = entryTop + (entryHeight / 2) - (badgeHeight / 2);
                const badgeBottom = badgeTop + badgeHeight;

                let colIndex = 0;
                while (colIndex < columns.length) {
                    const collision = columns[colIndex].some((placed) => !(badgeBottom <= placed.top || badgeTop >= placed.bottom));
                    if (!collision) break;
                    colIndex++;
                }

                if (colIndex === columns.length) {
                    columns.push([]);
                }

                columns[colIndex].push({ top: badgeTop, bottom: badgeBottom });
                badge.style.top = "50%";
                badge.style.left = "100%";
                badge.style.transform = "translateY(-50%)";
                badge.style.marginLeft = `${colIndex * (badge.offsetWidth + gap) + 6}px`;
            });
            return;
        }

        const rows: { left: number; right: number }[][] = [];

        badges.forEach((badge) => {
            const entry = badge.parentElement as HTMLElement | null;
            if (!entry) return;

            const entryLeft = entry.offsetLeft;
            const entryWidth = entry.offsetWidth;
            const badgeWidth = badge.offsetWidth;
            const badgeLeft = entryLeft + (entryWidth / 2) - (badgeWidth / 2);
            const badgeRight = badgeLeft + badgeWidth;

            let rowIndex = 0;
            while (rowIndex < rows.length) {
                const collision = rows[rowIndex].some((placed) => !(badgeRight <= placed.left || badgeLeft >= placed.right));
                if (!collision) break;
                rowIndex++;
            }

            if (rowIndex === rows.length) {
                rows.push([]);
            }

            rows[rowIndex].push({ left: badgeLeft, right: badgeRight });
            badge.style.top = `${baseTop - (rowIndex * rowHeight)}px`;
            badge.style.left = "50%";
            badge.style.transform = "translateX(-50%)";
            badge.style.marginLeft = "0";
        });
    }

    public getOutputs(): IOutputs { return {}; }
    public destroy(): void { 
        this._container.innerHTML = ""; 
        if (this._liveTooltip && this._liveTooltip.parentNode) {
            this._liveTooltip.parentNode.removeChild(this._liveTooltip);
        }
    }
}