import { IInputs, IOutputs } from "./generated/ManifestTypes";

interface ITimeEntry extends ComponentFramework.WebApi.Entity {
    msdyn_timeentryid?: string;
    msdyn_start?: string;
    msdyn_end?: string;
    msdyn_type?: number;
    "_msdyn_workorder_value"?: string;
    "msdyn_type@OData.Community.Display.V1.FormattedValue"?: string;
    "_msdyn_workorder_value@OData.Community.Display.V1.FormattedValue"?: string;
}

export class ParteDiario implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _version = "v1.0.34";
    private _palette = ["#0078d4", "#107c10", "#d83b01", "#5c2d91", "#008272", "#a80000", "#e3008c", "#ff8c00"];

    private _timelineEl: HTMLDivElement;
    private _liveTooltip: HTMLDivElement;
    private _isVertical = false;
    private _isDragging = false;
    private _lastRenderId = 0;
    private _dragType: 'move' | 'left' | 'right' | null = null;
    private _dragTarget: HTMLElement | null = null;
    private _dragData = { id: "", originalStart: 0, originalEnd: 0, offsetDecimal: 0, newStart: 0, newEnd: 0, origStartDate: "", origEndDate: "" };

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._context = context;
        this._container = container;
        
        this._container.style.position = "relative";
        this._container.style.width = "100%";
        this._container.style.height = "100%";

        const orientationRaw = context.parameters.orientacion?.raw;
        this._isVertical = orientationRaw === true || orientationRaw === null;

        this._liveTooltip = document.createElement("div");
        this._liveTooltip.className = "pd-live-tooltip";
        document.body.appendChild(this._liveTooltip);

        document.addEventListener("pointermove", this.onPointerMove.bind(this));
        document.addEventListener("pointerup", this.onPointerUp.bind(this));
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        this.renderTimeline().catch(err => console.error(err));
    }

    private async renderTimeline(): Promise<void> {
        const renderId = ++this._lastRenderId;
        this._container.innerHTML = "";
        
        const toolbar = document.createElement("div");
        toolbar.className = "pd-toolbar";
        
        const toggleBtn = document.createElement("button");
        toggleBtn.className = "pd-btn pd-btn-primary";
        toggleBtn.innerText = this._isVertical ? "↔️ Vista Horizontal" : "↕️ Vista Vertical";
        toggleBtn.onclick = () => {
            this._isVertical = !this._isVertical;
            this.renderTimeline().catch(err => console.error(err));
        };
        toolbar.appendChild(toggleBtn);

        const fillBtn = document.createElement("button");
        fillBtn.className = "pd-btn pd-btn-secondary";
        fillBtn.innerText = "Completar Huecos";
        fillBtn.onclick = () => this.fillGaps(fillBtn).catch(err => console.error(err));
        toolbar.appendChild(fillBtn);

        this._container.appendChild(toolbar);

        const timelineWrapper = document.createElement("div");
        timelineWrapper.className = this._isVertical ? "pd-timeline-wrapper pd-vertical" : "pd-timeline-wrapper pd-horizontal";

        this._timelineEl = document.createElement("div");
        this._timelineEl.className = this._isVertical ? "pd-timeline pd-vertical" : "pd-timeline pd-horizontal";
        
        const axisDiv = document.createElement("div");
        axisDiv.className = this._isVertical ? "pd-axis pd-axis-vertical" : "pd-axis pd-axis-horizontal";
        
        for (let i = 0; i <= 24; i += 2) {
            const tick = document.createElement("div");
            tick.className = this._isVertical ? "pd-tick pd-tick-vertical" : "pd-tick pd-tick-horizontal";
            if (this._isVertical) tick.style.top = `${(i / 24) * 100}%`;
            else tick.style.left = `${(i / 24) * 100}%`;
            tick.innerText = `${i}:00`;
            axisDiv.appendChild(tick);
        }

        if (this._isVertical) {
            timelineWrapper.appendChild(axisDiv);
            timelineWrapper.appendChild(this._timelineEl);
        } else {
            timelineWrapper.appendChild(this._timelineEl);
            timelineWrapper.appendChild(axisDiv);
        }
        
        this._container.appendChild(timelineWrapper);

        const params = this._context.parameters;
        const fechaRaw = params.sec_fecha?.raw;
        const inicioRaw = params.sec_horainicio?.raw;
        const finRaw = params.sec_horafin?.raw;
        const recursoRaw = params.sec_recursoid?.raw;

        let recursoId = "";
        if (Array.isArray(recursoRaw) && recursoRaw.length > 0) recursoId = recursoRaw[0].id;

        if (!fechaRaw || !inicioRaw || !finRaw || !recursoId) return;

        recursoId = recursoId.replace(/[{}]/g, "").toLowerCase();
        const hInicio = new Date(inicioRaw as Date);
        const hFin = new Date(finRaw as Date);
        const inicioDecimal = hInicio.getHours() + (hInicio.getMinutes() / 60);
        const finDecimal = hFin.getHours() + (hFin.getMinutes() / 60);

        const jornadaDiv = document.createElement("div");
        jornadaDiv.className = "pd-jornada";
        if (this._isVertical) {
            jornadaDiv.style.top = `${(inicioDecimal / 24) * 100}%`;
            jornadaDiv.style.height = `${((finDecimal - inicioDecimal) / 24) * 100}%`;
        } else {
            jornadaDiv.style.left = `${(inicioDecimal / 24) * 100}%`;
            jornadaDiv.style.width = `${((finDecimal - inicioDecimal) / 24) * 100}%`;
        }
        this._timelineEl.appendChild(jornadaDiv);

        try {
            const entries = await this.fetchTimeEntries(this._context, fechaRaw as Date, recursoId);
            
            if (renderId !== this._lastRenderId) return;

            const typeColorMap = new Map<string, string>();
            let colorIndex = 0;

            entries.forEach((entry: ITimeEntry) => {
                if (entry.msdyn_start && entry.msdyn_end) {
                    const startEntry = new Date(entry.msdyn_start);
                    const endEntry = new Date(entry.msdyn_end);
                    const startDec = startEntry.getHours() + (startEntry.getMinutes() / 60);
                    const endDec = endEntry.getHours() + (endEntry.getMinutes() / 60);

                    const isOutOfHours = (startDec < inicioDecimal - 0.01) || (endDec > finDecimal + 0.01);
                    const typeName = entry["msdyn_type@OData.Community.Display.V1.FormattedValue"] || "General";

                    let entryColor = typeColorMap.get(typeName);
                    if (!entryColor) {
                        entryColor = this._palette[colorIndex % this._palette.length];
                        typeColorMap.set(typeName, entryColor);
                        colorIndex++;
                    }

                    const entryDiv = document.createElement("div");
                    entryDiv.className = isOutOfHours ? "pd-entry pd-out-of-hours" : "pd-entry";
                    
                    if (this._isVertical) {
                        entryDiv.style.top = `${(startDec / 24) * 100}%`;
                        entryDiv.style.height = `${((endDec - startDec) / 24) * 100}%`;
                        entryDiv.style.width = `100%`;
                        const innerTime = document.createElement("div");
                        innerTime.className = "pd-entry-inner-time";
                        innerTime.innerText = `${this.formatTimeObj(startEntry)}-${this.formatTimeObj(endEntry)}`;
                        entryDiv.appendChild(innerTime);
                    } else {
                        entryDiv.style.left = `${(startDec / 24) * 100}%`;
                        entryDiv.style.width = `${((endDec - startDec) / 24) * 100}%`;
                        entryDiv.classList.add("pd-entry-horizontal");
                    }
                    
                    entryDiv.style.backgroundColor = entryColor;
                    entryDiv.dataset.id = entry.msdyn_timeentryid;
                    entryDiv.dataset.startDec = startDec.toString();
                    entryDiv.dataset.endDec = endDec.toString();
                    entryDiv.dataset.origStart = entry.msdyn_start;
                    entryDiv.dataset.origEnd = entry.msdyn_end;

                    const otName = entry["_msdyn_workorder_value@OData.Community.Display.V1.FormattedValue"] || "N/A";
                    const badge = document.createElement("div");
                    badge.className = "pd-entry-badge";
                    badge.innerText = `OT:${otName}`;
                    entryDiv.appendChild(badge);

                    const resizerLeft = document.createElement("div");
                    resizerLeft.className = "pd-resizer pd-resizer-left";
                    const resizerRight = document.createElement("div");
                    resizerRight.className = "pd-resizer pd-resizer-right";

                    entryDiv.appendChild(resizerLeft);
                    entryDiv.appendChild(resizerRight);
                    entryDiv.addEventListener("pointerdown", this.onPointerDown.bind(this));

                    this._timelineEl.appendChild(entryDiv);
                }
            });

            this.arrangeEntryBadges();

            if (typeColorMap.size > 0) {
                const legendDiv = document.createElement("div");
                legendDiv.className = "pd-legend";
                typeColorMap.forEach((color, name) => {
                    const item = document.createElement("div");
                    item.className = "pd-legend-item";
                    const box = document.createElement("div");
                    box.className = "pd-legend-color";
                    box.style.backgroundColor = color;
                    const span = document.createElement("span");
                    span.innerText = name;
                    item.appendChild(box);
                    item.appendChild(span);
                    legendDiv.appendChild(item);
                });
                this._container.appendChild(legendDiv);
            }

            const versionEl = document.createElement("div");
            versionEl.className = "pd-version-tag";
            versionEl.innerText = this._version;
            this._container.appendChild(versionEl);

        } catch (err) {
            console.error("Error WebAPI:", err);
        }
    }

    private async fillGaps(btn: HTMLButtonElement): Promise<void> {
        try {
            if (btn.disabled) return;
            btn.disabled = true;
            btn.innerText = "Procesando...";
            this.showLoadingOverlay();

            const params = this._context.parameters;
            const fechaRaw = params.sec_fecha?.raw;
            const inicioRaw = params.sec_horainicio?.raw;
            const finRaw = params.sec_horafin?.raw;
            const recursoRaw = params.sec_recursoid?.raw;

            if (!fechaRaw || !inicioRaw || !finRaw || !recursoRaw) {
                throw new Error("Faltan parámetros del formulario para completar los huecos.");
            }

            let recursoId = "";
            if (Array.isArray(recursoRaw) && recursoRaw.length > 0) recursoId = recursoRaw[0].id;
            recursoId = recursoId.replace(/[{}]/g, "").toLowerCase();

            const hInicio = new Date(inicioRaw as Date);
            const hFin = new Date(finRaw as Date);
            const inicioDec = hInicio.getHours() + (hInicio.getMinutes() / 60);
            const finDec = hFin.getHours() + (hFin.getMinutes() / 60);

            const entries = await this.fetchTimeEntries(this._context, fechaRaw as Date, recursoId);
            
            const sorted = entries
                .map(e => ({ 
                    start: new Date(e.msdyn_start!).getHours() + (new Date(e.msdyn_start!).getMinutes() / 60), 
                    end: new Date(e.msdyn_end!).getHours() + (new Date(e.msdyn_end!).getMinutes() / 60) 
                }))
                .sort((a, b) => a.start - b.start);

            const gaps: { s: number, e: number }[] = [];
            let currentPos = inicioDec;

            for (const entry of sorted) {
                if (entry.start > currentPos + 0.02) {
                    gaps.push({ s: currentPos, e: entry.start });
                }
                currentPos = Math.max(currentPos, entry.end);
            }

            if (currentPos < finDec - 0.02) {
                gaps.push({ s: currentPos, e: finDec });
            }

            if (gaps.length > 0) {
                const baseDate = new Date(fechaRaw as Date);
                for (const gap of gaps) {
                    // Cálculo de duración en minutos
                    const durationMinutes = Math.round((gap.e - gap.s) * 60);
                    
                    const data: ComponentFramework.WebApi.Entity = {
                        "msdyn_start": this.applyTimeToDateStr(baseDate.toISOString(), gap.s),
                        "msdyn_end": this.applyTimeToDateStr(baseDate.toISOString(), gap.e),
                        "msdyn_duration": durationMinutes,
                        "msdyn_type": 192355000, 
                        "msdyn_bookableresource@odata.bind": `/bookableresources(${recursoId})`
                    };
                    await this._context.webAPI.createRecord("msdyn_timeentry", data);
                }
            }

            await this.renderTimeline();

        } catch (error: unknown) {
            console.error("Error al completar huecos:", error);
            this.hideLoadingOverlay();
            btn.disabled = false;
            btn.innerText = "Completar Huecos";
            this.showErrorModal(error);
        }
    }

    private showLoadingOverlay(): void {
        let overlay = this._container.querySelector('.pd-loading-overlay');
        if (!overlay) {
            overlay = document.createElement('div');
            overlay.className = 'pd-loading-overlay';
            const spinner = document.createElement('div');
            spinner.className = 'pd-spinner';
            overlay.appendChild(spinner);
            this._container.appendChild(overlay);
        }
    }

    private hideLoadingOverlay(): void {
        const overlay = this._container.querySelector('.pd-loading-overlay');
        if (overlay && overlay.parentNode) {
            overlay.parentNode.removeChild(overlay);
        }
    }

    private showErrorModal(error: unknown): void {
        const backdrop = document.createElement("div");
        backdrop.className = "pd-modal-backdrop";

        const modal = document.createElement("div");
        modal.className = "pd-modal";

        const title = document.createElement("h3");
        title.innerText = "⚠️ Error en la operación";
        modal.appendChild(title);

        const text = document.createElement("p");
        text.innerText = "Ha ocurrido un problema. Puedes copiar el texto inferior para reportarlo:";
        modal.appendChild(text);

        const textarea = document.createElement("textarea");
        textarea.className = "pd-modal-textarea";
        textarea.readOnly = true;
        
        let errMsg = "Error desconocido";
        if (error instanceof Error) {
            errMsg = error.message;
        } else if (typeof error === "string") {
            errMsg = error;
        } else if (error && typeof error === "object") {
            const errObj = error as Record<string, unknown>;
            if (typeof errObj.message === "string") {
                errMsg = errObj.message;
            } else {
                errMsg = JSON.stringify(error, null, 2);
            }
        }
        textarea.value = errMsg;
        modal.appendChild(textarea);

        const closeBtn = document.createElement("button");
        closeBtn.className = "pd-btn pd-btn-primary";
        closeBtn.innerText = "Cerrar";
        closeBtn.onclick = () => {
            if (backdrop.parentNode) backdrop.parentNode.removeChild(backdrop);
        };
        modal.appendChild(closeBtn);

        backdrop.appendChild(modal);
        this._container.appendChild(backdrop);
    }

    private onPointerDown(e: PointerEvent): void {
        const target = e.target as HTMLElement;
        const entryDiv = target.closest('.pd-entry') as HTMLElement;
        if (!entryDiv) return;

        this._isDragging = true;
        this._dragTarget = entryDiv;
        this._dragType = target.classList.contains('pd-resizer-left') ? 'left' : (target.classList.contains('pd-resizer-right') ? 'right' : 'move');

        this._dragData.id = entryDiv.dataset.id || "";
        this._dragData.originalStart = parseFloat(entryDiv.dataset.startDec || "0");
        this._dragData.originalEnd = parseFloat(entryDiv.dataset.endDec || "0");
        this._dragData.newStart = this._dragData.originalStart;
        this._dragData.newEnd = this._dragData.originalEnd;
        this._dragData.origStartDate = entryDiv.dataset.origStart || "";
        this._dragData.origEndDate = entryDiv.dataset.origEnd || "";

        const rect = this._timelineEl.getBoundingClientRect();
        const pointerDec = this._isVertical ? (e.clientY - rect.top) / rect.height * 24 : (e.clientX - rect.left) / rect.width * 24;
        this._dragData.offsetDecimal = pointerDec - this._dragData.originalStart;

        entryDiv.classList.add("pd-dragging");
        this._liveTooltip.style.opacity = "1";
    }

    private onPointerMove(e: PointerEvent): void {
        if (!this._isDragging || !this._dragTarget || !this._timelineEl) return;

        const rect = this._timelineEl.getBoundingClientRect();
        const pointerDec = this._isVertical ? Math.max(0, Math.min(24, ((e.clientY - rect.top) / rect.height) * 24)) : Math.max(0, Math.min(24, ((e.clientX - rect.left) / rect.width) * 24));
        const snappedDec = Math.round(pointerDec * 12) / 12;

        if (this._dragType === 'move') {
            const duration = this._dragData.originalEnd - this._dragData.originalStart;
            let newStart = Math.round((pointerDec - this._dragData.offsetDecimal) * 12) / 12;
            newStart = Math.max(0, Math.min(24 - duration, newStart));
            this._dragData.newStart = newStart;
            this._dragData.newEnd = newStart + duration;
        } else if (this._dragType === 'left') {
            this._dragData.newStart = Math.min(snappedDec, this._dragData.newEnd - (5/60));
        } else if (this._dragType === 'right') {
            this._dragData.newEnd = Math.max(snappedDec, this._dragData.newStart + (5/60));
        }

        if (this._isVertical) {
            this._dragTarget.style.top = `${(this._dragData.newStart / 24) * 100}%`;
            this._dragTarget.style.height = `${((this._dragData.newEnd - this._dragData.newStart) / 24) * 100}%`;
            const inner = this._dragTarget.querySelector('.pd-entry-inner-time') as HTMLElement;
            if(inner) inner.innerText = `${this.formatDecimalTime(this._dragData.newStart)}-${this.formatDecimalTime(this._dragData.newEnd)}`;
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
            if (Math.abs(this._dragData.originalStart - this._dragData.newStart) > 0.01 || Math.abs(this._dragData.originalEnd - this._dragData.newEnd) > 0.01) {
                this._dragTarget.style.opacity = "0.5";
                const payload = {
                    msdyn_start: this.applyTimeToDateStr(this._dragData.origStartDate, this._dragData.newStart),
                    msdyn_end: this.applyTimeToDateStr(this._dragData.origEndDate, this._dragData.newEnd)
                };
                
                try {
                    this.showLoadingOverlay();
                    await this._context.webAPI.updateRecord("msdyn_timeentry", this._dragData.id, payload);
                    await this.renderTimeline();
                } catch (error: unknown) {
                    this.hideLoadingOverlay();
                    this.showErrorModal(error);
                    this._dragTarget.style.opacity = "1";
                    await this.renderTimeline(); 
                }
            }
        }
        this._dragTarget = null;
    }

    private updateLiveTooltip(x: number, y: number): void {
        this._liveTooltip.innerText = `${this.formatDecimalTime(this._dragData.newStart)} - ${this.formatDecimalTime(this._dragData.newEnd)}`;
        this._liveTooltip.style.left = `${x}px`;
        this._liveTooltip.style.top = `${y}px`;
    }

    private applyTimeToDateStr(origIsoString: string, decimalHours: number): string {
        const d = new Date(origIsoString);
        const h = Math.floor(decimalHours);
        const m = Math.round((decimalHours - h) * 60);
        d.setHours(h, m, 0, 0);
        return d.toISOString();
    }

    private formatDecimalTime(decimalHours: number): string {
        const h = Math.floor(decimalHours);
        const m = Math.round((decimalHours - h) * 60);
        return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
    }

    private formatTimeObj(date: Date): string {
        return `${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
    }

    private async fetchTimeEntries(context: ComponentFramework.Context<IInputs>, fecha: Date, recursoId: string): Promise<ITimeEntry[]> {
        const y = fecha.getFullYear();
        const m = String(fecha.getMonth() + 1).padStart(2, '0');
        const d = String(fecha.getDate()).padStart(2, '0');
        const filter = `_msdyn_bookableresource_value eq ${recursoId} and msdyn_start ge ${y}-${m}-${d}T00:00:00Z and msdyn_start le ${y}-${m}-${d}T23:59:59Z`;
        const query = `?$filter=${filter}&$select=msdyn_timeentryid,msdyn_start,msdyn_end,msdyn_type,_msdyn_workorder_value`;
        const res = await context.webAPI.retrieveMultipleRecords("msdyn_timeentry", query);
        return res.entities as ITimeEntry[];
    }

    private arrangeEntryBadges(): void {
        if (!this._timelineEl) return;
        const badges = Array.from(this._timelineEl.querySelectorAll<HTMLDivElement>(".pd-entry-badge"));
        if (this._isVertical) {
            const columns: { top: number; bottom: number }[][] = [];
            badges.forEach((badge) => {
                const entry = badge.parentElement as HTMLElement;
                const bTop = entry.offsetTop;
                const bBottom = bTop + entry.offsetHeight;
                let col = 0;
                while (col < columns.length) {
                    if (!columns[col].some(p => !(bBottom <= p.top || bTop >= p.bottom))) break;
                    col++;
                }
                if (col === columns.length) columns.push([]);
                columns[col].push({ top: bTop, bottom: bBottom });
                badge.style.marginLeft = `${col * 110 + 6}px`;
            });
        } else {
            const rows: { left: number; right: number }[][] = [];
            badges.forEach((badge) => {
                const entry = badge.parentElement as HTMLElement;
                const bLeft = entry.offsetLeft;
                const bRight = bLeft + entry.offsetWidth;
                let row = 0;
                while (row < rows.length) {
                    if (!rows[row].some(p => !(bRight <= p.left || bLeft >= p.right))) break;
                    row++;
                }
                if (row === rows.length) rows.push([]);
                rows[row].push({ left: bLeft, right: bRight });
                badge.style.top = `${-24 - (row * 18)}px`;
            });
        }
    }

    public getOutputs(): IOutputs { return {}; }
    public destroy(): void { 
        if (this._liveTooltip?.parentNode) this._liveTooltip.parentNode.removeChild(this._liveTooltip);
    }
}