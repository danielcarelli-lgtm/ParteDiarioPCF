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
    private _version = "v1.0.38";
    private _palette = ["#0078d4", "#107c10", "#d83b01", "#5c2d91", "#008272", "#a80000", "#e3008c", "#ff8c00"];

    private _timelineEl: HTMLDivElement;
    private _liveTooltip: HTMLDivElement;
    private _isVertical = false;
    private _isDragging = false;
    private _isZoomed = false;
    private _minViewHour = 0;
    private _maxViewHour = 24;
    private _lastRenderId = 0;
    
    private _currentEntries: { id: string, startDec: number, endDec: number }[] = [];

    private _dragType: 'move' | 'left' | 'right' | null = null;
    private _dragTarget: HTMLElement | null = null;
    private _dragData = { id: "", originalStart: 0, originalEnd: 0, offsetDecimal: 0, newStart: 0, newEnd: 0, origStartDate: "", origEndDate: "", minBound: 0, maxBound: 24 };

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
        this._currentEntries = []; 
        
        const params = this._context.parameters;
        const fechaRaw = params.sec_fecha?.raw;
        const inicioRaw = params.sec_horainicio?.raw;
        const finRaw = params.sec_horafin?.raw;
        const recursoRaw = params.sec_recursoid?.raw;

        let recursoId = "";
        if (Array.isArray(recursoRaw) && recursoRaw.length > 0) recursoId = recursoRaw[0].id;

        // Extraer horas estrictamente en UTC para evitar desfases de Zona Horaria
        const inicioDecimal = this.getTimeDecFromDate(inicioRaw as Date);
        const finDecimal = this.getTimeDecFromDate(finRaw as Date);

        if (this._isZoomed && inicioRaw && finRaw) {
            this._minViewHour = Math.max(0, inicioDecimal - 0.5);
            this._maxViewHour = Math.min(24, finDecimal + 0.5);
            if (this._maxViewHour <= this._minViewHour) this._maxViewHour = this._minViewHour + 1;
        } else {
            this._minViewHour = 0;
            this._maxViewHour = 24;
        }

        const range = this._maxViewHour - this._minViewHour;

        const toolbar = document.createElement("div");
        toolbar.className = "pd-toolbar";
        
        const actionsDiv = document.createElement("div");
        actionsDiv.className = "pd-toolbar-actions";

        const toggleBtn = document.createElement("button");
        toggleBtn.className = "pd-btn pd-btn-secondary";
        toggleBtn.innerText = this._isVertical ? "↔️ Vista Horizontal" : "↕️ Vista Vertical";
        toggleBtn.onclick = () => { this._isVertical = !this._isVertical; this.renderTimeline(); };
        actionsDiv.appendChild(toggleBtn);

        const zoomBtn = document.createElement("button");
        zoomBtn.className = "pd-btn pd-btn-secondary";
        zoomBtn.innerText = this._isZoomed ? "🔍 Ver 24h" : "🔍 Zoom Jornada";
        zoomBtn.onclick = () => { this._isZoomed = !this._isZoomed; this.renderTimeline(); };
        actionsDiv.appendChild(zoomBtn);

        const refreshBtn = document.createElement("button");
        refreshBtn.className = "pd-btn pd-btn-secondary";
        refreshBtn.innerText = "🔄 Refrescar";
        refreshBtn.onclick = () => { this.showLoadingOverlay(); this.renderTimeline(); };
        actionsDiv.appendChild(refreshBtn);

        const lunchBtn = document.createElement("button");
        lunchBtn.className = "pd-btn pd-btn-primary";
        lunchBtn.innerText = "🍔 Crear Almuerzo";
        lunchBtn.onclick = () => this.showLunchModal();
        actionsDiv.appendChild(lunchBtn);

        const fillBtn = document.createElement("button");
        fillBtn.className = "pd-btn pd-btn-primary";
        fillBtn.innerText = "Completar Huecos";
        fillBtn.onclick = () => this.fillGaps(fillBtn).catch(err => console.error(err));
        actionsDiv.appendChild(fillBtn);

        toolbar.appendChild(actionsDiv);

        const totalsDiv = document.createElement("div");
        totalsDiv.className = "pd-totals";
        toolbar.appendChild(totalsDiv);

        this._container.appendChild(toolbar);

        const timelineWrapper = document.createElement("div");
        timelineWrapper.className = this._isVertical ? "pd-timeline-wrapper pd-vertical" : "pd-timeline-wrapper pd-horizontal";

        this._timelineEl = document.createElement("div");
        this._timelineEl.className = this._isVertical ? "pd-timeline pd-vertical" : "pd-timeline pd-horizontal";
        
        const axisDiv = document.createElement("div");
        axisDiv.className = this._isVertical ? "pd-axis pd-axis-vertical" : "pd-axis pd-axis-horizontal";
        
        const tickStep = this._isZoomed ? 1 : 2;
        const startTick = Math.ceil(this._minViewHour);
        for (let i = startTick; i <= this._maxViewHour; i += tickStep) {
            const tick = document.createElement("div");
            tick.className = this._isVertical ? "pd-tick pd-tick-vertical" : "pd-tick pd-tick-horizontal";
            const percent = ((i - this._minViewHour) / range) * 100;
            if (this._isVertical) tick.style.top = `${percent}%`;
            else tick.style.left = `${percent}%`;
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

        if (!fechaRaw || !inicioRaw || !finRaw || !recursoId) return;

        recursoId = recursoId.replace(/[{}]/g, "").toLowerCase();

        const jornadaDiv = document.createElement("div");
        jornadaDiv.className = "pd-jornada";
        const startJornadaPercent = ((inicioDecimal - this._minViewHour) / range) * 100;
        const sizeJornadaPercent = ((finDecimal - inicioDecimal) / range) * 100;
        
        if (this._isVertical) {
            jornadaDiv.style.top = `${startJornadaPercent}%`;
            jornadaDiv.style.height = `${sizeJornadaPercent}%`;
        } else {
            jornadaDiv.style.left = `${startJornadaPercent}%`;
            jornadaDiv.style.width = `${sizeJornadaPercent}%`;
        }
        this._timelineEl.appendChild(jornadaDiv);

        try {
            const entries = await this.fetchTimeEntries(this._context, fechaRaw as Date, recursoId);
            
            if (renderId !== this._lastRenderId) return;

            const typeColorMap = new Map<string, string>();
            let colorIndex = 0;
            let totalMinsLogged = 0;

            entries.forEach((entry: ITimeEntry) => {
                if (entry.msdyn_start && entry.msdyn_end) {
                    const startEntry = new Date(entry.msdyn_start);
                    const endEntry = new Date(entry.msdyn_end);
                    
                    const startDec = this.getTimeDecFromDate(startEntry);
                    const endDec = this.getTimeDecFromDate(endEntry);

                    this._currentEntries.push({ id: entry.msdyn_timeentryid!, startDec, endDec });
                    totalMinsLogged += (endDec - startDec) * 60;

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

                    // Si está haciendo zoom y la entrada está fuera de la vista, la ocultamos
                    if (this._isZoomed && (endDec <= this._minViewHour || startDec >= this._maxViewHour)) {
                        entryDiv.style.display = "none";
                    }
                    
                    const startPercent = ((startDec - this._minViewHour) / range) * 100;
                    const sizePercent = ((endDec - startDec) / range) * 100;

                    if (this._isVertical) {
                        entryDiv.style.top = `${startPercent}%`;
                        entryDiv.style.height = `${sizePercent}%`;
                        entryDiv.style.width = `100%`;
                        const innerTime = document.createElement("div");
                        innerTime.className = "pd-entry-inner-time";
                        innerTime.innerText = `${this.formatTimeObj(startEntry)}-${this.formatTimeObj(endEntry)}`;
                        entryDiv.appendChild(innerTime);
                    } else {
                        entryDiv.style.left = `${startPercent}%`;
                        entryDiv.style.width = `${sizePercent}%`;
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

                    if (entry.msdyn_type === 192355000) {
                        const deleteBtn = document.createElement("div");
                        deleteBtn.className = "pd-delete-btn";
                        deleteBtn.innerHTML = "&times;";
                        deleteBtn.title = "Eliminar entrada";
                        deleteBtn.addEventListener("pointerdown", async (ev) => {
                            ev.stopPropagation(); 
                            if(confirm("¿Estás seguro de que quieres eliminar esta entrada de descanso?")) {
                                try {
                                    this.showLoadingOverlay();
                                    await this._context.webAPI.deleteRecord("msdyn_timeentry", entry.msdyn_timeentryid!);
                                    await this.renderTimeline();
                                } catch (error) {
                                    this.hideLoadingOverlay();
                                    this.showErrorModal(error);
                                }
                            }
                        });
                        entryDiv.appendChild(deleteBtn);
                    }

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

            const workMins = Math.round((finDecimal - inicioDecimal) * 60);
            const loggedMins = Math.round(totalMinsLogged);
            const missingMins = workMins - loggedMins;

            const hT = Math.floor(loggedMins / 60);
            const mT = Math.round(loggedMins % 60);
            const hW = Math.floor(workMins / 60);
            const mW = Math.round(workMins % 60);
            
            let missingHtml = "";
            if (missingMins > 0) {
                const hF = Math.floor(missingMins / 60);
                const mF = Math.round(missingMins % 60);
                missingHtml = `<span class="pd-total-falta" style="color: #a80000; border-left: 1px solid #c8c6c4; padding-left: 15px;">Faltan: <strong>${hF}h ${String(mF).padStart(2,'0')}m</strong></span>`;
            } else if (missingMins < 0) {
                const overMins = Math.abs(missingMins);
                const hE = Math.floor(overMins / 60);
                const mE = Math.round(overMins % 60);
                missingHtml = `<span class="pd-total-falta" style="color: #107c10; border-left: 1px solid #c8c6c4; padding-left: 15px;">Extra: <strong>${hE}h ${String(mE).padStart(2,'0')}m</strong></span>`;
            } else {
                missingHtml = `<span class="pd-total-falta" style="color: #107c10; border-left: 1px solid #c8c6c4; padding-left: 15px;"><strong>Jornada Completa ✔️</strong></span>`;
            }

            totalsDiv.innerHTML = `<span class="pd-total-imputado">⏳ Imputado: <strong>${hT}h ${String(mT).padStart(2,'0')}m</strong></span><span class="pd-total-jornada">Jornada: ${hW}h ${String(mW).padStart(2,'0')}m</span>${missingHtml}`;

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
            
            // Si todo fue bien y había overlay, lo ocultamos
            this.hideLoadingOverlay();

        } catch (err) {
            console.error("Error WebAPI:", err);
            this.hideLoadingOverlay();
        }
    }

    private showLunchModal(): void {
        const backdrop = document.createElement("div");
        backdrop.className = "pd-modal-backdrop";

        const modal = document.createElement("div");
        modal.className = "pd-modal";

        const title = document.createElement("h3");
        title.innerText = "🍔 Registrar Almuerzo";
        modal.appendChild(title);

        const lblTime = document.createElement("label");
        lblTime.innerText = "Hora de Inicio (HH:MM):";
        modal.appendChild(lblTime);

        const inputTime = document.createElement("input");
        inputTime.type = "time";
        inputTime.value = "14:00";
        inputTime.className = "pd-input";
        modal.appendChild(inputTime);

        const lblDur = document.createElement("label");
        lblDur.innerText = "Duración:";
        modal.appendChild(lblDur);

        const selectDur = document.createElement("select");
        selectDur.className = "pd-input";
        const options = [ {v:30, l:"30 minutos"}, {v:45, l:"45 minutos"}, {v:60, l:"1 hora"}, {v:90, l:"1.5 horas"} ];
        options.forEach(o => {
            const opt = document.createElement("option");
            opt.value = o.v.toString();
            opt.innerText = o.l;
            if(o.v === 60) opt.selected = true;
            selectDur.appendChild(opt);
        });
        modal.appendChild(selectDur);

        const btnDiv = document.createElement("div");
        btnDiv.style.display = "flex";
        btnDiv.style.gap = "10px";
        btnDiv.style.marginTop = "15px";

        const saveBtn = document.createElement("button");
        saveBtn.className = "pd-btn pd-btn-primary";
        saveBtn.innerText = "Guardar";
        saveBtn.onclick = async () => {
            saveBtn.disabled = true;
            saveBtn.innerText = "Guardando...";
            const timeVal = inputTime.value; 
            const durVal = parseInt(selectDur.value, 10);
            
            if(!timeVal) {
                alert("Introduce una hora válida");
                saveBtn.disabled = false;
                saveBtn.innerText = "Guardar";
                return;
            }

            const match = timeVal.match(/(\d{1,2}):(\d{2})/);
            if(match) {
                const h = parseInt(match[1], 10);
                const m = parseInt(match[2], 10);
                const hDec = h + (m/60);
                const endDec = hDec + (durVal/60);

                const params = this._context.parameters;
                const fechaRaw = params.sec_fecha?.raw;
                const recursoRaw = params.sec_recursoid?.raw;
                
                if (fechaRaw && recursoRaw && recursoRaw.length > 0) {
                    const recursoId = recursoRaw[0].id.replace(/[{}]/g, "").toLowerCase();
                    const baseDateStr = (fechaRaw as Date).toISOString();
                    
                    const payload = {
                        "msdyn_start": this.applyTimeToDateStr(baseDateStr, hDec),
                        "msdyn_end": this.applyTimeToDateStr(baseDateStr, endDec),
                        "msdyn_duration": durVal,
                        "msdyn_type": 192355000,
                        "msdyn_description": "Almuerzo",
                        "msdyn_bookableresource@odata.bind": `/bookableresources(${recursoId})`
                    };

                    try {
                        await this._context.webAPI.createRecord("msdyn_timeentry", payload);
                        if (backdrop.parentNode) backdrop.parentNode.removeChild(backdrop);
                        this.showLoadingOverlay();
                        await this.renderTimeline();
                    } catch (e) {
                        saveBtn.disabled = false;
                        saveBtn.innerText = "Guardar";
                        this.showErrorModal(e);
                    }
                }
            }
        };

        const cancelBtn = document.createElement("button");
        cancelBtn.className = "pd-btn pd-btn-secondary";
        cancelBtn.innerText = "Cancelar";
        cancelBtn.onclick = () => {
            if (backdrop.parentNode) backdrop.parentNode.removeChild(backdrop);
        };

        btnDiv.appendChild(saveBtn);
        btnDiv.appendChild(cancelBtn);
        modal.appendChild(btnDiv);

        backdrop.appendChild(modal);
        this._container.appendChild(backdrop);
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

            const inicioDec = this.getTimeDecFromDate(inicioRaw as Date);
            const finDec = this.getTimeDecFromDate(finRaw as Date);

            const entries = await this.fetchTimeEntries(this._context, fechaRaw as Date, recursoId);
            
            const sorted = entries
                .map(e => ({ 
                    start: this.getTimeDecFromDate(new Date(e.msdyn_start!)), 
                    end: this.getTimeDecFromDate(new Date(e.msdyn_end!))
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
        
        if (target.classList.contains('pd-delete-btn')) return;

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

        let minBound = 0;
        let maxBound = 24;
        
        this._currentEntries.forEach(other => {
            if (other.id === this._dragData.id) return;
            if (other.endDec <= this._dragData.originalStart + 0.01) {
                minBound = Math.max(minBound, other.endDec);
            }
            if (other.startDec >= this._dragData.originalEnd - 0.01) {
                maxBound = Math.min(maxBound, other.startDec);
            }
        });

        this._dragData.minBound = minBound;
        this._dragData.maxBound = maxBound;

        const rect = this._timelineEl.getBoundingClientRect();
        const range = this._maxViewHour - this._minViewHour;
        const pointerDec = this._isVertical ? 
            this._minViewHour + ((e.clientY - rect.top) / rect.height) * range : 
            this._minViewHour + ((e.clientX - rect.left) / rect.width) * range;

        this._dragData.offsetDecimal = pointerDec - this._dragData.originalStart;

        entryDiv.classList.add("pd-dragging");
        this._liveTooltip.style.opacity = "1";
    }

    private onPointerMove(e: PointerEvent): void {
        if (!this._isDragging || !this._dragTarget || !this._timelineEl) return;

        const rect = this._timelineEl.getBoundingClientRect();
        const range = this._maxViewHour - this._minViewHour;
        const pointerDec = this._isVertical ? 
            this._minViewHour + ((e.clientY - rect.top) / rect.height) * range : 
            this._minViewHour + ((e.clientX - rect.left) / rect.width) * range;

        const duration = this._dragData.originalEnd - this._dragData.originalStart;

        if (this._dragType === 'move') {
            let proposedStart = Math.round((pointerDec - this._dragData.offsetDecimal) * 12) / 12; // Snap cada 5 min
            const halfHourSnap = Math.round(proposedStart * 2) / 2; // Imán a las medias horas (30 min / 00 min)
            
            if (Math.abs(proposedStart - halfHourSnap) < (10 / 60)) proposedStart = halfHourSnap;
            
            proposedStart = Math.max(this._dragData.minBound, Math.min(this._dragData.maxBound - duration, proposedStart));
            
            this._dragData.newStart = proposedStart;
            this._dragData.newEnd = proposedStart + duration;

        } else if (this._dragType === 'left') {
            let proposedStart = Math.round(pointerDec * 12) / 12;
            const halfHourSnap = Math.round(pointerDec * 2) / 2;
            
            if (Math.abs(pointerDec - halfHourSnap) < (10 / 60)) proposedStart = halfHourSnap;

            proposedStart = Math.max(this._dragData.minBound, proposedStart);
            this._dragData.newStart = Math.min(proposedStart, this._dragData.newEnd - (5/60));

        } else if (this._dragType === 'right') {
            let proposedEnd = Math.round(pointerDec * 12) / 12;
            const halfHourSnap = Math.round(pointerDec * 2) / 2;
            
            if (Math.abs(pointerDec - halfHourSnap) < (10 / 60)) proposedEnd = halfHourSnap;

            proposedEnd = Math.min(this._dragData.maxBound, proposedEnd);
            this._dragData.newEnd = Math.max(proposedEnd, this._dragData.newStart + (5/60));
        }

        const startPercent = ((this._dragData.newStart - this._minViewHour) / range) * 100;
        const sizePercent = ((this._dragData.newEnd - this._dragData.newStart) / range) * 100;

        if (this._isVertical) {
            this._dragTarget.style.top = `${startPercent}%`;
            this._dragTarget.style.height = `${sizePercent}%`;
            const inner = this._dragTarget.querySelector('.pd-entry-inner-time') as HTMLElement;
            if(inner) inner.innerText = `${this.formatDecimalTime(this._dragData.newStart)}-${this.formatDecimalTime(this._dragData.newEnd)}`;
        } else {
            this._dragTarget.style.left = `${startPercent}%`;
            this._dragTarget.style.width = `${sizePercent}%`;
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
                
                const durationMins = Math.round((this._dragData.newEnd - this._dragData.newStart) * 60);

                const payload = {
                    msdyn_start: this.applyTimeToDateStr(this._dragData.origStartDate, this._dragData.newStart),
                    msdyn_end: this.applyTimeToDateStr(this._dragData.origEndDate, this._dragData.newEnd),
                    msdyn_duration: durationMins
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

    private getTimeDecFromDate(d: Date | undefined): number {
        if (!d) return 0;
        // Obliga a usar la hora UTC para evitar el desfase de la zona horaria del PC
        return d.getUTCHours() + (d.getUTCMinutes() / 60);
    }

    private applyTimeToDateStr(origIsoString: string, decimalHours: number): string {
        const d = new Date(origIsoString);
        const h = Math.floor(decimalHours);
        const m = Math.round((decimalHours - h) * 60);
        // Obliga a setear la hora en UTC para que Field Service reciba la hora exacta que el usuario ve
        d.setUTCHours(h, m, 0, 0);
        return d.toISOString();
    }

    private formatDecimalTime(decimalHours: number): string {
        const h = Math.floor(decimalHours);
        const m = Math.round((decimalHours - h) * 60);
        return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
    }

    private formatTimeObj(date: Date): string {
        return `${String(date.getUTCHours()).padStart(2, '0')}:${String(date.getUTCMinutes()).padStart(2, '0')}`;
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