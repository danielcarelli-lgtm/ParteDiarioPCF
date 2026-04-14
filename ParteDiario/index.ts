import { IInputs, IOutputs } from "./generated/ManifestTypes";

interface ITimeEntry extends ComponentFramework.WebApi.Entity {
    msdyn_start?: string;
    msdyn_end?: string;
}

export class ParteDiario implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _version = "v1.0.13";

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._context = context;
        this._container = container;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        this._container.innerHTML = "";

        try {
            const timelineWrapper = document.createElement("div");
            timelineWrapper.className = "pd-timeline-wrapper";

            const timeline = document.createElement("div");
            timeline.className = "pd-timeline";
            timelineWrapper.appendChild(timeline);
            this._container.appendChild(timelineWrapper);

            const versionEl = document.createElement("div");
            versionEl.className = "pd-version-tag";
            versionEl.innerText = this._version;
            timelineWrapper.appendChild(versionEl);

            const params = context.parameters;

            const fechaRaw = params.sec_fecha?.raw;
            const inicioRaw = params.sec_horainicio?.raw;
            const finRaw = params.sec_horafin?.raw;

            // CORRECCIÓN: sec_recursoid ahora es SingleLine.Text → raw es string directamente
            const recursoId = params.sec_recursoid?.raw as string;

            if (!fechaRaw || !inicioRaw || !finRaw || !recursoId) {
                const errorMsg = document.createElement("div");
                errorMsg.innerText = "Faltan datos de configuración en el formulario.";
                errorMsg.className = "pd-error";
                timelineWrapper.appendChild(errorMsg);
                return;
            }

            const hInicio = new Date(inicioRaw as string | number | Date);
            const hFin = new Date(finRaw as string | number | Date);

            const inicioDecimal = hInicio.getHours() + (hInicio.getMinutes() / 60);
            const finDecimal = hFin.getHours() + (hFin.getMinutes() / 60);

            const jornadaDiv = document.createElement("div");
            jornadaDiv.className = "pd-jornada";
            jornadaDiv.style.left = `${(inicioDecimal / 24) * 100}%`;
            jornadaDiv.style.width = `${((finDecimal - inicioDecimal) / 24) * 100}%`;
            jornadaDiv.title = `Jornada: ${this.formatTime(hInicio)} - ${this.formatTime(hFin)}`;
            timeline.appendChild(jornadaDiv);

            // CORRECCIÓN: fechaRaw parseado sin conversión UTC para evitar desfase de zona horaria
            this.fetchTimeEntries(context, fechaRaw as string | number | Date, recursoId).then((entries) => {
                if (!entries) return;

                entries.forEach((entry: ITimeEntry) => {
                    if (entry.msdyn_start && entry.msdyn_end) {
                        const startEntry = new Date(entry.msdyn_start);
                        const endEntry = new Date(entry.msdyn_end);

                        const startDec = startEntry.getHours() + (startEntry.getMinutes() / 60);
                        const endDec = endEntry.getHours() + (endEntry.getMinutes() / 60);

                        const isOutOfHours = (startDec < inicioDecimal) || (endDec > finDecimal);

                        const entryDiv = document.createElement("div");
                        entryDiv.className = isOutOfHours ? "pd-entry pd-out-of-hours" : "pd-entry pd-in-hours";
                        entryDiv.style.left = `${(startDec / 24) * 100}%`;
                        entryDiv.style.width = `${((endDec - startDec) / 24) * 100}%`;
                        entryDiv.title = `Imputación: ${this.formatTime(startEntry)} - ${this.formatTime(endEntry)}${isOutOfHours ? ' (Fuera de jornada)' : ''}`;

                        timeline.appendChild(entryDiv);
                    }
                });
                return undefined;
            }).catch(err => {
                console.error("Error WebAPI:", err);
                const errorMsg = document.createElement("div");
                errorMsg.innerText = "Error consultando los time entries.";
                errorMsg.className = "pd-error";
                timelineWrapper.appendChild(errorMsg);
            });

        } catch (error: unknown) {
            console.error("Error en updateView:", error);
            const errorMsg = document.createElement("div");
            const msg = error instanceof Error ? error.message : String(error);
            errorMsg.innerText = "Error interno del control: " + msg;
            errorMsg.className = "pd-error";
            this._container.appendChild(errorMsg);
        }
    }

    private async fetchTimeEntries(context: ComponentFramework.Context<IInputs>, fechaRaw: string | number | Date, recursoId: string): Promise<ITimeEntry[]> {
        // CORRECCIÓN: parseo seguro de fecha sin desfase UTC
        const fechaStr = String(fechaRaw).split('T')[0];
        const [y, m, d] = fechaStr.split('-');

        const fechaInicio = `${y}-${m}-${d}T00:00:00Z`;
        const fechaFin = `${y}-${m}-${d}T23:59:59Z`;

        // CORRECCIÓN: filtro OData estándar en lugar de Microsoft.Dynamics.CRM.On (más compatible)
        const filter = `_msdyn_bookableresource_value eq '${recursoId}' and msdyn_start ge ${fechaInicio} and msdyn_start le ${fechaFin}`;
        const query = `?$filter=${filter}&$select=msdyn_start,msdyn_end`;

        const response = await context.webAPI.retrieveMultipleRecords("msdyn_timeentry", query);
        return response.entities as ITimeEntry[];
    }

    private formatTime(date: Date): string {
        if (isNaN(date.getTime())) return "--:--";
        return `${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
    }

    public getOutputs(): IOutputs { return {}; }

    public destroy(): void {
        this._container.innerHTML = "";
    }
}