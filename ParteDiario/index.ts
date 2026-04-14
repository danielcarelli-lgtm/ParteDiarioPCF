import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as ComponentFramework from "powerapps-component-framework";

interface ITimeEntry extends ComponentFramework.WebApi.Entity {
    msdyn_start?: string;
    msdyn_end?: string;
}

export class ParteDiario implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _version = "v1.0.9";

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._context = context;
        this._container = container;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        this._container.innerHTML = "";

        const timelineWrapper = document.createElement("div");
        timelineWrapper.className = "pd-timeline-wrapper";

        const timeline = document.createElement("div");
        timeline.className = "pd-timeline";
        timelineWrapper.appendChild(timeline);
        this._container.appendChild(timelineWrapper);

        // Versión en la esquina
        const versionEl = document.createElement("div");
        versionEl.className = "pd-version-tag";
        versionEl.innerText = this._version;
        timelineWrapper.appendChild(versionEl);

        // Validación de parámetros
        const params = context.parameters;
        if (!params.sec_fecha.raw || !params.sec_horainicio.raw || !params.sec_horafin.raw || !params.sec_recursoid.raw?.[0]) {
            const errorMsg = document.createElement("div");
            errorMsg.innerText = "Faltan datos de configuración (Fecha, Horas o Recurso).";
            errorMsg.className = "pd-error";
            timelineWrapper.appendChild(errorMsg);
            return;
        }

        // Definición de jornada
        const hInicio = params.sec_horainicio.raw;
        const hFin = params.sec_horafin.raw;
        const inicioDecimal = hInicio.getHours() + (hInicio.getMinutes() / 60);
        const finDecimal = hFin.getHours() + (hFin.getMinutes() / 60);
        
        const jornadaDiv = document.createElement("div");
        jornadaDiv.className = "pd-jornada";
        jornadaDiv.style.left = `${(inicioDecimal / 24) * 100}%`;
        jornadaDiv.style.width = `${((finDecimal - inicioDecimal) / 24) * 100}%`;
        jornadaDiv.title = `Jornada: ${this.formatTime(hInicio)} - ${this.formatTime(hFin)}`;
        timeline.appendChild(jornadaDiv);

        // Fetch de registros
        this.fetchTimeEntries(context).then((entries) => {
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
            console.error(err);
            return undefined;
        });
    }

    private async fetchTimeEntries(context: ComponentFramework.Context<IInputs>): Promise<ITimeEntry[]> {
        const fecha = context.parameters.sec_fecha.raw as Date;
        const recursoId = context.parameters.sec_recursoid.raw[0].id;

        const y = fecha.getFullYear();
        const m = String(fecha.getMonth() + 1).padStart(2, '0');
        const d = String(fecha.getDate()).padStart(2, '0');
        
        const filter = `_msdyn_bookableresource_value eq ${recursoId} and Microsoft.Dynamics.CRM.On(PropertyName='msdyn_start',PropertyValue='${y}-${m}-${d}')`;
        const query = `?$filter=${filter}&$select=msdyn_start,msdyn_end`;

        const response = await context.webAPI.retrieveMultipleRecords("msdyn_timeentry", query);
        return response.entities as ITimeEntry[];
    }

    private formatTime(date: Date): string {
        return `${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
    }

    public getOutputs(): IOutputs { return {}; }

    public destroy(): void {
        this._container.innerHTML = "";
    }
}