import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as XLSX from 'xlsx';   // npm install --save xlsx
import { saveAs } from 'file-saver';  // npm install --save file-saver



export class JsonToExcel implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    /**
     * Variables
     */
    private _notifyOutputChanged: () => void;
    private _fileName: string;
    private _json: string;
    private _sortOrder: string;
    private _trigger: boolean;

    /**
     * Empty constructor.
     */
    constructor() {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        // Add control initialization code
        this._notifyOutputChanged = notifyOutputChanged;
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Add code to update control view

        this._fileName = context.parameters.fileName.raw ?? "";

        this._json = context.parameters.JSONContent.raw ?? "";

        this._sortOrder = context.parameters.sortOrder.raw ?? "";

        this._trigger = context.parameters.trigger.raw ?? false;

        if (this._trigger) {
            this.downloadExcel();
            this._notifyOutputChanged();
        }
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs {
        return {
            trigger: false
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }

    private downloadExcel() {
        //convert json to array
        let data = JSON.parse(this._json);
        let sortOrder = JSON.parse(this._sortOrder)
        //new workbook
        let wb = XLSX.utils.book_new();
        //new worksheet
        let ws = XLSX.utils.json_to_sheet(data, { header: sortOrder });
        //add worksheet to workbook
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        //adjust column width
        const columnWidths = sortOrder.map((column: string) => ({
            wch: Math.max(column.length, ...data.map((row: any) => String(row[column]).length))
        }));
        ws['!cols'] = columnWidths;
        //save workbook
        const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        saveAs(blob, `${this._fileName}.xlsx`);

        // const workbook: XLSX.WorkBook = { Sheets: { 'data': ws }, SheetNames: ['data'] };
        // XLSX.writeFile(workbook, this.filename);

    }

}
