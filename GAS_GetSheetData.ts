//@ts-check
class getSheetData {
    protected _titleRow: number =1;
    protected _dataArray: Array<any>;

    protected _spreadsheet: any;
    protected _sheetName: string;

    protected index: number;
    protected columnNumbers: Object;

    static new(id, sheetName, titleRow) {
        let returnClass = new clientCloseDate();
        returnClass.openFileByID(id);
        returnClass.sheetName = sheetName;
        returnClass.titleRow = titleRow;
        return returnClass;
    }

    public openFileByID(id: string) {
        try {
            this._spreadsheet = SpreadsheetApp.openById(id);  // https://developers.google.com/drive/v3/web/picker  
        } catch (e) {
            throw new Error("It was not possible to open ID '{id}'. Do you have access permissions to this file?");
        }
    }

    public openFileByURL(url: string) {
        try {
            this._spreadsheet = SpreadsheetApp.openByUrl(url);  // https://developers.google.com/drive/v3/web/picker  
        } catch (e) {
            throw new Error("It was not possible to open URL '{url}'. Do you have access permissions to this file?");
        }
    }

    public set titleRow(row: number) {
        this._titleRow = row - 1; //Array start at 0
        this.index = row;
    }

    public set sheetName(name: string) {
        if (this._spreadsheet && !this._spreadsheet.getSheetByName(name)) {
            throw "Sheet name '{name}' does not exist.";
        }
        this._sheetName = name;
    }

    public getAllData(): void {
        return this._spreadsheet.getSheetByName(this._sheetName).getDataRange().getValues();
    }

    protected getColumnNumberByText(columnTitle: string): number {
        if (this.columnNumbers[columnTitle]) {
            return this.columnNumbers[columnTitle];
        }

        for (let i in this._spreadsheet[this._titleRow]) {
            if (this._spreadsheet[this.titleRow][i] == columnTitle) {
                return this.columnNumbers[columnTitle] = Number(i);
            }
        }

        throw "Column \"{columnTitle}\" not found on row " + (this._titleRow + 1) + ".";
    }

    public next() {
        this.index++;
        if (this.index > this._dataArray.length) {
            return {done: true};
        }
        return {value: this._dataArray[this.index], done: false};
    }

    public previous() {
        this.index--;
        if (this.index <= this._titleRow) {
            return {done: true};
        }
        return {value: this._dataArray[this.index], done: false};
    }

    public getValue(columnTitle: string) {
        return this._dataArray[this.index][this.getColumnNumberByText(columnTitle)];
    }

    [Symbol.iterator]() {
        return { next: this.next() };
    }
}