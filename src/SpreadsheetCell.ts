import GoogleSpreadsheet from './GoogleSpreadsheet';
import { xmlSafeValue, forceArray } from './util';

export default class SpreadsheetCell {

	private spreadsheet: GoogleSpreadsheet;
	private batchID: string;
	private worksheetID: number;
	private spreadsheetKey: string;
	private id: string;
	private links: Map<string, string> = new Map();
	private _formula: string;
	private _numericValue: number;
	private _value: string;
	private _needsSave: boolean = false;

	public row: number;
	public col: number;

	public constructor(spreadsheet: GoogleSpreadsheet, spreadsheetKey: string, worksheetID: number, data) {
		this.spreadsheet = spreadsheet;
		// eslint-disable-next-line dot-notation
		this.row = parseInt(data['gs:cell']['$']['row']);
		// eslint-disable-next-line dot-notation
		this.col = parseInt(data['gs:cell']['$']['col']);
		this.batchID = `R${this.row}C${this.col}`;

		if (data.id === `https://spreadsheets.google.com/feeds/cells/${spreadsheetKey}/${worksheetID}/${this.batchID}`) {
			this.worksheetID = worksheetID;
			this.spreadsheetKey = spreadsheetKey;
		} else {
			this.id = data.id;
		}

		for (const link of forceArray(data.link)) {
			if (link.$.rel === 'self' && link.$.href === this.getSelf()) continue;
			if (link.$.rel === 'edit' && link.$.href === this.getEdit()) continue;
			this.links.set(link.$.rel, link.$.href);
		}

		this.updateValuesFromResponseData(data);
	}

	public getID(): string {
		return this.id || `https://spreadsheets.google.com/feeds/cells/${this.spreadsheetKey}/${this.worksheetID}/${this.batchID}`;
	}

	public getEdit(): string {
		return this.links.get('edit') || this.getID().replace(this.batchID, `private/full/${this.batchID}`);
	}

	public getSelf(): string {
		return this.links.get('edit') || this.getID().replace(this.batchID, `private/full/${this.batchID}`);
	}

	public getXML(link: string): string {
		this._needsSave = false;
		return [
			'	<entry>',
			`		<batch:id>${this.batchID}</batch:id>`,
			'		<batch:operation type="update"/>',
			`		<id>${link}/${this.batchID}</id>`,
			`		<link rel="edit" type="application/atom+xml" href="${this.getEdit()}"/>`,
			`		<gs:cell row="${this.row}" col="${this.col}" inputValue="${this.valueForSave}"/>`,
			'	</entry>'
		].join('\n');
	}

	public updateValuesFromResponseData(data): void {
		// formula value
		// eslint-disable-next-line dot-notation
		const inputVal = data['gs:cell']['$']['inputValue'];
		// inputValue can be undefined so substr throws an error
		// still unsure how this situation happens
		this._formula = inputVal && inputVal.startsWith('=') ? inputVal : undefined;

		// numeric values
		// eslint-disable-next-line dot-notation
		this._numericValue = data['gs:cell']['$']['numericValue'] !== undefined ? parseFloat(data['gs:cell']['$']['numericValue']) : undefined;

		// the main "value" - its always a string
		// eslint-disable-next-line dot-notation
		this._value = data['gs:cell']['_'] || '';
	}

	public async setValue(value): Promise<void> {
		this.value = value;
		await this.save();
	}

	public get value(): string {
		return this._needsSave ? '*SAVE TO GET NEW VALUE*' : this._value;
	}

	public set value(val) {
		if (!val) {
			this._clearValue();
			return;
		}

		const numericVal = parseFloat(val);
		if (!isNaN(numericVal)) {
			this._numericValue = numericVal;
			this._value = val.toString();
		} else {
			this._numericValue = undefined;
			this._value = val;
		}

		if (typeof val === 'string' && val.startsWith('=')) {
			// use the getter to clear the value
			this.formula = val;
		} else {
			this._formula = undefined;
		}
	}

	public get formula(): string {
		return this._formula;
	}

	public set formula(val) {
		if (!val) {
			this._clearValue();
			return;
		}
		if (!val.startsWith('=')) throw new Error('Formulas must start with "="');

		this._numericValue = undefined;
		this._needsSave = true;
		this._formula = val;
	}

	// eslint-disable-next-line @typescript-eslint/explicit-function-return-type
	public get numericValue() {
		return this._numericValue;
	}

	public set numericValue(val: string | number) {
		if (val === undefined || val === null) {
			this._clearValue();
			return;
		}

		const parsed = typeof val === 'string' ? parseFloat(val) : val;
		if (Number.isNaN(parsed) || !isFinite(parsed)) throw new Error('Invalid numeric value assignment');

		this._value = parsed.toString();
		this._numericValue = parsed;
		this._formula = undefined;
	}

	public get valueForSave(): string {
		return xmlSafeValue(this._formula || this._value);
	}

	public async save(): Promise<void> {
		this._needsSave = false;

		const id = this.getID();
		const dataXML = [
			'<entry xmlns=\'http://www.w3.org/2005/Atom\' xmlns:gs=\'http://schemas.google.com/spreadsheets/2006\'>',
			`<id>${id}</id>`,
			`<link rel="edit" type="application/atom+xml" href="${id}"/>`,
			`<gs:cell row="${this.row}" col="${this.col}" inputValue="${this.valueForSave}"/>`,
			'</entry>'
		].join('');

		const { data } = await this.spreadsheet.makeFeedRequest(this.getEdit(), 'PUT', dataXML);
		this.updateValuesFromResponseData(data);
	}

	public async del(): Promise<void> {
		await this.setValue('');
	}

	private _clearValue(): void {
		this._formula = undefined;
		this._numericValue = undefined;
		this._value = '';
	}

}
