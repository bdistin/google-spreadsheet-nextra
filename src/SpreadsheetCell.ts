import GoogleSpreadsheet from './GoogleSpreadsheet';
import { xmlSafeValue, forceArray } from './util';

export default class SpreadsheetCell {

	private spreadsheet: GoogleSpreadsheet;
	private batchID: string;
	private wsID: string;
	private ss: string;
	private id: string;
	private links: Map<string, string> = new Map();
	private _formula: string;
	private _numericValue: number;
	private _value: string;
	private _needsSave: boolean = false;

	public row: number;
	public col: number;

	public constructor(spreadsheet: GoogleSpreadsheet, ssKey: string, worksheetID: string, data) {
		this.spreadsheet = spreadsheet;
		this.row = parseInt(data['gs:cell']['$']['row']);
		this.col = parseInt(data['gs:cell']['$']['col']);
		this.batchID = `R${this.row}C${this.col}`;

		if (data.id === `https://spreadsheets.google.com/feeds/cells/${ssKey}/${worksheetID}/${this.batchID}`) {
			this.wsID = worksheetID;
			this.ss = ssKey;
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

	public getID() {
		return this.id || `https://spreadsheets.google.com/feeds/cells/${this.ss}/${this.wsID}/${this.batchID}`;
	}

	public getEdit() {
		return this.links.get('edit') || this.getID().replace(this.batchID, `private/full/${this.batchID}`);
	}

	public getSelf() {
		return this.links.get('edit') || this.getID().replace(this.batchID, `private/full/${this.batchID}`);
	}

	public updateValuesFromResponseData(_data) {
		// formula value
		const input_val = _data['gs:cell']['$']['inputValue'];
		// inputValue can be undefined so substr throws an error
		// still unsure how this situation happens
		this._formula = input_val && input_val.startsWith('=') ? input_val : undefined;

		// numeric values
		this._numericValue = _data['gs:cell']['$']['numericValue'] !== undefined ? parseFloat(_data['gs:cell']['$']['numericValue']) : undefined;

		// the main "value" - its always a string
		this._value = _data['gs:cell']['_'] || '';
	}

	public async setValue(new_value) {
		this.value = new_value;
		await this.save();
	}

	public get value() {
		return this._needsSave ? '*SAVE TO GET NEW VALUE*' : this._value;
	}

	public set value(val) {
		if (!val) {
			this._clearValue();
			return;
		}

		const numeric_val = parseFloat(val);
		if (!isNaN(numeric_val)) {
			this._numericValue = numeric_val;
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

	public get formula() {
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

	public get valueForSave() {
		return xmlSafeValue(this._formula || this._value);
	}

	public async save() {
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

	public async del() {
		await this.setValue('');
	}

	private _clearValue() {
		this._formula = undefined;
		this._numericValue = undefined;
		this._value = '';
	}

}
