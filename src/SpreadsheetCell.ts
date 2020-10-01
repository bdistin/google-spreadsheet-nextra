import { xmlSafeValue, forceArray } from './util';

import type { GoogleSpreadsheet } from './GoogleSpreadsheet';

/**
 * Represents a cell in a given worksheet for the connected spreadsheet
 */
export class SpreadsheetCell {

	/**
	 * The google spreadsheet this belongs to
	 */
	private spreadsheet: GoogleSpreadsheet;

	/**
	 * The batch id
	 */
	private batchID: string;

	/**
	 * The worksheet id
	 */
	private worksheetID: number;

	/**
	 * The spreadsheet key
	 */
	private spreadsheetKey: string;

	/**
	 * The id for this cell
	 */
	private id: string;

	/**
	 * The links for this cell
	 */
	private links: Map<string, string> = new Map();

	/**
	 * The formula stored in this cell
	 */
	private _formula: string;

	/**
	 * The numeric value of this cell
	 */
	private _numericValue: number;

	/**
	 * The actual value of this cell
	 */
	private _value: string;

	/**
	 * If this cell needs to be synced back to google sheets
	 */
	private _needsSave = false;

	/**
	 * The row number for this cell
	 */
	public row: number;

	/**
	 * The column number for this cell
	 */
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

	/**
	 * The id for this cell
	 */
	public getID(): string {
		return this.id || `https://spreadsheets.google.com/feeds/cells/${this.spreadsheetKey}/${this.worksheetID}/${this.batchID}`;
	}

	/**
	 * The edit link for this cell
	 */
	public getEdit(): string {
		return this.links.get('edit') || this.getID().replace(this.batchID, `private/full/${this.batchID}`);
	}

	/**
	 * The link to this cell
	 */
	public getSelf(): string {
		return this.links.get('edit') || this.getID().replace(this.batchID, `private/full/${this.batchID}`);
	}

	/**
	 * Gets the xml for this cell
	 * @param link The link
	 */
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

	/**
	 * Updates this cell from the response data
	 * @param data The data from the response
	 */
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

	/**
	 * Sets the value of this cell
	 * @param value The data you want the value to be
	 */
	public async setValue(value): Promise<void> {
		this.value = value;
		await this.save();
	}

	/**
	 * Gets the value of this cell
	 */
	public get value(): string {
		return this._needsSave ? '*SAVE TO GET NEW VALUE*' : this._value;
	}

	/**
	 * Allows you to set the value of this cell
	 */
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

	/**
	 * Gets the formula of this cell
	 */
	public get formula(): string {
		return this._formula;
	}

	/**
	 * Allows you to set the value of the formula for this cell
	 */
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

	/**
	 * Gets the numeric value of this cell
	 */
	// eslint-disable-next-line @typescript-eslint/explicit-function-return-type
	public get numericValue() {
		return this._numericValue;
	}

	/**
	 * Sets the numeric value of this cell
	 */
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

	/**
	 * Makes the value safe for serializing to xml
	 */
	public get valueForSave(): string {
		return xmlSafeValue(this._formula || this._value);
	}

	/**
	 * Saves changes to the value to google sheets
	 */
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

	/**
	 * Deletes the data from this cell
	 */
	public async del(): Promise<void> {
		await this.setValue('');
	}

	/**
	 * Clears the values of this cell
	 */
	private _clearValue(): void {
		this._formula = undefined;
		this._numericValue = undefined;
		this._value = '';
	}

}
