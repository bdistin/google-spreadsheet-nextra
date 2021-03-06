import { forceArray } from './util';

import type { GoogleSpreadsheet, CellsQuery, RowsQuery } from './GoogleSpreadsheet';
import type { SpreadsheetCell } from './SpreadsheetCell';
import type { SpreadsheetRow } from './SpreadsheetRow';

export interface EditOptions {
	rowCount?: number;
	colCount?: number;
	title?: string;
}

/**
 * Represents a worksheet in the connected spreadsheet
 */
export class SpreadsheetWorksheet {

	/**
	 * The links for this worksheet
	 */
	private links = new Map();

	/**
	 * The google spreadsheet this worksheet belongs to
	 */
	private spreadsheet: GoogleSpreadsheet;

	/**
	 * The url to this worksheet
	 */
	public url: string;

	/**
	 * The id of this worksheet
	 */
	public id: number;

	/**
	 * The title of this worksheet
	 */
	public title: string;

	/**
	 * The number of rows in this worksheet
	 */
	public rowCount: number;

	/**
	 * The number of columns in this worksheet
	 */
	public colCount: number;

	public constructor(spreadsheet: GoogleSpreadsheet, data) {
		this.spreadsheet = spreadsheet;
		this.url = data.id;
		this.id = parseInt(data.id.substring(data.id.lastIndexOf('/') + 1));
		this.title = data.title;
		this.rowCount = parseInt(data['gs:rowCount']);
		this.colCount = parseInt(data['gs:colCount']);

		for (const link of forceArray(data.link)) this.links.set(link.$.rel, link.$.href);

		const cells = this.links.get('http://schemas.google.com/spreadsheets/2006#cellsfeed');

		this.links.set('cells', cells);
		this.links.set('bulkcells', `${cells}/batch`);
	}

	/**
	 * Edits title, rowcount, and or column count for this worksheet
	 * @param param0 The edit options
	 */
	public async edit({ title, rowCount, colCount }: EditOptions): Promise<void> {
		const xml = [
			'<entry xmlns="http://www.w3.org/2005/Atom" xmlns:gs="http://schemas.google.com/spreadsheets/2006">',
			`<title>${title || this.title}</title>`,
			`<gs:rowCount>${rowCount || this.rowCount}</gs:rowCount>`,
			`<gs:colCount>${colCount || this.colCount}</gs:colCount>`,
			'</entry>'
		].join('');

		const { data } = await this.spreadsheet.makeFeedRequest(this.links.get('edit'), 'PUT', xml);

		this.title = data.title;
		this.rowCount = parseInt(data['gs:rowCount']);
		this.colCount = parseInt(data['gs:colCount']);
	}

	/**
	 * Resizes this worksheet
	 * @param rowCount The new row count
	 * @param colCount The new column count
	 */
	public resize(rowCount: number, colCount: number): Promise<void> {
		return this.edit({ rowCount, colCount });
	}

	/**
	 * Sets the title of this worksheet
	 * @param title The new title
	 */
	public setTitle(title: string): Promise<void> {
		return this.edit({ title });
	}

	/**
	 * Clears this worksheet
	 */
	public async clear(): Promise<void> {
		const { colCount, rowCount } = this;
		await this.resize(1, 1);
		const cells = await this.getCells();
		await cells[0].setValue(null);
		await this.resize(rowCount, colCount);
	}

	/**
	 * Gets rows from this worksheet
	 * @param options The row query
	 */
	public getRows(options: RowsQuery): Promise<SpreadsheetRow[]> {
		return this.spreadsheet.getRows(this.id, options);
	}

	/**
	 * Gets cells from this worksheet
	 * @param options The cells query
	 */
	public getCells(options: CellsQuery = {}): Promise<SpreadsheetCell[]> {
		return this.spreadsheet.getCells(this.id, options);
	}

	/**
	 * Adds a row to this worksheet
	 * @param data The row data
	 */
	public addRow(data): Promise<SpreadsheetRow> {
		return this.spreadsheet.addRow(this.id, data);
	}

	/**
	 * Updates many cells at the same time
	 * @param cells The cells to be updated
	 */
	public async bulkUpdateCells(cells: SpreadsheetCell[]): Promise<void> {
		const link = this.links.get('cells');
		const entries = cells.map((cell): string => cell.getXML(link));

		const dataXML = [
			'<feed xmlns="http://www.w3.org/2005/Atom" xmlns:batch="http://schemas.google.com/gdata/batch" xmlns:gs="http://schemas.google.com/spreadsheets/2006">',
			`	  <id>${link}</id>`,
			entries.join('\n'),
			'</feed>'
		].join('\n');

		const { data } = await this.spreadsheet.makeFeedRequest(this.links.get('bulkcells'), 'POST', dataXML);

		// update all the cells
		if (data.entry && data.entry.length) {
			const cellsByBatchID = cells.reduce((acc, entry): any => {
				// eslint-disable-next-line dot-notation
				acc[entry['batchId']] = entry;
				return acc;
			}, {});

			for (const cellData of data.entry) cellsByBatchID[cellData['batch:id']].updateValuesFromResponseData(cellData);
		}
	}

	/**
	 * Deletes this worksheet
	 */
	public async del(): Promise<void> {
		await this.spreadsheet.makeFeedRequest(this.links.get('edit'), 'DELETE', null);
	}

	/**
	 * Sets a header row
	 * @param values The values for the header row
	 */
	public async setHeaderRow(values: string[]): Promise<void> {
		if (!values) return undefined;
		if (values.length > this.colCount) throw new Error(`Sheet is not large enough to fit ${values.length} columns. Resize the sheet first.`);

		const cells = await this.getCells({
			'min-row': 1,
			'max-row': 1,
			'min-col': 1,
			'max-col': this.colCount,
			'return-empty': true
		});

		for (const cell of cells) cell.value = values[cell.col - 1] || '';

		return this.bulkUpdateCells(cells);
	}

}
