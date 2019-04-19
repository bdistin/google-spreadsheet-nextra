import GoogleSpreadsheet, { CellsQuery, RowsQuery } from './GoogleSpreadsheet';
import { forceArray } from './util';
import SpreadsheetCell from './SpreadsheetCell';
import SpreadsheetRow from './SpreadsheetRow';

export interface EditOptions {
	rowCount?: number;
	colCount?: number;
	title?: string;
}

export default class SpreadsheetWorksheet {

	private links = new Map();
	private spreadsheet: GoogleSpreadsheet;
	public url: string;
	public id: number;
	public title: string;
	public rowCount: number;
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

	public resize(rowCount: number, colCount: number): Promise<void> {
		return this.edit({ rowCount, colCount });
	}

	public setTitle(title: string): Promise<void> {
		return this.edit({ title });
	}

	public async clear(): Promise<void> {
		const { colCount, rowCount } = this;
		await this.resize(1, 1);
		const cells = await this.getCells();
		await cells[0].setValue(null);
		await this.resize(rowCount, colCount);
	}

	public getRows(options: RowsQuery): Promise<SpreadsheetRow[]> {
		return this.spreadsheet.getRows(this.id, options);
	}

	public getCells(options: CellsQuery = {}): Promise<SpreadsheetCell[]> {
		return this.spreadsheet.getCells(this.id, options);
	}

	public addRow(data): Promise<SpreadsheetRow> {
		return this.spreadsheet.addRow(this.id, data);
	}

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

	public async del(): Promise<void> {
		await this.spreadsheet.makeFeedRequest(this.links.get('edit'), 'DELETE', null);
	}

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
