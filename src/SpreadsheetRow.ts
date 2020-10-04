import { xmlSafeColumnName, forceArray, xmlSafeValue } from './util';

import type { GoogleSpreadsheet } from './GoogleSpreadsheet';

/**
 * Represents a row in a worksheet
 */
export class SpreadsheetRow extends Map {

	/**
	 * The spreadsheet this belongs to
	 */
	private spreadsheet: GoogleSpreadsheet;

	/**
	 * The xml value of this row
	 */
	private xml: string;

	/**
	 * The links for this row
	 */
	private links = new Map<string, string>();

	public constructor(spreadsheet: GoogleSpreadsheet, data: any, xml: string) {
		super();
		this.spreadsheet = spreadsheet;
		this.xml = xml;

		for (const [key, value] of Object.entries(data)) {
			if (key.startsWith('gsx:')) this.set(key === 'gsx:' ? key.substring(0, 3) : key.substring(4), typeof value === 'object' && Object.keys(value).length === 0 ? null : value);
			else if (key === 'id') this.set(key, value);
			// eslint-disable-next-line dot-notation
			else if (value['_']) this.set(key, value['_']);
			// eslint-disable-next-line dot-notation
			else if (key === 'link') for (const link of forceArray(value)) this.links.set(link['$'].rel, link['$'].href);
		}
	}

	/**
	 * Saves changes to this row
	 */
	public async save(): Promise<void> {
		let dataXML = this.xml.replace('<entry>', '<entry xmlns=\'http://www.w3.org/2005/Atom\' xmlns:gsx=\'http://schemas.google.com/spreadsheets/2006/extended\'>');
		for (const [key, value] of this) {
			// Need to double check against RegExp Redos
			dataXML = dataXML.replace(
				new RegExp(`<gsx:${xmlSafeColumnName(key)}>([\\s\\S]*?)</gsx:${xmlSafeColumnName(key)}>`),
				`<gsx:${xmlSafeColumnName(key)}>${xmlSafeValue(value)}</gsx:${xmlSafeColumnName(key)}>`
			);
		}
		await this.spreadsheet.makeFeedRequest(this.links.get('edit'), 'PUT', dataXML);
	}

	/**
	 * Deletes this row
	 */
	public async del(): Promise<void> {
		await this.spreadsheet.makeFeedRequest(this.links.get('edit'), 'DELETE', null);
	}

}
