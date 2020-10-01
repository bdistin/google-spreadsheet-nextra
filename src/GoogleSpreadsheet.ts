import fetch from 'node-fetch';
import * as xml2js from 'xml2js';
import { JWT } from 'google-auth-library';

import { STATUS_CODES } from 'http';

import { SpreadsheetWorksheet } from './SpreadsheetWorksheet';
import { SpreadsheetCell } from './SpreadsheetCell';
import { SpreadsheetRow } from './SpreadsheetRow';
import { forceArray, mergeDefault, deepClone, xmlSafeColumnName, xmlSafeValue } from './util';
import { URLSearchParams } from 'url';

const parser = new xml2js.Parser({
	explicitArray: false,
	explicitRoot: false
});
const parseString = (str: string): Promise<unknown> => new Promise((resolve, reject) => {
	parser.parseString(str, (data, error) => {
		if (error) return reject(error);
		return resolve(data);
	});
});

const GOOGLE_FEED_URL = 'https://spreadsheets.google.com/feeds/';
const GOOGLE_AUTH_SCOPE = ['https://spreadsheets.google.com/feeds'];

const REQUIRE_AUTH_MESSAGE = 'You must authenticate to modify sheet data';

type HTTP_METHODS = 'GET' | 'HEAD' | 'POST' | 'PUT' | 'DELETE' | 'CONNECT' | 'OPTIONS' | 'TRACE';

export enum GoogleSpreadsheetVisibility {
	public,
	private
}

export enum GoogleSpreadsheetProjection {
	full,
	values
}

export enum GoogleSpreadsheetAuthMode {
	anonymous,
	token,
	jwt
}

export interface SpreadsheetInfo {
	id: string;
	title: string;
	updated: string;
	author: string;
	worksheets: SpreadsheetWorksheet[];
}

export interface CellsQuery {
	['min-row']?: number;
	['max-row']?: number;
	['min-col']?: number;
	['max-col']?: number;
	['return-empty']?: boolean;
}

export interface APIRowQuery {
	['start-index']?: number;
	['max-results']?: number;
	['orderby']?: string;
	['reverse']?: string;
	['sq']?: string;
}

export interface RowsQuery {
	offset?: number;
	limit?: number;
	orderBy?: string;
	reverse?: boolean;
	query?: string;
}

export interface WorksheetOptions {
	title?: string;
	rowCount?: number;
	colCount?: number;
	headers?: string[];
}

/**
 * The main controlling class for creating connections to GoogleSheets
 */
export class GoogleSpreadsheet {

	/**
	 * The authorization data
	 */
	private googleAuth = null;

	/**
	 * The visibility of the connected spreadsheet
	 */
	private visibility = GoogleSpreadsheetVisibility.public;

	/**
	 * The projection of the connected spreadsheet
	 */
	private projection = GoogleSpreadsheetProjection.values;

	/**
	 * The authorization mode of the connected spreadsheet
	 */
	private authMode = GoogleSpreadsheetAuthMode.anonymous;

	/**
	 * The key for the connected spreadsheet
	 */
	private spreadsheetKey: string;

	/**
	 * The json webtoken client for the connected spreadsheet
	 */
	private jwtClient: JWT = null;

	/**
	 * The options passed to the constructor
	 */
	private options;

	/**
	 * The metadata about the connected spreadsheet
	 */
	public info: SpreadsheetInfo;

	/**
	 * The worksheets for the connected spreadsheet
	 */
	public worksheets: SpreadsheetWorksheet[];

	public constructor(spreadsheetKey: string, authID, options) {
		this.options = options || {};
		if (!spreadsheetKey) throw new Error('Spreadsheet key not provided');
		this.spreadsheetKey = spreadsheetKey;

		this.setAuthAndDependencies(authID);
	}

	/**
	 * Sets the authorization, visibility, and projection for this spreadsheet
	 * @param auth The authorization
	 */
	private setAuthAndDependencies(auth: string): void {
		this.googleAuth = auth;

		if (!this.options.visibility) this.visibility = this.googleAuth ? GoogleSpreadsheetVisibility.private : GoogleSpreadsheetVisibility.public;
		if (!this.options.projection) this.projection = this.googleAuth ? GoogleSpreadsheetProjection.full : GoogleSpreadsheetProjection.values;
	}

	/**
	 * Sets the authorization for this spreadsheet
	 * @param auth Auth data
	 */
	public setAuthToken(auth): void {
		if (this.authMode === GoogleSpreadsheetAuthMode.anonymous) this.authMode = GoogleSpreadsheetAuthMode.token;
		this.setAuthAndDependencies(auth);
	}

	/**
	 * Sets authorization to use a service account for authorization
	 * @param creds The credentials
	 */
	public useServiceAccountAuth(creds): Promise<void> {
		if (typeof creds === 'string') creds = require(creds);
		this.jwtClient = new JWT(creds.client_email, null, creds.private_key, GOOGLE_AUTH_SCOPE, null);
		return this.renewJwtAuth();
	}

	/**
	 * Renews the json webtoken client authorization
	 */
	private async renewJwtAuth(): Promise<void> {
		this.authMode = GoogleSpreadsheetAuthMode.jwt;
		const credentials = await this.jwtClient.authorize();
		this.setAuthToken({
			type: credentials.token_type,
			value: credentials.access_token,
			expires: credentials.expiry_date
		});
	}

	/**
	 * Returns if this class is authorized
	 */
	public get isAuthActive(): boolean {
		return !!this.googleAuth;
	}

	/**
	 * Downloads info from google sheets for the connected spreadsheet
	 */
	public async getInfo(): Promise<SpreadsheetInfo> {
		const { data } = await this.makeFeedRequest(['worksheets', this.spreadsheetKey], 'GET', null);
		if (!data) throw new Error('No response to getInfo call');

		this.info = {
			id: data.id,
			title: data.title,
			updated: data.updated,
			author: data.author,
			worksheets: []
		};

		for (const workSheet of forceArray(data.entry)) this.info.worksheets.push(new SpreadsheetWorksheet(this, workSheet));

		this.worksheets = this.info.worksheets;
		return this.info;
	}

	/**
	 * Adds a new worksheet in the connected spreadsheet
	 * @param options The worksheet options
	 */
	public async addWorksheet(options: WorksheetOptions = {}): Promise<SpreadsheetWorksheet> {
		if (!this.isAuthActive) throw new Error(REQUIRE_AUTH_MESSAGE);

		mergeDefault({
			// need a unique title
			title: `Worksheet ${new Date()}`,
			rowCount: 50,
			colCount: 20
		}, options);

		// if column headers are set, make sure the sheet is big enough for them
		if (options.headers && options.headers.length > options.colCount) options.colCount = options.headers.length;

		const dataXML = [
			'<entry xmlns="http://www.w3.org/2005/Atom" xmlns:gs="http://schemas.google.com/spreadsheets/2006">',
			`<title>${options.title}</title>`,
			`<gs:rowCount>${options.rowCount}</gs:rowCount>`,
			`<gs:colCount>${options.colCount}</gs:colCount>`,
			'</entry>'
		].join('');

		const { data } = await this.makeFeedRequest(['worksheets', this.spreadsheetKey], 'POST', dataXML);

		const sheet = new SpreadsheetWorksheet(this, data);
		this.worksheets = this.worksheets || [];
		this.worksheets.push(sheet);
		await sheet.setHeaderRow(options.headers);
		return sheet;
	}

	/**
	 * Deletes a worksheet from the connected spreadsheet
	 * @param worksheet The worksheet id or SpreadsheetWorksheet instance to delete
	 */
	public async removeWorksheet(worksheet: number | SpreadsheetWorksheet): Promise<void> {
		if (!this.isAuthActive) throw new Error(REQUIRE_AUTH_MESSAGE);
		if (worksheet instanceof SpreadsheetWorksheet) return worksheet.del();
		await this.makeFeedRequest(`${GOOGLE_FEED_URL}worksheets/${this.spreadsheetKey}/private/full/${worksheet}`, 'DELETE', null);
		return undefined;
	}

	/**
	 * Gets rows from the connected spreadsheet for the given worksheet and options
	 * @param worksheetID The worksheet id
	 * @param options The rows query
	 */
	public async getRows(worksheetID: number, options: RowsQuery = {}): Promise<SpreadsheetRow[]> {
		// the first row is used as titles/keys and is not included
		const query: APIRowQuery = {};

		if (options.offset) query['start-index'] = options.offset;
		if (options.limit) query['max-results'] = options.limit;
		// eslint-disable-next-line dot-notation
		if (options.orderBy) query['orderby'] = options.orderBy;
		// eslint-disable-next-line dot-notation
		if (options.reverse) query['reverse'] = 'true';
		// eslint-disable-next-line dot-notation
		if (options.query) query['sq'] = options.query;

		const { data, xml } = await this.makeFeedRequest(['list', this.spreadsheetKey, worksheetID], 'GET', query);
		if (!data) throw new Error('No response to getRows call');

		// gets the raw xml for each entry -- this is passed to the row object so we can do updates on it later

		let entriesXML = xml.match(/<entry[^>]*>([\s\S]*?)<\/entry>/g);

		// need to add the properties from the feed to the xml for the entries
		const feedProps = deepClone(data.$);
		delete feedProps['gd:etag'];
		const feedPropsStr = feedProps.reduce((str, val, key): string => `${str}${key}='${val}' `, '');
		entriesXML = entriesXML.map((_xml): string => _xml.replace('<entry ', `<entry ${feedPropsStr}`));

		return forceArray(data.entry).map((entry, i): SpreadsheetRow => new SpreadsheetRow(this, entry, entriesXML[i]));
	}

	/**
	 * Adds a row in the specified worksheet
	 * @param worksheetID The worksheet id
	 * @param rowData The row data to add
	 */
	public async addRow(worksheetID: number, rowData): Promise<SpreadsheetRow> {
		const dataXML = ['<entry xmlns="http://www.w3.org/2005/Atom" xmlns:gsx="http://schemas.google.com/spreadsheets/2006/extended">'];

		for (const [key, value] of Object.entries(rowData)) {
			if (key !== 'id' && key !== 'title' && key !== 'content' && key !== '_links') {
				dataXML.push(`<gsx:${xmlSafeColumnName(key)}>${xmlSafeValue(value)}</gsx:${xmlSafeColumnName(key)}>`);
			}
		}

		dataXML.push('</entry>');
		const { data, xml } = await this.makeFeedRequest(['list', this.spreadsheetKey, worksheetID], 'POST', dataXML.join('\n'));
		const entriesXML = xml.match(/<entry[^>]*>([\s\S]*?)<\/entry>/g);
		return new SpreadsheetRow(this, data, entriesXML[0]);
	}

	/**
	 * Get cells from the specified worksheet
	 * @param worksheetID The worksheet id
	 * @param options The cells query
	 */
	public async getCells(worksheetID: number, options: CellsQuery = {}): Promise<SpreadsheetCell[]> {
		const { data } = await this.makeFeedRequest(['cells', this.spreadsheetKey, worksheetID], 'GET', options);
		if (!data) throw new Error('No response to getCells call');
		// eslint-disable-next-line dot-notation
		return forceArray(data['entry']).map((entry): SpreadsheetCell => new SpreadsheetCell(this, this.spreadsheetKey, worksheetID, entry));
	}

	/**
	 * Performs all http requests to google sheets
	 * @param urlParams The url params
	 * @param method The http method
	 * @param queryOrData The query or data to send
	 */
	public async makeFeedRequest(urlParams: string | (string | number)[], method: HTTP_METHODS, queryOrData: string | CellsQuery | APIRowQuery): Promise<{xml: string, data: any}> {
		let url;
		let body: string | null = null;

		const headers = {};
		if (typeof urlParams === 'string') {
			// used for edit / delete requests
			url = urlParams;
		} else if (Array.isArray(urlParams)) {
			// used for get and post requests
			urlParams.push(GoogleSpreadsheetVisibility[this.visibility], GoogleSpreadsheetProjection[this.projection]);
			url = GOOGLE_FEED_URL + urlParams.join('/');
		}

		// auth
		if (this.authMode === GoogleSpreadsheetAuthMode.jwt && this.googleAuth && this.googleAuth.expires > Date.now()) await this.renewJwtAuth();

		// request
		headers['Gdata-Version'] = '3.0';

		// eslint-disable-next-line dot-notation
		if (this.googleAuth) headers['Authorization'] = this.googleAuth.type === 'Bearer' ? `Bearer ${this.googleAuth.value}` : `GoogleLogin auth=${this.googleAuth}`;

		if (method === 'POST' || method === 'PUT') {
			headers['content-type'] = 'application/atom+xml';
			if (url.includes('/batch')) headers['If-Match'] = '*';
			body = queryOrData as string;
		}

		if (method === 'GET' && queryOrData) {
			url += `?${new URLSearchParams(queryOrData as string | { [key: string]: string | string[] })}`
				// replacements are needed for using structured queries on getRows
				.replace(/%3E/g, '>')
				.replace(/%3D/g, '=')
				.replace(/%3C/g, '<');
		}

		const response = await fetch(url, { method, headers, body });

		if (response.status === 200 && response.headers.get('content-type').includes('text/html')) throw new Error(`Sheet is private. Use authentication or make public.`);
		if (response.status === 401) throw new Error('Invalid authorization key.');

		const xml = await response.text();

		if (response.status >= 400) throw new Error(`HTTP error ${response.status} (${STATUS_CODES[response.status]}) - ${xml.replace(/&quot;/g, '"')}`);

		return { xml, data: !xml ? null : await parseString(xml) };
	}

}
