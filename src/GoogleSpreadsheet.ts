import fetch from 'node-fetch';
import * as xml2js from 'xml2js';
import { JWT } from 'google-auth-library';

import * as http from 'http';
import * as querystring from 'querystring';
import { promisify } from 'util';

import SpreadsheetWorksheet from './SpreadsheetWorksheet';
import SpreadsheetCell from './SpreadsheetCell';
import SpreadsheetRow from './SpreadsheetRow';
import { forceArray, mergeDefault, deepClone, xmlSafeColumnName, xmlSafeValue, isObject } from './util';

const parser = new xml2js.Parser({
	explicitArray: false,
	explicitRoot: false
});
const parseString = promisify(parser.parseString);

const GOOGLE_FEED_URL = 'https://spreadsheets.google.com/feeds/';
const GOOGLE_AUTH_SCOPE = ['https://spreadsheets.google.com/feeds'];

const REQUIRE_AUTH_MESSAGE = 'You must authenticate to modify sheet data';


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
	worksheets: Array<SpreadsheetWorksheet>;
}

export default class GoogleSpreadsheet {

	private googleAuth = null;
	private visibility = GoogleSpreadsheetVisibility.public;
	private projection = GoogleSpreadsheetProjection.values;
	private authMode = GoogleSpreadsheetAuthMode.anonymous;
	private ssKey: string;
	private jwtClient: JWT = null;
	private options;

	public info: SpreadsheetInfo;
	public worksheets: Array<SpreadsheetWorksheet>;

	constructor(ssKey, authID, options) {
		this.options = options || {};
		if (!ssKey) throw new Error('Spreadsheet key not provided');
		this.ssKey = ssKey;

		this.setAuthAndDependencies(authID);
	}

	private setAuthAndDependencies(auth) {
		this.googleAuth = auth;
		if (!this.options.visibility) {
			this.visibility = this.googleAuth ? GoogleSpreadsheetVisibility.private : GoogleSpreadsheetVisibility.public;
		}
		if (!this.options.projection) {
			this.projection = this.googleAuth ? GoogleSpreadsheetProjection.full : GoogleSpreadsheetProjection.values;
		}
	}

	public setAuthToken(auth) {
		if (this.authMode === GoogleSpreadsheetAuthMode.anonymous) this.authMode = GoogleSpreadsheetAuthMode.token;
		this.setAuthAndDependencies(auth);
	}

	public useServiceAccountAuth(creds) {
		if (typeof creds === 'string') {
			creds = require(creds);
		}
		this.jwtClient = new JWT(creds.client_email, null, creds.private_key, GOOGLE_AUTH_SCOPE, null);
		return this.renewJwtAuth();
	}

	private async renewJwtAuth() {
		this.authMode = GoogleSpreadsheetAuthMode.jwt;
		const credentials = await this.jwtClient.authorize();
		this.setAuthToken({
			type: credentials.token_type,
			value: credentials.access_token,
			expires: credentials.expiry_date
		});
	}

	public get isAuthActive() {
		return !!this.googleAuth;
	}

	public async getInfo() {
		const { data } = await this.makeFeedRequest(['worksheets', this.ssKey], 'GET', null);
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

	public async addWorksheet(options) {
		if (!this.isAuthActive) throw new Error(REQUIRE_AUTH_MESSAGE);

		mergeDefault({
			title: `Worksheet ${new Date()}`,	// need a unique title
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

		const { data } = await this.makeFeedRequest( ['worksheets', this.ssKey], 'POST', dataXML);

		const sheet = new SpreadsheetWorksheet(this, data);
		this.worksheets = this.worksheets || [];
		this.worksheets.push(sheet);
		await sheet.setHeaderRow(options.headers);
		return sheet;
	}

	public async removeWorksheet(sheetID) {
		if (!this.isAuthActive) throw new Error(REQUIRE_AUTH_MESSAGE);
		if (sheetID instanceof SpreadsheetWorksheet) return sheetID.del();
		await this.makeFeedRequest(`${GOOGLE_FEED_URL}worksheets/${this.ssKey}/private/full/${sheetID}`, 'DELETE', null);
	}

	public async getRows(worksheetID, options) {
		// the first row is used as titles/keys and is not included
		const query = {};

		if ( options.offset ) query['start-index'] = options.offset;
		else if ( options.start ) query['start-index'] = options.start;

		if ( options.limit ) query['max-results'] = options.limit;
		else if ( options.num ) query['"max-results'] = options.num;

		if ( options.orderby ) query['orderby'] = options.orderby;
		if ( options.reverse ) query['reverse'] = 'true';
		if ( options.query ) query['sq'] = options.query;

		const { data, xml } = await this.makeFeedRequest(['list', this.ssKey, worksheetID], 'GET', query);
		if (!data) throw new Error('No response to getRows call');

		// gets the raw xml for each entry -- this is passed to the row object so we can do updates on it later

		let entriesXML = xml.match(/<entry[^>]*>([\s\S]*?)<\/entry>/g);

		// need to add the properties from the feed to the xml for the entries
		const feedProps = deepClone(data.$);
		delete feedProps['gd:etag'];
		const feedPropsStr = feedProps.reduce((str, val, key) => `${str}${key}='${val}' `, '');
		entriesXML = entriesXML.map((xml) => xml.replace('<entry ', `<entry ${feedPropsStr}`));

		return forceArray(data.entry).map((entry, i) => new SpreadsheetRow(this, entry, entriesXML[i]));
	}

	public async addRow(worksheetID, rowData) {
		const dataXML = ['<entry xmlns="http://www.w3.org/2005/Atom" xmlns:gsx="http://schemas.google.com/spreadsheets/2006/extended">'];

		for (const [key, value] of Object.entries(rowData)) {
			if (key !== 'id' && key !== 'title' && key !== 'content' && key !== '_links') {
				dataXML.push(`<gsx:${xmlSafeColumnName(key)}>${xmlSafeValue(value)}</gsx:${xmlSafeColumnName(key)}>`);
			}
		}

		dataXML.push('</entry>');
		const { data, xml } = await this.makeFeedRequest(['list', this.ssKey, worksheetID], 'POST', dataXML.join('\n'));
		const entriesXML = xml.match(/<entry[^>]*>([\s\S]*?)<\/entry>/g);
		return new SpreadsheetRow(this, data, entriesXML[0]);
	}

	public async getCells(worksheetID, options = {}) {
		const { data } = await this.makeFeedRequest(['cells', this.ssKey, worksheetID], 'GET', options);
		if (!data) throw new Error('No response to getCells call');
		return forceArray(data['entry']).map(entry => new SpreadsheetCell(this, this.ssKey, worksheetID, entry));
	}

	public async makeFeedRequest(urlParams, method, queryOrData) {
		let url;
		const headers = {};
		if (typeof urlParams === 'string') {
			// used for edit / delete requests
			url = urlParams;
		} else if (Array.isArray(urlParams)) {
			// used for get and post requets
			urlParams.push(GoogleSpreadsheetVisibility[this.visibility], GoogleSpreadsheetProjection[this.projection]);
			url = GOOGLE_FEED_URL + urlParams.join('/');
		}

		// auth
		if (this.authMode === GoogleSpreadsheetAuthMode.jwt && this.googleAuth && this.googleAuth.expires > + new Date()) await this.renewJwtAuth();

		// request
		headers['Gdata-Version'] = '3.0';

		if (this.googleAuth) headers['Authorization'] = this.googleAuth.type === 'Bearer' ? `Bearer ${this.googleAuth.value}` : `GoogleLogin auth=${this.googleAuth}`;

		if (method === 'POST' || method === 'PUT') {
			headers['content-type'] = 'application/atom+xml';
			if (url.includes('/batch')) headers['If-Match'] = '*';
		}

		if (method === 'GET' && queryOrData) {
			url += `?${querystring.stringify(queryOrData)}`
				// replacements are needed for using structured queries on getRows
				.replace(/%3E/g, '>')
				.replace(/%3D/g, '=')
				.replace(/%3C/g, '<');
		}

		const response = await fetch(url, { method, headers, body: method === 'POST' || method === 'PUT' ? queryOrData : null });

		if (response.status === 200 && response.headers.get('content-type').includes('text/html')) throw new Error(`Sheet is private. Use authentication or make public.`);
		if (response.status === 401) throw new Error('Invalid authorization key.');

		const body = await response.text();

		if (response.status >= 400) {
			const stringifiedBody = isObject(body) ? JSON.stringify(body) : body.replace(/&quot;/g, '"');
			throw new Error(`HTTP error ${response.status} (${http.STATUS_CODES[response.status]}) - ${stringifiedBody}`);
		}

		return { xml: body, data: !body ? null : await parseString(body) };
	}

}
