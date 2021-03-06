export const PRIMITIVE_TYPES = ['string', 'bigint', 'number', 'boolean'];

export const forceArray = <T>(val: T | T[]): T[] => {
	if (Array.isArray(val)) return val;
	if (!val) return [];
	return [val];
};

export const xmlSafeValue = (val): string => val === undefined || val === null ? '' : String(val).replace(/&/g, '&amp;')
	.replace(/</g, '&lt;')
	.replace(/>/g, '&gt;')
	.replace(/"/g, '&quot;')
	.replace(/\n/g, '&#10;')
	.replace(/\r/g, '&#13;');

export const xmlSafeColumnName = (val): string => !val ? '' : String(val).replace(/[\s_]+/g, '').toLowerCase();

export const mergeDefault = (def, given): any => {
	if (!given) return deepClone(def);
	for (const key in def) {
		if (typeof given[key] === 'undefined') given[key] = deepClone(def[key]);
		else if (isObject(given[key])) given[key] = mergeDefault(def[key], given[key]);
	}

	return given;
};

export const deepClone = (source): any => {
	// Check if it's a primitive (with exception of function and null, which is typeof object)
	if (source === null || isPrimitive(source)) return source;
	if (Array.isArray(source)) {
		const output = [];
		for (const value of source) output.push(deepClone(value));
		return output;
	}
	if (isObject(source)) {
		const output = {};
		for (const [key, value] of Object.entries(source)) output[key] = deepClone(value);
		return output;
	}
	if (source instanceof Map) {
		const output = new (source.constructor as typeof Map)();
		for (const [key, value] of source.entries()) output.set(key, deepClone(value));
		return output;
	}
	if (source instanceof Set) {
		const output = new (source.constructor as typeof Set)();
		for (const value of source.values()) output.add(deepClone(value));
		return output;
	}
	return source;
};

export const isObject = (input): boolean => input && input.constructor === Object;

export const isPrimitive = (value): boolean => PRIMITIVE_TYPES.includes(typeof value);
