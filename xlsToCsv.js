"use strict";

let Stream = require("stream");
let path = require("path");

let gutil = require("gulp-util");

let plexer = require("plexer");
let XLSX = require("xlsx");

function buildFilename(file) {
	let fullpath = file.history[0];
	return path.basename(fullpath, path.extname(fullpath));
}

function concatenate(arrays) {
	let totalLength = 0;
	for (let arr of arrays) {
		totalLength += arr.length;
	}
	let result = new Uint8Array(totalLength);
	let offset = 0;
	for (let arr of arrays) {
		result.set(arr, offset);
		offset += arr.length;
	}
	return Uint8Array.from(result);
}

function streamWorkbook(workbook, file, fileName, stream) {
	let length = workbook.SheetNames.length;
	workbook.SheetNames.forEach((name) => {
		let postfix = length === 1 ? "" : `-${name}`;

		let csvFile = new gutil.File({
			cwd: file.cwd,
			base: file.base,
			path: `${path.join(file.base, fileName)}${postfix}.csv`
		});

		let ws = workbook.Sheets[name];
		let csvString = XLSX.utils.sheet_to_csv(ws);

		csvFile.contents = new Buffer(csvString);
		stream.push(csvFile);
	});
}

module.exports = function(PLUGIN_NAME) {
	return function(options) {
		options = options || {};

		let inputStream = new Stream.Transform({objectMode: true});
		let outputStream = new Stream.PassThrough({objectMode: true});
		let stream = plexer({objectMode: true}, inputStream, outputStream);

		inputStream._transform = function(file, encoding, done) {
			if (file.isNull()) {
				outputStream.write(file);
				done();
				return;
			}

			if (path.extname(file.history[0]) !== ".xlsx") {
				outputStream.emit("error", new gutil.PluginError(PLUGIN_NAME, `Extension ".xlsx" expected but ${path.extname(file.history[0])} was found`));
				done();
				return;
			}

			let filename = options.filename || buildFilename(file);

			if (file.isStream()) {
				let acc = [];
				file.contents
					.on("data", (d) => {
						acc = acc.concat(d);
					})
					.on("error", (err) => {
						outputStream.emit("error", new gutil.PluginError(PLUGIN_NAME, `Error while reading the stream! ${err.message}`));
						done();
					})
					.on("end", () => {
						let data = concatenate(acc);
						let workbook = XLSX.read(data, {
							type: "buffer"
						});
						streamWorkbook(workbook, file, filename, outputStream);
						outputStream.end();
					});
			}

			if (file.isBuffer()) {
				let workbook = XLSX.read(file.contents);
				streamWorkbook(workbook, file, filename, outputStream);
				outputStream.end();
			}

			done();
		};

		return stream;
	};
};
