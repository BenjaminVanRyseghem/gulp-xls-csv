"use strict";

let Stream = require("stream");
let path = require("path");

let gutil = require("gulp-util");

let plexer = require("plexer");
let csv = require("csv-parser");
let excelbuilder = require("msexcel-builder");

function buildFilename(file) {
	let fullpath = file.history[0];
	return path.basename(fullpath, path.extname(fullpath));
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

			if (path.extname(file.history[0]) !== ".csv") {
				outputStream.emit("error", new gutil.PluginError(PLUGIN_NAME, `Extension ".csv" expected but ${path.extname(file.history[0])} was found`));
				done();
				return;
			}

			let keys;
			let data = [];
			let maxColumns = 0;

			let xlsStream;

			let filename = options.filename || buildFilename(file);

			if (file.isBuffer()) {
				xlsStream = new Stream.PassThrough();
				xlsStream.push(file.contents);
				xlsStream.end();
			}

			if (file.isStream()) {
				xlsStream = file.contents;
			}

			let xlsFile = new gutil.File({
				cwd: file.cwd,
				base: file.base,
				path: `${path.join(file.base, filename)}.xlsx`
			});

			xlsStream
				.pipe(csv())
				.on("data", (row) => {
					if (!keys) {
						keys = Object.keys(row);
						maxColumns = keys.length;
					}

					data.push(row);
				})
				.on("error", (err) => {
					outputStream.emit("error", new gutil.PluginError(PLUGIN_NAME, `Error while reading the stream! ${err.message}`));
				})
				.on("end", () => {
					let maxRows = data.length;

					let workbook = excelbuilder.createWorkbook(file.base, `${path.join(file.base, filename)}.xlsx`);
					let sheet1 = workbook.createSheet("Sheet 1", maxColumns, maxRows + 1);

					keys.forEach((key, j) => {
						sheet1.set(j + 1, 1, key);
					});

					for (let i = 0; i < maxRows; i++) {
						let line = data[i];

						keys.forEach((key, j) => {
							let cell = line[key];
							sheet1.set(j + 1, i + 2, cell);
						});
					}

					workbook.generate((err, zip) => {
						if (err) {
							outputStream.emit("error", new gutil.PluginError(PLUGIN_NAME, "Error while generating workbook!"));
							done();
						}

						xlsFile.contents = zip.generate({
							type: "nodebuffer"
						});
						outputStream.push(xlsFile);
						outputStream.end();
					});
				});

			done();
		};

		return stream;
	};
};
