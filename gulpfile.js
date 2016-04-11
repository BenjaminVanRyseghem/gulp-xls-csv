"use strict";

let gulp = require("gulp");
let fs = require("fs");
let path = require("path");
let commandLineArgs = require("command-line-args");

gulp.task("xls-to-csv", function() {
	let knownOptions = [
		{
			name: "verbose", alias: "v", type: Boolean
		},
		{
			name: "dest", alias: "d", type: String
		},
		{
			name: "src", type: String, defaultOption: true
		}
	];

	let cli = commandLineArgs(knownOptions);

	let options = {};

	// Support the `default` task
	let args = process.argv.slice(2);
	if (this.seq.indexOf(args[0]) !== -1) {
		args = process.argv.slice(3);
	}

	try {
		options = cli.parse(args);
	} catch (e) {
		exit(e.message, options.verbose);
	}

	if (!options.src) {
		exit("Source file is mandatory", options.verbose);
	}

	if (options.verbose) {
		console.log(`About to import ${options.src}`);
	}

	try {
		fs.accessSync(options.src, fs.R_OK);
	} catch (e) {
		exit(`${options.src} can not be accessed.`, options.verbose);
	}

	let stats = fs.statSync(options.src);

	if (!stats.isFile()) {
		exit(`${options.src} MUST be a file.`, options.verbose);
	}

	let XLSX = require("xlsx");
	let workbook = XLSX.readFile(options.src);

	let length = workbook.SheetNames.length;
	workbook.SheetNames.forEach((name) => {
		let ws = workbook.Sheets[name];
		let csvString = XLSX.utils.sheet_to_csv(ws);

		let maxColumn = getMaxColumn(ws);

		let postfix = length === 1 ? null : name;
		let destination = options.dest || buildDestinationName(options.src, "csv", postfix);

		if (options.verbose) {
			console.log(`About to write ${options.src} to ${destination}`);
		}

		csvString = removeTableName(csvString);
		csvString = trimColumns(csvString, maxColumn);
		fs.writeFileSync(destination, csvString, "utf8");
	});
});

gulp.task("csv-to-xls", function(callback) {
	let knownOptions = [
		{
			name: "verbose", alias: "v", type: Boolean
		},
		{
			name: "dest", alias: "d", type: String
		},
		{
			name: "src", type: String, defaultOption: true
		}
	];

	let cli = commandLineArgs(knownOptions);

	let options = {};

	// Support the `default` task
	let args = process.argv.slice(2);
	if (this.seq.indexOf(args[0]) !== -1) {
		args = process.argv.slice(3);
	}

	try {
		options = cli.parse(args);
	} catch (e) {
		exit(e.message, options.verbose);
	}

	if (!options.src) {
		exit("Source file is mandatory", options.verbose);
	}

	if (options.verbose) {
		console.log(`About to import ${options.src}`);
	}

	try {
		fs.accessSync(options.src, fs.R_OK);
	} catch (e) {
		exit(`${options.src} can not be accessed.`, options.verbose);
	}

	let stats = fs.statSync(options.src);

	if (!stats.isFile()) {
		exit(`${options.src} MUST be a file.`, options.verbose);
	}

	let csv = require("csv-parser");

	let data = [];
	let maxColumns = 0;
	let keys;

	fs.createReadStream(options.src)
		.pipe(csv())
		.on("data", (d) => {
			if (!keys) {
				keys = Object.keys(d);
				maxColumns = keys.length;
			}

			data.push(d);
		})
		.on("end", () => {
			let maxRows = data.length;

			let excelbuilder = require("msexcel-builder");

			let destination = options.dest || buildDestinationName(options.src, "xls");

			let workbook = excelbuilder.createWorkbook(path.dirname(destination), path.basename(destination));
			let sheet1 = workbook.createSheet("Sheet 1", maxColumns, maxRows + 1);

			keys.forEach((key, j) => {
				sheet1.set(j + 1, 1, key);
			});

			for (let i = 2; i < maxRows; i++) {
				let line = data[i];

				keys.forEach((key, j) => {
					let cell = line[key];
					sheet1.set(j + 1, i, cell);
				});
			}

			if (options.verbose) {
				console.log(`About to write ${options.src} to ${destination}`);
			}
			workbook.save(callback);
		});
});

//
// DEFAULT
//

gulp.task("default", ["xls-to-csv"]);

//
// Helpers
//

function trimColumns(csv, max) {
	let lines = csv.split("\n");
	lines = lines.map((line) => {
		let data = line.split(",");
		data = data.slice(0, max);
		return data.join(",");
	});
	return lines.join("\n");
}

function getMaxColumn(worksheet) {
	let max = 0;
	for (let cell in worksheet) {
		if (worksheet.hasOwnProperty(cell)) {
			if (cell[0] !== "!") {
				let column = cell[0];
				let index = convertToNumbers(column);
				max = Math.max(max, index);
			}
		}
	}
	return max;
}

function convertToNumbers(str) {
	let arr = "abcdefghijklmnopqrstuvwxyz".split("");
	return str.replace(/[a-z]/ig, (match) => {
		return arr.indexOf(match.toLowerCase()) + 1;
	});
}

function removeTableName(csv) {
	let lines = csv.split("\n");
	lines.splice(0, 1);
	return lines.join("\n");
}

function exit(error, verbose) {
	console.error(error);
	if (verbose) {
		console.log("Exiting");
	}
	process.exit(1);
}

function buildDestinationName(source, extension, postfix) {
	let ext = path.extname(source);
	let base = path.basename(source, ext);
	let post = "";
	if (postfix) {
		post = ` - ${postfix}`;
	}
	return `./${base}${post}.${extension}`;
}
