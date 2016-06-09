"use strict";

let gulp = require("gulp");
let csvToXls = require("./index").csvToXls;
let xlsToCsv = require("./index").xlsToCsv;

gulp.task("xls-to-csv", () => {
	return gulp.src("excel-file.xlsx")
		.pipe(xlsToCsv({
			filename: "comma-separated-file"
		}))
		.pipe(gulp.dest("./dest"));
});

gulp.task("csv-to-xls", () => {
	return gulp.src("comma-separated-file.csv")
		.pipe(csvToXls())
		.pipe(gulp.dest("./dest"));
});