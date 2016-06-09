(function() {
	"use strict";

	// consts
	let PLUGIN_NAME = "gulp-xls-csv";

	module.exports = {
		xlsToCsv: require("./xlsToCsv")(PLUGIN_NAME),
		csvToXls: require("./csvToXls")(PLUGIN_NAME)
	};
})();
