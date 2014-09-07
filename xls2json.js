"use strict";
var XLS = require('xlsjs');
var _ = require('lodash');
module.exports = function(fileName){
	var workbook = XLS.readFile(fileName, {cellNF: true});

	var sheetNameList = workbook.SheetNames;

	var parseResult = [];

	_.forEach(sheetNameList, function(sheetName){
		var sheet = workbook.Sheets[sheetName];

		var data = {};

		for(var index in sheet){
			if(sheet.hasOwnProperty(index)){
				if(!index.match(/^!/)){
					var item = sheet[index];
					var col, row;
					var matchResult = index.match(/\d+$/);
					if(matchResult !== null){
						row = parseInt(matchResult[0], 10);
						col = index.substring(0, matchResult.index);

						if(data[col] === undefined){
							data[col] = {
								colName: col,
								values: []
							};
						}

						if(row === 1){
							data[col].description = item.w;
						}
						else{
							data[col].values[row-2] = item.w;
						}
					}
				}
			}
		}

		var parseResultItem = {
			sheetName: sheetName,
			data: []
		};

		for(var colIndex in data){
			if(data.hasOwnProperty(colIndex)){
				parseResultItem.data.push(data[colIndex]);
			}
		}

		parseResult.push(parseResultItem);
	});

	return parseResult;
};