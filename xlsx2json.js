"use strict";
var XLSX = require('xlsx');
var _ = require('lodash');
var pinyin = require('pinyin');

function getStartRow(exclude){
	if(exclude === undefined || exclude === null){
		return 1;
	}
	else{
		var startRow = 1;
		var t = _.range(1, exclude[exclude.length-1]+1);

		_.forEach(t, function(i){
			if(_.findIndex(exclude, function(e){return i === e;}) === -1){
				startRow = i;
				return false;
			}
			else{
				startRow ++;
			}
		});
		return startRow;
	}
}

module.exports = function(fileName){
	var workbook = XLSX.readFile(fileName);
//	var workbook = xlsxObj.workbook;

	var sheetNameList = workbook.SheetNames;
	var parseResult = [];

	_.forEach(sheetNameList, function(sheetName){
		var sheet = workbook.Sheets[sheetName];

		var data = {};
		var colNameArray = [];

		var excludeRows = null;

		if(sheet['!merges'] !== undefined){
			excludeRows = [];
			var merges = sheet['!merges'];

			_.forEach(merges, function(mergeInfo){
				var startRow = mergeInfo.s.r + 1;
				var endRow = parseInt(mergeInfo.e.r, 10) + 2;

				var mergeRowsArray = _.range(startRow, endRow);

				_.forEach(mergeRowsArray, function(row){
					if(_.findIndex(excludeRows, function(r){return r === row;}) === -1){
						excludeRows.push(row);
					}
				});
			});
		}

		var startRow = getStartRow(excludeRows);

		for(var index in sheet) {
			if (sheet.hasOwnProperty(index)) {
				if (!index.match(/^!/)) {
					var item = sheet[index];
					var col, row;
					var matchResult = index.match(/\d+$/);
					if(matchResult !== null) {
						row = parseInt(matchResult[0], 10);
						col = index.substring(0, matchResult.index);

						if(row >= startRow && _.findIndex(excludeRows, row) === -1){
							if(data[col] === undefined){
								data[col] = {
									label: col,
									values: []
								};
							}

							if(row === startRow){
								data[col].description = item.w;

								var colName = pinyin(item.w, {
									style: pinyin.STYLE_NORMAL,
									heteronym: false
								}).join('_');

								var colNameSuffix = 0;

								_.forEach(colNameArray, function(c){
									if(c.name === colName){
										colNameSuffix = c.suffix++;
									}
								});

								colNameArray.push({
									name: colName,
									suffix: colNameSuffix
								});

								if(colNameSuffix > 0){
									colName = colName+'_'+colNameSuffix;
								}

								data[col].name = colName;
							}
							else{
								if(data[col].dataType === undefined){
									var t = item.t;
									switch(t){
										case 'n':
											data[col].dataType = 'number';
											break;
										case 's':
											data[col].dataType = 'string';
											break;
										case 'str':
											data[col].dataType = 'string';
											break;
										case 'b':
											data[col].dataType = 'boolean';
											break;
										case 'e':
											data[col].dataType = 'error';
											break;
										default:
											data[col].dataType = 'string';
											break;
									}
								}
								data[col].values[row-startRow-1] = item.w;
							}
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

		if(parseResultItem.data.length > 0){
			parseResult.push(parseResultItem);
		}
	});

	return parseResult;
};