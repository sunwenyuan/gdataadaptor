"use strict";
var XLS = require('xlsjs');
var _ = require('lodash');
var pinyin = require('pinyin');
module.exports = function(fileName){
	var workbook = XLS.readFile(fileName, {cellNF: true});

	var sheetNameList = workbook.SheetNames;

	var parseResult = [];

	_.forEach(sheetNameList, function(sheetName){
		var sheet = workbook.Sheets[sheetName];

		var data = {};

		var colNameArray = [];

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
								label: col,
								values: []
							};
						}

						if(row === 1){
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

		if(parseResultItem.data.length > 0){
			parseResult.push(parseResultItem);
		}
	});

	return parseResult;
};