"use strict";
var adaptors = require('../index.js');
var xls2json = adaptors.xls2json;
var xlsx2json = adaptors.xlsx2json;


xls2json('./test/data/上海工具公司11390_20.xls');

xlsx2json('./test/data/Country and Region Source (ISO 3166-1).xlsx');