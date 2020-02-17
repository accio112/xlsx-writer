declare var require: any
var _ = require('lodash');
import * as sampleData from './sample-data.json';
import { XlsxData } from '../interface/xlsx-data.js';

import {CONSTANTS, DELL_INFORMATION, DATE_OF_SERVICE, CUSTOMER_INFORMATION} from '../static/constants';

export function createDataOne() {

        const receivedData = sampleData['default'];
        const worksheetName = "WorkSheet";
        const columnWidth = 20;
        const logoData = {
            "name":"data:image/jpeg;base64,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",
            "topLeft":{
                "col":0.1,
                "row":0.1
            },
            "bottomRight":{
                "col":1.2,
                "row":4
            },
            "mergeCells":{
                "start":0,
                "end":0
            }
        };
        const titleData =    {
            "name": "Onsite Data Sanitization Report",
            "mergeCells": {
                "start": "A1",
                "end": "H4"
            },
            "style": {
                "name": "Arial",
                "size": "20",
                "bold": "true",
                "underline": "true",
                "alignment":{
                    "vertical":"top",
                    "horizontal" : "center"
                }
            }
        };

        const table1Style = {"name": "Arial",
                            "size": "14",
                            "bold": "true",
                            "underline": "single",
                            "color": "027CBB"
                        };
        let table1 = {
            "headers":{
                "data": [
                    {
                        "name":DELL_INFORMATION,
                        "mergeCells": {
                            "start": "A5",
                            "end": "B5"
                        },      
                        "style": table1Style
                    },
                    {
                        "name": CUSTOMER_INFORMATION,
                        "mergeCells": {
                            "start": "D5",
                            "end": "E5"
                        },
                        "style": table1Style
                    },
                    {
                        "name": DATE_OF_SERVICE,
                        "mergeCells": {
                            "start": "G5",
                            "end": "H5"
                        },
                        "style": table1Style
                    }
                ]
            },
            "rowsData":[
                [{"name":CONSTANTS.arsjobNum, "mergeCells":{"start":"A6"}},{"name":receivedData.arsjobNum, "mergeCells":{"start":"B6"}},{},{"name":CONSTANTS.customerName, "mergeCells":{"start":"D6"}},{"name":receivedData.customerName, "mergeCells":{"start":"E6"}},{},{"name":CONSTANTS.dateOfService, "mergeCells":{"start":"G6"}},{"name":receivedData.dateOfService, "mergeCells":{"start":"H6"}}],
                [{"name":CONSTANTS.vendorName, "mergeCells":{"start":"A7"}},{"name":receivedData.vendorName, "mergeCells":{"start":"B7"}},{},{"name":CONSTANTS.projectName, "mergeCells":{"start":"D7"}},{"name":receivedData.projectName, "mergeCells":{"start":"E7"}},{},{"name":CONSTANTS.startTime, "mergeCells":{"start":"G7"}},{"name":receivedData.startTime, "mergeCells":{"start":"H7"}}],
                [{"name":CONSTANTS.technicianName, "mergeCells":{"start":"A8"}},{"name":receivedData.technicianName, "mergeCells":{"start":"B8"}},{},{"name":CONSTANTS.collectionAddress, "mergeCells":{"start":"D8"}},{"name":receivedData.collectionAddress, "mergeCells":{"start":"E8"}},{},{"name":CONSTANTS.finishTime, "mergeCells":{"start":"G8"}},{"name":receivedData.finishTime, "mergeCells":{"start":"H8"}}],
                [{},{},{},{},{},{},{"name":CONSTANTS.numberOfSystemsProcessed, "mergeCells":{"start":"G9"}},{"name":receivedData.numberOfSystemsProcessed, "mergeCells":{"start":"H9"}}],
                [{"name":CONSTANTS.softwareName, "mergeCells":{"start":"A10"}},{"name":receivedData.softwareName, "mergeCells":{"start":"B10"}},{},{"name":CONSTANTS.country, "mergeCells":{"start":"D10"}},{"name":receivedData.country, "mergeCells":{"start":"E10"}},{},{"name":CONSTANTS.numberOfSystemsPassed, "mergeCells":{"start":"G10"}},{"name":receivedData.numberOfSystemsPassed, "mergeCells":{"start":"H10"}}],
                [{"name":CONSTANTS.softwareVersion, "mergeCells":{"start":"A11"}},{"name":receivedData.softwareVersion, "mergeCells":{"start":"B11"}},{},{"name":CONSTANTS.countryName, "mergeCells":{"start":"D11"}},{"name":receivedData.countryName, "mergeCells":{"start":"E11"}},{},{"name":CONSTANTS.numberOfSystemsFailed, "mergeCells":{"start":"G11"}},{"name":receivedData.numberOfSystemsFailed, "mergeCells":{"start":"H11"}}],
            ]
        };

        // start with A14
        let rows = [];
        const items = receivedData.items;
        _.forEach(items, item => {
            const details = item.driveDetails;
            let row = [];
            _.forEach(details, detail => {
                row= [];
                row.push(item.make);
                row.push(item.model);
                row.push(item.serviceTag);
                row.push(detail.driveModel);
                row.push(detail.driveSerialNumber);
                row.push(detail.status);
                row.push(detail.result);
                row.push(detail.exceptionsComment);
                rows.push(row);
            });
        });
        // console.log('rows', rows);
        let styledRows = [];
        let rowNumber = 14;
        _.forEach(rows, row=>{
            let styledRow = [];
            let col = "A";
            _.forEach(row, cell=>{
                let cellEntry = {};
                cellEntry["name"] = cell;
                cellEntry["mergeCells"] = {
                    "start": rowNumber+col
                }
                let num = 1+ col.charCodeAt(0);
                col = String.fromCharCode(num);
                styledRow.push(cellEntry);
            })
            rowNumber = rowNumber+1
            styledRows.push(styledRow);
        })
        // console.log('styledRows', styledRows);
        const table2HeaderStyle = { "underline": "true",
                                    "color": "027CBB",
                                    "bgColor": "000000",
                                    "fgColor": "D3D3D3",
                                    "border": "true"}
        let table2 = {
            "headers":{
                "data":[
                    {
                        "name":CONSTANTS.make,
                        "mergeCells": {
                            "start": "A13",
                        },      
                        "style": table2HeaderStyle
                    },
                    {
                        "name":CONSTANTS.model,
                        "mergeCells": {
                            "start": "B13",
                        },      
                        "style": table2HeaderStyle
                    },
                    {
                        "name":CONSTANTS.serviceTag,
                        "mergeCells": {
                            "start": "C13",
                        },      
                        "style":table2HeaderStyle
                    },
                    {
                        "name":CONSTANTS.driveModel,
                        "mergeCells": {
                            "start": "D13",
                        },      
                        "style": table2HeaderStyle
                    },
                    {
                        "name":CONSTANTS.driveSerialNumber,
                        "mergeCells": {
                            "start": "E13",
                        },      
                        "style": table2HeaderStyle
                    },
                    {
                        "name":CONSTANTS.status,
                        "mergeCells": {
                            "start": "F13",
                        },      
                        "style": table2HeaderStyle
                    },
                    {
                        "name":CONSTANTS.result,
                        "mergeCells": {
                            "start": "G13",
                        },      
                        "style":table2HeaderStyle
                    },
                    {
                        "name":CONSTANTS.exceptionsComment,
                        "mergeCells": {
                            "start": "H13",
                        },      
                        "style": table2HeaderStyle
                    }
                ]
            },
            "rowsData":styledRows
        };
        const dataToSend = {
            "worksheetName": worksheetName,
            "columnWidth":columnWidth,
            "image":{
                "data":[logoData]
            },
            "title":{
                "data":[titleData]
            },
            "tables":[table1, table2]
        }
        console.log('dataToSendOne', dataToSend);
        return dataToSend;
}
