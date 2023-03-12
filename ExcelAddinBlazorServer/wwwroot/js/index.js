let sheetsCreated = []
let dotNetReference;


async function createWorksheets(sourceSheet, sourceTable, variable, value) {
    await Office.onReady();
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sourceSheet);
        const table = sheet.tables.getItem(sourceTable);
        let filter = table.columns.getItem(variable).filter;
        console.log(value);
        filter.apply({
            filterOn: Excel.FilterOn.values,
            values: [value]
        });
        const visibleRange = table
            .getRange()
            .getVisibleView()
            .load("values");
        await context.sync();
        await context.application.suspendScreenUpdatingUntilNextSync();

        let values = visibleRange.values;
        let rowCount = values.length;
        let columnCount = values[0].length;
        let worksheetDest = sourceSheet + variable + value;
        context.workbook.worksheets.getItemOrNullObject(worksheetDest).delete();
        let sheetDest = context.workbook.worksheets.add(worksheetDest);

        let range = sheetDest.getRangeByIndexes(0, 0, rowCount, columnCount);
        range.values = values;
        sheetDest.getUsedRange().format.autofitColumns();
        sheetDest.getUsedRange().format.autofitRows();
        let newTable = sheetDest.tables.add(range, true);

        newTable.name = worksheetDest;
        table.clearFilters();
        sheetsCreated.push(worksheetDest)
    });
}


async function getValuesFromColumn(worksheetSource, tableSource, column) {
    return new Office.Promise(async function (resolve) {
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getItem(worksheetSource);
            const table = sheet.tables.getItem(tableSource);
            const columnRange = table.columns.getItem(column).getDataBodyRange().load("values");
            await sheet.context.sync();
            const columnValues = columnRange.values;
            await context.sync();
            resolve(columnValues);
        });
    });
}

async function createBoxplotFormulas(sheetName, tableName, colName) {
    await Office.onReady();
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem(sheetName);
        let table = sheet.tables.getItem(tableName);
        const columnRange = table.columns
            .getItem(colName)
            .getDataBodyRange()
            .load("values");
        columnRange.load("address");
        await context.sync();
        let valueRange = columnRange.address;
        
        let minFormula = "=MIN(" + valueRange + ")";
        let lowerQuartile = "=QUARTILE(" + valueRange + "," +"1)";
        let medianFormula = "=MEDIAN(" + valueRange + ")";
        let upperQuartile = "=QUARTILE(" + valueRange + "," + "3)";
        let maxFormula = "=MAX(" + valueRange + ")";
        let rangeFormula = maxFormula + "-MIN(" + valueRange + ")";
        let interquartileFormula = upperQuartile + "-QUARTILE(" + valueRange + "," + "1)";
        let stdDevFormula = "=STDEV.S(" + valueRange + ")";

        const tblrange = table.getRange();
        let range = tblrange.getColumnsAfter(2).getCell(0, 1);
        range.load("address");
        await context.sync();
        let addrFirst = range.address;
        for (let i = 0; i < 7; i++) {
            range.insert(Excel.InsertShiftDirection.right);            
        }        
        range.insert(Excel.InsertShiftDirection.down);
        range.load("address");
        await context.sync();
        let addrLast = range.address;

        let finalRange = sheet.getRange(addrFirst + ':' + addrLast);
        console.log(addrFirst + ':' + addrLast);
        console.log(finalRange);
        finalRange.load("address");
        await context.sync();
        console.log(finalRange.address);

        finalRange.formulas = [
            ['Min', 'Q25', 'Median', 'Q75', 'Max', 'Spannweite', 'IQR', 'St.Dev.'],
            [minFormula, lowerQuartile, medianFormula, upperQuartile, maxFormula, rangeFormula, interquartileFormula, stdDevFormula]
        ];
        finalRange.format.autofitColumns();

        await context.sync();
    });
}

async function deleteWorksheets(deleteAll) {
    await Office.onReady();
    await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        if (sheets.items.length > 1) {
            if (deleteAll){
                for (let i in sheets.items) {
                    if (sheetsCreated.includes(sheets.items[i].name)){
                        const sheet = sheets.items[i];
                        sheet.delete();                        
                        await context.sync();                
                    }
                }
            }else{
                const lastSheet = sheets.items[sheets.items.length - 1];
                lastSheet.delete();
                await context.sync();                
            }
        }
    });
}

async function registerOnActivateHandler(callbackRef) {
    dotNetReference = callbackRef;
    await Office.onReady();
    await Excel.run(async (context) => {
        let sheets = context.workbook.worksheets;
        sheets.onActivated.add(onActivate);

        await context.sync();
        console.log("A handler has been registered for the OnActivate event.");
    });
}

async function onActivate(args) {
    await Excel.run(async (context) => {
        console.log("The activated worksheet Id : " + args.worksheetId);
        await listWorksheets();
    });     
}

async function getTablesFromActiveWorksheet() {
    await Office.onReady();
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("items/name");
        await context.sync();
        const tables = sheet.tables;
        tables.load('name, count, headers, columns');
        await context.sync();
        let allTables = [];
        for (let j in tables.items) {
            let tableheaders = tables.items[j].columns.items;
            let alltableheaders = [];
            for (let k in tableheaders) {
                alltableheaders.push(tableheaders[k].name);
            }
            allTables.push({tablename: tables.items[j].name, categories: alltableheaders});
        }
        await dotNetReference.invokeMethodAsync("CallbackAllTablesInActiveWorksheet", allTables, sheet.name);
    });
}

