import { Component } from '@angular/core';
import * as OfficeHelpers from '@microsoft/office-js-helpers';

const template = require('./app.component.html');

const RawDataTableName = "PopulationTable";
const OutputSheetName = "Top 10 Growing Cities";

@Component({
    selector: 'app-home',
    template
})
export default class AppComponent {
    welcomeMessage = 'Welcome';
    //todo: add a button in the task pane to retrieve top 10 city where population exploed like anything
    // the input for that scenario will be the population table.
    async run() {
        try {
            await Excel.run(async context => {
               let orginalTable = context.workbook.tables.getItem(RawDataTableName);

               // get the proxy objects for city, latest data column and earliest data column.
               let nameColumn = orginalTable.columns.getItem("City");
               let latestDataColumn = orginalTable.columns.getItem("07-01-2014 population estimate");
               let earliestDataColumn = orginalTable.columns.getItem("04-01-1990 population estimate");

               // get the values of all the above three columns

               nameColumn.load("values");
               latestDataColumn.load("values");
               earliestDataColumn.load("values");

               await context.sync();

               let citiesData: Array<{name: string, growth: number }> =[];

               // here we are starting from second index as first row is city.
               // office.js expects 0 based indexing whereas interop requires 1 based indexing. 
               for( let i = 1; i< nameColumn.values.length; i++){
                    let name = nameColumn.values[i][0];

                    let pop1990 = earliestDataColumn.values[i][0];
                    let popLatest = earliestDataColumn.values[i][0];

                    if(isNaN(pop1990) || isNaN(popLatest)) {
                        console.log('Skipping "'+ name +'"');
                    }
                    let growth = popLatest - pop1990;
                    citiesData.push({name: name, growth: growth});
               }

               let sorted = citiesData.sort((city1, city2) => {
                 return city2.growth - city1.growth;
               });

               let top10 = sorted.slice(0, 10);

               // to create the report in a new worksheet
               let outputSheet = context.workbook.worksheets.add(OutputSheetName);
               let sheetHeaderTitle = "Population Growth 1990 - 2014";
               let tableCategories = ["Rank", "City", "Population Growth"];

               let reportStartCell = outputSheet.getRange("B2");
               reportStartCell.values = [[sheetHeaderTitle]];

               reportStartCell.format.font.bold = true;
               reportStartCell.format.font.size = 14;
               reportStartCell.getResizedRange(0, tableCategories.length -1).merge();

               let tableHeader = reportStartCell.getOffsetRange(2, 0).getResizedRange(0, tableCategories.length -1 );
               tableHeader.values = [tableCategories];
               let table = outputSheet.tables.add(tableHeader, true);

               for(let i=0; i< top10.length; i++){
                    let cityData = top10[i];
                    table.rows.add(null, [[i+1, cityData.name, cityData.growth]]);
               }

               table.getRange().getEntireColumn().format.autofitColumns();
               table.getDataBodyRange().getLastColumn().numberFormat = [["#,##"]];
               
               
               // to create a chart using the above drawn table

               let fullTableRange = table.getRange();

               let dataRangeForChart = fullTableRange.getColumn(1).getResizedRange(0, 1);

               let chart = outputSheet.charts.add(Excel.ChartType.columnClustered, dataRangeForChart, Excel.ChartSeriesBy.columns);

               chart.title.text = "Population Growth between 1990 and 2014";

               let chartPositionStart = fullTableRange.getLastRow().getOffsetRange(2, 0);
               chart.setPosition(chartPositionStart, chartPositionStart.getOffsetRange(14, 0));

               outputSheet.activate();

               await context.sync();
            });
        } catch (error) {
            console.log(error);
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
    }
}