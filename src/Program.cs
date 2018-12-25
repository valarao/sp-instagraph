using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace pleaseWork
{
    class Program
    {
        const double MAX_SCALE_FACTOR = 1.05; // Adjusts how high the top of chart is from max price
        const double MIN_SCALE_FACTOR = 0.95; // Adjusts how low the bottom of chart is from min price

        /// <summary>
        /// Run application to process CSV/XLS/XLSX file.
        /// </summary>
        static void Main(string[] args)
        {
            processFile();
        }

        /// <summary>
        /// Deletes unused columns and constructs formatted line chart given company's historical stock prices.
        /// </summary>
        static void processFile() {
            // Declare common sheet variables 
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook book = excel.Workbooks.Open("C:\\Temp\\mySheety14.xlsx");
            Microsoft.Office.Interop.Excel.Worksheet sheet = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            object misValue = System.Reflection.Missing.Value;
            Excel.Range userRange = sheet.UsedRange;
            int rowCount = userRange.Rows.Count - 3;

            // Execute process
            deleteColumns(sheet);
            formatArrays(sheet, rowCount);
            makeChart(sheet, misValue, rowCount);
            defaultView(sheet, book, excel);

        }

        /// <summary>
        /// Deletes the 4 unused columns: Open, Low, High, Close
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        static void deleteColumns(Excel.Worksheet sheet) {
            for (int i = 0; i < 4; i++)
            {
                Excel.Range toBeDeletedRange = sheet.get_Range("B:B", System.Type.Missing);
                toBeDeletedRange.EntireColumn.Delete();
            }
            Excel.Range volRange = sheet.get_Range("C:C", System.Type.Missing);
            volRange.EntireColumn.Delete();
        }

        /// <summary>
        /// Format column width and number formats.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="rowCount">Number of rows filled with data.</param>
        static void formatArrays(Excel.Worksheet sheet, int rowCount) {
            Excel.Range dateRange = sheet.get_Range("A:A", System.Type.Missing);
            dateRange.EntireColumn.ColumnWidth = 12.21;
            string dateFormat = "m/d/yyyy";
            string priceFormat = "_($* 0.00_);_($* (0.00);_($* '-'??_);_(@_)";
            for (int i = 1; i <= rowCount; i++)
                sheet.Cells[i, 1].NumberFormat = dateFormat;
            for (int i = 1; i <= rowCount; i++)
                sheet.Cells[i, 2].NumberFormat = priceFormat;
        }

        /// <summary>
        /// Plot line chart with date and adjusted close prices.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="misValue">Object in case of mishandled value.</param>
        /// <param name="rowCount">Number of rows filled with data.</param>
        static void makeChart(Excel.Worksheet sheet, object misValue, int rowCount) {
            Excel.ChartObjects xlCharts = (Excel.ChartObjects)sheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(150, 20, 300, 250);
            Excel.Chart chartPage = myChart.Chart;

            // Classify Cells
            string firstPriceCell = "B2";
            string lastPriceCell = "B" + rowCount;
            string firstDateCell = "A2";
            string lastDateCell = "A" + rowCount;

            formatChart(sheet, chartPage, firstPriceCell, lastPriceCell, firstDateCell, lastDateCell, misValue, rowCount);
        }
        /// <summary>
        /// Perform chart formatting based on WestPeak 2018-2019 guidelines.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="chartPage">1-year historical price chart object.</param>
        /// <param name="firstPriceCell">First adjusted price in worksheet.</param>
        /// <param name="lastPriceCell">Last adjusted price in worksheet.</param>
        /// <param name="firstDateCell">First date in worksheet.</param>
        /// <param name="lastDateCell">Last date in worksheet.</param>
        /// <param name="misValue">Object in case of mishandled value.</param>
        /// <param name="rowCount">Number of rows filled with data.</param>
        static void formatChart(Excel.Worksheet sheet, Excel.Chart chartPage, string firstPriceCell, string lastPriceCell,
                                string firstDateCell, string lastDateCell, object misValue, int rowCount) {
            chartPage.SetSourceData(sheet.get_Range(firstPriceCell, lastPriceCell), misValue); // Set Y-Axis (Price)
            chartPage.ChartArea.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse; // Delete chart area fill
            chartPage.PlotArea.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse; // Delete plot area fill
            chartPage.ChartArea.Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone; // Delete chart border
            chartPage.ChartType = Excel.XlChartType.xlLine; // Convert to line chart
            chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).TickLabels.NumberFormat = choosePriceFormat(sheet, rowCount); // Set Y-Axis Format (Price)
            chartPage.SeriesCollection(1).XValues = sheet.get_Range(firstDateCell, lastDateCell); // Set X-Axis (Date)
            chartPage.Axes(Excel.XlAxisGroup.xlPrimary).MajorUnit = 2; // Set date unit frequency
            chartPage.Axes(Excel.XlAxisGroup.xlPrimary).TickLabels.NumberFormat = "[$-en-US]mmm-yyyy;@"; // Set X-Axis Format (Date)
            chartPage.Legend.LegendEntries(1).LegendKey.Border.ColorIndex = 1; // Change line color to black
            chartPage.Legend.Delete(); // Delete legend
            chartPage.ChartTitle.Delete(); // Delete chart title
            adjustPriceScale(sheet, rowCount, chartPage); // Adjust price scale
        }

        /// <summary>
        /// Choose price number format depending on the minimum price value.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="rowCount">Number of rows filled with data.</param>
        /// <returns></returns>
        static string choosePriceFormat(Excel.Worksheet sheet, int rowCount)
        {
            double min = findMin(sheet, rowCount);
            string format;
            if (min > 10.0) {
                format = "_($* 0_);_($* (0);_($* '-'??_);_(@_)";
            }
            else {
                format = "_($* 0.00_);_($* (0.00);_($* '-'??_);_(@_)";
            }
            return format;
        }

        /// <summary>
        /// Adjust the maximum and minimum values for the y-axis price scale.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="rowCount">Number of rows filled with data.</param>
        /// <param name="chartPage">1-year historical price chart object.</param>
        static void adjustPriceScale(Excel.Worksheet sheet, int rowCount, Excel.Chart chartPage)
        {
            double graphMax = findMax(sheet, rowCount);
            double graphMin = findMin(sheet, rowCount);

            var yAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
            yAxis.MaximumScale = graphMax;
            yAxis.MinimumScale = graphMin;
        }

        /// <summary>
        /// Find maximum adjusted price over the past year.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="rowCount">Number of rows filled with data.</param>
        /// <returns>Maximum adjusted price over the past year.</returns>
        static double findMax(Excel.Worksheet sheet, int rowCount)
        {
            double maxPrice = 0;
            for (int i = 2; i < rowCount; i++)
            {
                var cellValueStr = (double)(sheet.Cells[i, 2] as Excel.Range).Value;
                double cellValue = Convert.ToDouble(cellValueStr);

                if (cellValue > maxPrice)
                    maxPrice = cellValue;
            }
            return Math.Ceiling(maxPrice * MAX_SCALE_FACTOR);
        }

        /// <summary>
        /// Find minimum adjusted price over the past year.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="rowCount">Number of rows filled with data.</param>
        /// <returns>Minimum adjusted price over the past year.</returns>
        static double findMin(Excel.Worksheet sheet, int rowCount) {
            double minPrice = 1000000;
            for (int i = 2; i < rowCount; i++)
            {
                var cellValueStr = (double)(sheet.Cells[i, 2] as Excel.Range).Value;
                double cellValue = Convert.ToDouble(cellValueStr);

                if (cellValue < minPrice)
                    minPrice = cellValue;
            }
            return Math.Floor(minPrice * MIN_SCALE_FACTOR);
        }

        /// <summary>
        /// Return view to top-left of the page, close the Workbook, and close Excel application.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="book">Excel workbook with pricing data.</param>
        /// <param name="excel">Excel application.</param>
        static void defaultView(Excel.Worksheet sheet, Excel.Workbook book, Excel.Application excel) {
            sheet.Cells[1, 1].Select();
            book.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
        }

    }
}
