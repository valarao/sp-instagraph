﻿using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace pleaseWork
{
    class Program
    {
        const double MAX_SCALE_FACTOR = 1.05; // Adjusts how high the top of chart is from max price
        const double MIN_SCALE_FACTOR = 0.95; // Adjusts how low the bottom of chart is from min price

        /// <summary>
        /// Run application to process XLSX file.
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
            Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook book = convertCSVtoXLSX(excel);
            Excel.Worksheet sheet = excel.ActiveSheet as Excel.Worksheet;
            object misValue = System.Reflection.Missing.Value;
            Excel.Range userRange = sheet.UsedRange;
            int rowCount = userRange.Rows.Count;

            // Execute process
            formatDataBackground(sheet, rowCount);
            formatHeaders(sheet);
            formatDataArrays(sheet);
            formatDataTitles(sheet);
            formatSummaryBox(sheet);
            makeChart(sheet, misValue, rowCount);
            setSummaryBoxValues(sheet, rowCount);
            formatSummaryBoxNumbers(sheet);
            defaultView(sheet, book, excel);
        }

        static Excel.Workbook convertCSVtoXLSX(Excel.Application excel) {
            Excel.Workbook bookCSV = excel.Workbooks.Open("C:\\Temp\\AAPL2.csv");
            bookCSV.SaveAs("C:\\Temp\\AAPLFormatted11.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            bookCSV.Close();
            return excel.Workbooks.Open("C:\\Temp\\AAPLFormatted11.xlsx");
        }

        static void formatDataBackground(Excel.Worksheet sheet, int rowCount)
        {
            int stopPoint = rowCount + 1;
            string range = "A1" + ":S" + stopPoint;
            Excel.Range bgRange = sheet.get_Range(range, Type.Missing);
            bgRange.Interior.Color = Excel.XlRgbColor.rgbWhite;
        }

        static void formatHeaders(Excel.Worksheet sheet) {
            Excel.Range totalHeaderRange = sheet.get_Range("I2:R5", Type.Missing);
            totalHeaderRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            Excel.Range headerRange = sheet.get_Range("I2:R4", Type.Missing);
            headerRange.Merge();
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerRange.Font.Color = Excel.XlRgbColor.rgbWhite;
            headerRange.Interior.Color = Excel.XlRgbColor.rgbBlack;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.Value = "Stock Price Instagraph";
            headerRange.Font.Size = 20;

            Excel.Range subheaderRange = sheet.get_Range("I5:R5", Type.Missing);
            subheaderRange.Merge();
            subheaderRange.Font.Italic = true;
            subheaderRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            subheaderRange.Font.Size = 10;
            subheaderRange.Value = "For finance nerds too broke to afford Bloomberg/CapIQ " +
                "or too lazy to format an Excel chart themselves.";
        }

        static void formatDataTitles(Excel.Worksheet sheet) {
            Excel.Range titleRange = sheet.get_Range("A1:G1", Type.Missing);
            titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            titleRange.Font.Bold = true;
            titleRange.Interior.Color = Excel.XlRgbColor.rgbBlack;
            titleRange.Font.Color = Excel.XlRgbColor.rgbWhite;
        }

        /// <summary>
        /// Deletes the 4 unused columns: Open, Low, High, Close
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        static void formatDataArrays(Excel.Worksheet sheet) {
            formatArray(sheet, "A:A", 1, 12.21, "date"); // date
            formatArray(sheet, "B:B", 2, 10.00, "price"); // open
            formatArray(sheet, "C:C", 3, 10.00, "price"); // high
            formatArray(sheet, "D:D", 4, 10.00, "price"); // low
            formatArray(sheet, "E:E", 5, 10.00, "price"); // close
            formatArray(sheet, "F:F", 6, 10.00, "price"); // adjusted close
            formatArray(sheet, "G:G", 7, 10.50, "thousand"); // volume
        }

        /// <summary>
        /// Format column width and number formats.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="rowCount">Number of rows filled with data.</param>
        static void formatArray(Excel.Worksheet sheet, string range, int col, double colWidth, string type) {
            Excel.Range arrayRange = sheet.get_Range(range, Type.Missing);
            arrayRange.EntireColumn.ColumnWidth = colWidth;
            string format;
            if (type.Equals("date")) {
                format = "m/d/yyyy";
            }
            else if (type.Equals("price")) {
                format = "_($* 0.00_);_($* (0.00);_($* '-'??_);_(@_)";
            }
            else {
                format = "_(* #,##0_);_(* (#,##0);_(* ' - '??_);_(@_)";
            }
            arrayRange.NumberFormat = format;
        }

        static void formatSummaryBox(Excel.Worksheet sheet) {
            formatSummaryBoxColumns(sheet);
            drawSummaryBox(sheet);
            formatSummaryBoxTitle(sheet);
            setSummaryBoxDataTitles(sheet);
        }

        static void formatSummaryBoxColumns(Excel.Worksheet sheet) {
            sheet.get_Range("H:H", Type.Missing).EntireColumn.ColumnWidth = 2.58;
            sheet.get_Range("I:I", Type.Missing).EntireColumn.ColumnWidth = 2.58;
            sheet.get_Range("R:R", Type.Missing).EntireColumn.ColumnWidth = 2.58;
            sheet.get_Range("K:K", Type.Missing).EntireColumn.ColumnWidth = 10.25;
        }

        static void drawSummaryBox(Excel.Worksheet sheet) {
            Excel.Range boxRange = sheet.get_Range("I7:R25", Type.Missing);
            boxRange.Interior.Color = System.Drawing.Color.FromArgb(250, 250, 250);
            boxRange.BorderAround2(Excel.XlLineStyle.xlDash, Excel.XlBorderWeight.xlThick);
        }

        static void formatSummaryBoxTitle(Excel.Worksheet sheet)
        {
            Excel.Range titleRange = sheet.get_Range("J8:Q8", Type.Missing);
            titleRange.Merge();
            titleRange.Font.Size = 15;
            titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            titleRange.Font.Underline = true;
            titleRange.Value = "1-Year Historical Price Summary";
            titleRange.EntireRow.RowHeight = 14.40;
        }

        static void setSummaryBoxDataTitles(Excel.Worksheet sheet) {
            sheet.get_Range("J13:J13", Type.Missing).Value = "Company";
            sheet.get_Range("J14:J14", Type.Missing).Value = "Date";
            sheet.get_Range("J16:J16", Type.Missing).Value = "Last Price";
            sheet.get_Range("J17:J17", Type.Missing).Value = "High";
            sheet.get_Range("J18:J18", Type.Missing).Value = "Low";
            sheet.get_Range("J20:J20", Type.Missing).Value = "ADTV";
        }

        /// <summary>
        /// Plot line chart with date and adjusted close prices.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="misValue">Object in case of mishandled value.</param>
        /// <param name="rowCount">Number of rows filled with data.</param>
        static void makeChart(Excel.Worksheet sheet, object misValue, int rowCount) {
            Excel.ChartObjects xlCharts = (Excel.ChartObjects)sheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = xlCharts.Add(600, 120, 305, 240);
            Excel.Chart chartPage = myChart.Chart;

            // Classify Cells
            string firstPriceCell = "F2";
            string lastPriceCell = "F" + rowCount;
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
            // chartPage.ChartTitle.Delete(); // Delete chart title
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
            double graphMax = Math.Ceiling(findMax(sheet, rowCount) * MAX_SCALE_FACTOR);
            double graphMin = Math.Floor(findMin(sheet, rowCount) * MIN_SCALE_FACTOR);

            if (graphMin > 10.0) {
                graphMax = Math.Round(graphMax / 5) * 5.0;
                graphMin = Math.Round(graphMin / 5) * 5.0;
            }

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
            for (int i = 2; i < rowCount + 1; i++)
            {
                var cellValueStr = (double)(sheet.Cells[i, 6] as Excel.Range).Value;
                double cellValue = Convert.ToDouble(cellValueStr);

                if (cellValue > maxPrice)
                    maxPrice = cellValue;
            }
            return maxPrice;
        }

        /// <summary>
        /// Find minimum adjusted price over the past year.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="rowCount">Number of rows filled with data.</param>
        /// <returns>Minimum adjusted price over the past year.</returns>
        static double findMin(Excel.Worksheet sheet, int rowCount) {
            double minPrice = 1000;
            for (int i = 2; i < rowCount + 1; i++)
            {
                var cellValueStr = (double)(sheet.Cells[i, 6] as Excel.Range).Value;
                double cellValue = Convert.ToDouble(cellValueStr);

                if (cellValue < minPrice)
                    minPrice = cellValue;
            }
            return minPrice;
        }

        static void setSummaryBoxValues(Excel.Worksheet sheet, int rowCount) {
            string lastDate = "A" + rowCount + ":A" + rowCount;
            string lastPrice = "F" + rowCount + ":F" + rowCount;
            sheet.get_Range("K13:K13", Type.Missing).Value = "AAPL";
            sheet.get_Range("K14:K14", Type.Missing).Value = sheet.get_Range(lastDate, Type.Missing).Value;
            sheet.get_Range("K16:K16", Type.Missing).Value = sheet.get_Range(lastPrice, Type.Missing).Value;
            sheet.get_Range("K17:K17", Type.Missing).Value = findMax(sheet, rowCount);
            sheet.get_Range("K18:K18", Type.Missing).Value = findMin(sheet, rowCount);
            sheet.get_Range("K20:K20", Type.Missing).Value = Math.Round(findADTV(sheet, rowCount));
        }

        static double findADTV(Excel.Worksheet sheet, int rowCount)
        {
            double sum = 0;
            double count = rowCount - 1;
            for (int i = 2; i < rowCount + 1; i++)
            {
                var cellValueStr = (double)(sheet.Cells[i, 7] as Excel.Range).Value;
                double cellValue = Convert.ToDouble(cellValueStr);
                sum += (cellValue / count);
            }
            return sum;
        }

        static void formatSummaryBoxNumbers(Excel.Worksheet sheet) {
            sheet.get_Range("K13:K13", Type.Missing).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight; // Company
            sheet.get_Range("K14:K14", Type.Missing).NumberFormat = "m/d/yyyy"; // Date
            sheet.get_Range("K16:K16", Type.Missing).NumberFormat = "_($* 0.00_);_($* (0.00);_($* '-'??_);_(@_)"; // Last price
            sheet.get_Range("K17:K17", Type.Missing).NumberFormat = "_($* 0.00_);_($* (0.00);_($* '-'??_);_(@_)"; // High
            sheet.get_Range("K18:K18", Type.Missing).NumberFormat = "_($* 0.00_);_($* (0.00);_($* '-'??_);_(@_)"; // Low
            sheet.get_Range("K20:K20", Type.Missing).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ' - '??_);_(@_)"; // ADTV
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
