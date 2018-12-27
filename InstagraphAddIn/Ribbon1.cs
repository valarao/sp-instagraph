using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using HtmlAgilityPack;
using System.Net.Http;
using System.Collections;
using System.Threading.Tasks;

namespace InstagraphAddIn
{
   

    public partial class Ribbon1
    {
        const double MAX_SCALE_FACTOR = 1.05; // Adjusts how high the top of chart is from max price
        const double MIN_SCALE_FACTOR = 0.95; // Adjusts how low the bottom of chart is from min price
                                              // const string COMPANY_TICKER = "BBW";

        // Application excel = new Application();

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private async void btnIntagraph_Click(object sender, RibbonControlEventArgs e)
        {
            Application excel = new Application();
            Worksheet sheet = Globals.ThisAddIn.GetActiveWorksheet();
            string company = excel.Application.InputBox("What is the company's ticker?").ToUpper();
            string exchange = excel.Application.InputBox("What exchange is it on? (NYSE, NASDAQ, TO, V)").ToUpper();
            exchange = checkExchange(exchange);
            setTitles(sheet);
            await parseData(sheet, company, exchange);
            processFile(sheet, company);

        }

        private static string checkExchange(string exchange) {
            if (exchange.Equals("V") || exchange.Equals("TO"))
            {
                return exchange;
            }
            else {
                return "";
            }
        }

        private void setTitles(Worksheet sheet) {
            sheet.get_Range("A:G", Type.Missing).Clear();
            sheet.get_Range("A1:A1", Type.Missing).Value = "Date";
            sheet.get_Range("B1:B1", Type.Missing).Value = "Open";
            sheet.get_Range("C1:C1", Type.Missing).Value = "High";
            sheet.get_Range("D1:D1", Type.Missing).Value = "Low";
            sheet.get_Range("E1:E1", Type.Missing).Value = "Close";
            sheet.get_Range("F1:F1", Type.Missing).Value = "Adj Close";
            sheet.get_Range("G1:G1", Type.Missing).Value = "Volume";
        }

        private async Task parseData(Worksheet sheet, string company, string exchange) {
            int rowCount = 0;
            var today = convertToUnix(0);
            var threeMonthsPrior = convertToUnix(3);
            var sixMonthsPrior = convertToUnix(6);
            var nineMonthsPrior = convertToUnix(9);
            var twelveMonthsPrior = convertToUnix(12);
            rowCount = sheet.UsedRange.Rows.Count;
            await getAllHTMLData(sheet, company, exchange, today, threeMonthsPrior, sixMonthsPrior,
                nineMonthsPrior, twelveMonthsPrior, rowCount);
            rowCount = sheet.UsedRange.Rows.Count;

            Console.ReadLine();
        }

        private static async Task getAllHTMLData(Worksheet sheet, string company, string exchange, int today,
            int threeMonthsPrior, int sixMonthsPrior, int nineMonthsPrior, int twelveMonthsPrior, int rowCount)
        {
            await getHTMLData(sheet, company, exchange, threeMonthsPrior, today, rowCount); // Run quarterly
            rowCount = sheet.UsedRange.Rows.Count;
            await getHTMLData(sheet, company, exchange, sixMonthsPrior, threeMonthsPrior, rowCount); // Run quarterly
            rowCount = sheet.UsedRange.Rows.Count;
            await getHTMLData(sheet, company, exchange, nineMonthsPrior, sixMonthsPrior, rowCount); // Run quarterly
            rowCount = sheet.UsedRange.Rows.Count;
            await getHTMLData(sheet, company, exchange, twelveMonthsPrior, nineMonthsPrior, rowCount); // Run quarterly
        }

        private static Int32 convertToUnix(int timePeriod)
        {
            var standardTime = DateTime.Today.AddMonths(-timePeriod);
            standardTime = standardTime.AddDays(1);
            var unix = (Int32)standardTime.Subtract(new DateTime(1970, 1, 1)).TotalSeconds;
            return unix;
        }

        private static async Task getHTMLData(Worksheet sheet, string company, string exchange, Int32 startDate, Int32 endDate, int rowCount)
        {
            var url = "https://ca.finance.yahoo.com/quote/" + company + "." + exchange + "/history?period1=" + startDate +
                "&period2=" + endDate + "&interval=1d&filter=history&frequency=1d";
            var httpClient = new HttpClient();
            var html = await httpClient.GetStringAsync(url);

            var htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(html);
            var currRowCount = createMatrix(sheet, htmlDocument, rowCount);
        }

        private static int createMatrix(Worksheet sheet, HtmlDocument htmlDocument, int rowCount)
        {
            var dataHtml = htmlDocument.DocumentNode.Descendants("table")
                .Where(node => node.GetAttributeValue("data-test", "")
                .Equals("historical-prices")).ToList();

            var dataTable = dataHtml[0].ChildNodes[1];
            var dataListItems = dataTable.Descendants("tr").ToList();
            var count = dataListItems.Count;

            for (int i = count - 1; i > 0; i--) {
                if (dataListItems[i].ChildNodes.Count != 7) {
                    dataListItems.Remove(dataListItems[i]);
                }
            }

            var rows = new ArrayList();
            var endPoint = dataListItems.Count;
            for (int i = 0; i < endPoint; i++)
            {
                var row = new ArrayList(); // matrix
                for (int j = 0; j < 7; j++)
                {
                    var dataValue = dataListItems[i].ChildNodes[j].ChildNodes[0].ChildNodes[0].InnerHtml;
                    row.Add(dataValue);
                    Range refCell = (sheet.Cells[i + rowCount + 1, j + 1] as Range);
                    refCell.Value = dataValue;
                }
                
                rows.Add(row);
            };
            return rows.Count;
        }

        static void processFile(Worksheet sheet, string company)
        {
            object misValue = System.Reflection.Missing.Value;
            Range userRange = sheet.UsedRange;
            int rowCount = userRange.Rows.Count;

            // Execute process
            formatDataBackground(sheet, rowCount);
            formatHeaders(sheet);
            formatDataArrays(sheet, rowCount);
            formatDataTitles(sheet);
            formatSummaryBox(sheet);
            makeChart(sheet, misValue, rowCount);
            setSummaryBoxValues(sheet, rowCount, company);
            formatSummaryBoxValues(sheet);
        }

        /// <summary>
        /// Format the used area to have a white background.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="rowCount">Number of rows filled with data.</param>
        static void formatDataBackground(Worksheet sheet, int rowCount)
        {
            int stopPoint = rowCount + 1;
            string range = "A1" + ":S" + stopPoint;
            Range bgRange = sheet.get_Range(range, Type.Missing);
            bgRange.Interior.Color = XlRgbColor.rgbWhite;
        }

        /// <summary>
        /// Format the Instagraph title headers.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        static void formatHeaders(Worksheet sheet)
        {
            Range totalHeaderRange = sheet.get_Range("I2:R5", Type.Missing);
            totalHeaderRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThick);

            Range headerRange = sheet.get_Range("I2:R4", Type.Missing);
            headerRange.Merge();
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            headerRange.Font.Color = XlRgbColor.rgbWhite;
            headerRange.Interior.Color = XlRgbColor.rgbBlack;
            headerRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            headerRange.Value = "Stock Price Instagraph";
            headerRange.Font.Size = 20;

            Range subheaderRange = sheet.get_Range("I5:R5", Type.Missing);
            subheaderRange.Merge();
            subheaderRange.Font.Italic = true;
            subheaderRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            subheaderRange.Font.Size = 10;
            subheaderRange.Value = "For finance nerds too broke to afford Bloomberg/CapIQ " +
                "or too lazy to format an Excel chart themselves.";
        }

        /// <summary>
        /// Format the data title headers: Date, Open, Low, High, Close, Adjusted Close, Volume.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        static void formatDataTitles(Worksheet sheet)
        {
            Range titleRange = sheet.get_Range("A1:G1", Type.Missing);
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRange.Font.Bold = true;
            titleRange.Interior.Color = XlRgbColor.rgbBlack;
            titleRange.Font.Color = XlRgbColor.rgbWhite;
        }

        /// <summary>
        /// Format the width, background color, and number format of data cells.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="rowCount">Number of rows filled with data.</param>
        static void formatDataArrays(Worksheet sheet, int rowCount)
        {
            formatArray(sheet, rowCount, "A", 12.21, System.Drawing.Color.FromArgb(250, 250, 250), "date"); // date
            formatArray(sheet, rowCount, "B", 10.00, System.Drawing.Color.FromArgb(255, 255, 255), "price"); // open
            formatArray(sheet, rowCount, "C", 10.00, System.Drawing.Color.FromArgb(230, 241, 223), "price"); // high
            formatArray(sheet, rowCount, "D", 10.00, System.Drawing.Color.FromArgb(255, 185, 185), "price"); // low
            formatArray(sheet, rowCount, "E", 10.00, System.Drawing.Color.FromArgb(221, 235, 247), "price"); // close
            formatArray(sheet, rowCount, "F", 10.00, System.Drawing.Color.FromArgb(217, 225, 242), "price"); // adjusted close
            formatArray(sheet, rowCount, "G", 10.50, System.Drawing.Color.FromArgb(255, 242, 204), "thousand"); // volume
        }

        /// <summary>
        /// Format specific array based on given column, color, and number format type.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="rowCount">Number of rows filled with data.</param>
        static void formatArray(Worksheet sheet, int rowCount, string col, double colWidth, System.Drawing.Color color, string type)
        {
            string range = col + "2:" + col + rowCount;
            Range arrayRange = sheet.get_Range(range, Type.Missing);
            arrayRange.Interior.Color = color;
            arrayRange.EntireColumn.ColumnWidth = colWidth;
            string format;
            if (type.Equals("date"))
            {
                format = "m/d/yyyy";
            }
            else if (type.Equals("price"))
            {
                format = "_($* 0.00_);_($* (0.00);_($* '-'??_);_(@_)";
            }
            else
            {
                format = "_(* #,##0_);_(* (#,##0);_(* ' - '??_);_(@_)";
            }
            arrayRange.NumberFormat = format;
        }

        /// <summary>
        /// Format the base layout for the summary box.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        static void formatSummaryBox(Worksheet sheet)
        {
            formatSummaryBoxColumns(sheet);
            drawSummaryBox(sheet);
            formatSummaryBoxTitle(sheet);
            setSummaryBoxDataTitles(sheet);
        }

        /// <summary>
        /// Format the non-default widths of data-filled summary box columns.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        static void formatSummaryBoxColumns(Worksheet sheet)
        {
            sheet.get_Range("H:H", Type.Missing).EntireColumn.ColumnWidth = 2.58;
            sheet.get_Range("I:I", Type.Missing).EntireColumn.ColumnWidth = 2.58;
            sheet.get_Range("R:R", Type.Missing).EntireColumn.ColumnWidth = 2.58;
            sheet.get_Range("K:K", Type.Missing).EntireColumn.ColumnWidth = 10.25;
        }

        /// <summary>
        /// Make the border for the summary box.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        static void drawSummaryBox(Worksheet sheet)
        {
            Range boxRange = sheet.get_Range("I7:R25", Type.Missing);
            boxRange.Interior.Color = System.Drawing.Color.FromArgb(250, 250, 250);
            boxRange.BorderAround2(XlLineStyle.xlDash, XlBorderWeight.xlThick);
        }

        /// <summary>
        /// Format the main title of the summary box.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        static void formatSummaryBoxTitle(Worksheet sheet)
        {
            Range titleRange = sheet.get_Range("J8:Q8", Type.Missing);
            titleRange.Merge();
            titleRange.Font.Size = 15;
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRange.Font.Underline = true;
            titleRange.Value = "1-Year Historical Price Summary";
            titleRange.EntireRow.RowHeight = 14.40;
        }

        /// <summary>
        /// Set the data section titles of the summary box.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        static void setSummaryBoxDataTitles(Worksheet sheet)
        {
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
        static void makeChart(Worksheet sheet, object misValue, int rowCount)
        {
            ChartObjects xlCharts = (ChartObjects)sheet.ChartObjects(Type.Missing);
            ChartObject myChart = xlCharts.Add(600, 120, 305, 240);
            Chart chartPage = myChart.Chart;

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
        static void formatChart(Worksheet sheet, Chart chartPage, string firstPriceCell, string lastPriceCell,
                                string firstDateCell, string lastDateCell, object misValue, int rowCount)
        {
            chartPage.SetSourceData(sheet.get_Range(firstPriceCell, lastPriceCell), misValue); // Set Y-Axis (Price)
            chartPage.ChartArea.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse; // Delete chart area fill
            chartPage.PlotArea.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse; // Delete plot area fill
            chartPage.ChartArea.Border.LineStyle = XlLineStyle.xlLineStyleNone; // Delete chart border
            chartPage.ChartType = XlChartType.xlLine; // Convert to line chart
            chartPage.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).TickLabels.NumberFormat = choosePriceFormat(sheet, rowCount); // Set Y-Axis Format (Price)
            chartPage.SeriesCollection(1).XValues = sheet.get_Range(firstDateCell, lastDateCell); // Set X-Axis (Date)
            chartPage.Axes(XlAxisGroup.xlPrimary).MajorUnit = 2; // Set date unit frequency
            chartPage.Axes(XlAxisGroup.xlPrimary).TickLabels.NumberFormat = "[$-en-US]mmm-yyyy;@"; // Set X-Axis Format (Date)
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
        /// <returns>Number format for the price chart.</returns>
        static string choosePriceFormat(Worksheet sheet, int rowCount)
        {
            double min = findMin(sheet, rowCount);
            string format;
            if (min > 10.0)
            {
                format = "_($* 0_);_($* (0);_($* '-'??_);_(@_)";
            }
            else
            {
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
        static void adjustPriceScale(Worksheet sheet, int rowCount, Chart chartPage)
        {
            double graphMax = Math.Ceiling(findMax(sheet, rowCount) * MAX_SCALE_FACTOR);
            double graphMin = Math.Floor(findMin(sheet, rowCount) * MIN_SCALE_FACTOR);

            if (graphMin > 10.0)
            {
                graphMax = Math.Round(graphMax / 5) * 5.0;
                graphMin = Math.Round(graphMin / 5) * 5.0;
            }

            var yAxis = (Axis)chartPage.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
            yAxis.MaximumScale = graphMax;
            yAxis.MinimumScale = graphMin;
        }

        /// <summary>
        /// Find maximum adjusted price over the past year.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="rowCount">Number of rows filled with data.</param>
        /// <returns>Maximum adjusted price over the past year.</returns>
        static double findMax(Worksheet sheet, int rowCount)
        {
            double maxPrice = 0;
            for (int i = 2; i < rowCount + 1; i++)
            {
                var cellValueStr = (double)(sheet.Cells[i, 6] as Range).Value;
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
        static double findMin(Worksheet sheet, int rowCount)
        {
            double minPrice = 1000;
            for (int i = 2; i < rowCount + 1; i++)
            {
                var cellValueStr = (double)(sheet.Cells[i, 6] as Range).Value;
                double cellValue = Convert.ToDouble(cellValueStr);

                if (cellValue < minPrice)
                    minPrice = cellValue;
            }
            return minPrice;
        }

        /// <summary>
        /// Set summary box values given company price data.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="rowCount">Number of rows filled with data.</param>
        static void setSummaryBoxValues(Worksheet sheet, int rowCount, string company)
        {
            string lastDate = "A" + rowCount + ":A" + rowCount;
            string lastPrice = "F" + rowCount + ":F" + rowCount;
            sheet.get_Range("K13:K13", Type.Missing).Value = company;
            sheet.get_Range("K14:K14", Type.Missing).Value = sheet.get_Range(lastDate, Type.Missing).Value;
            sheet.get_Range("K16:K16", Type.Missing).Value = sheet.get_Range(lastPrice, Type.Missing).Value;
            sheet.get_Range("K17:K17", Type.Missing).Value = findMax(sheet, rowCount);
            sheet.get_Range("K18:K18", Type.Missing).Value = findMin(sheet, rowCount);
            sheet.get_Range("K20:K20", Type.Missing).Value = Math.Round(findADTV(sheet, rowCount));
        }

        /// <summary>
        /// Find average daily trading volume (ADTV) over the past year.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        /// <param name="rowCount">Number of rows filled with data.</param>
        /// <returns>Average daily trading volume (ADTV) over the past year</returns>
        static double findADTV(Worksheet sheet, int rowCount)
        {
            double sum = 0;
            double count = rowCount - 1;
            for (int i = 2; i < rowCount + 1; i++)
            {
                var cellValueStr = (double)(sheet.Cells[i, 7] as Range).Value;
                double cellValue = Convert.ToDouble(cellValueStr);
                sum += (cellValue / count);
            }
            return sum;
        }

        /// <summary>
        /// Format summary box company cell text alignment and other cell number formats.
        /// </summary>
        /// <param name="sheet">Excel worksheet with pricing data.</param>
        static void formatSummaryBoxValues(Worksheet sheet)
        {
            sheet.get_Range("K13:K13", Type.Missing).HorizontalAlignment = XlHAlign.xlHAlignRight; // Company
            sheet.get_Range("K14:K14", Type.Missing).NumberFormat = "m/d/yyyy"; // Date
            sheet.get_Range("K16:K16", Type.Missing).NumberFormat = "_($* 0.00_);_($* (0.00);_($* '-'??_);_(@_)"; // Last price
            sheet.get_Range("K17:K17", Type.Missing).NumberFormat = "_($* 0.00_);_($* (0.00);_($* '-'??_);_(@_)"; // High
            sheet.get_Range("K18:K18", Type.Missing).NumberFormat = "_($* 0.00_);_($* (0.00);_($* '-'??_);_(@_)"; // Low
            sheet.get_Range("K20:K20", Type.Missing).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ' - '??_);_(@_)"; // ADTV
        }






    }
}
