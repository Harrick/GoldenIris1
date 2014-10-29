using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Office.Core;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplicationAutoScan
{
    class StockMap
    {
        public TextBox log;
        public string sSourceCode;  /* The entire source code. */
        public int sort; /* 1:taobao or 2:tmall */

        public string[] sizeCharacters;
        public string[] sizeId;
        public int sizeLength;

        public string[] colourCharacters;
        public string[] colourId;
        public int colourLength;

        public int[,] stock;

        private int sizeStart;
        private int sizeEnd;
        private int colourStart;
        private int colourEnd;

        private string sSizeBlock;
        private string sColourBlock;
        private string sSkuMap;

        private string sSibUrl;
        private string sSibSourceCode;  /* For taobao urls */

        public StockMap(TextBox logBox)
        {
            log = logBox;
            sSourceCode = "";
            sSibSourceCode = "";
            sSibUrl = "";
            sort = 0;
            sizeLength = 0;
            colourLength = 0;
            sizeStart = -1;
            sizeEnd = -1;
            colourStart = -1;
            colourEnd = -1;
        }

        public void StockMapEnd()
        {
            sSourceCode = "";
            sSibSourceCode = "";
            sSibUrl = "";
            sort = 0;
            sizeLength = 0;
            colourLength = 0;
            sizeStart = -1;
            sizeEnd = -1;
            colourStart = -1;
            colourEnd = -1;
            sSizeBlock = "";
            sColourBlock = "";
            sSkuMap = "";
            sizeCharacters = null;
            sizeId = null;
            colourCharacters = null;
            colourId = null;
            stock = null;
        }

        private void countSize()
        {
            MatchCollection mSize;

            try
            {
                if (sSourceCode == null || sSourceCode == "")
                {
                    throw new Exception("Empty excel file.");
                }

                if (sort == 1)
                {
                    sizeStart = sSourceCode.IndexOf("tb-prop J_Prop tb-clearfix J_TMySizeProp");  /* Locating size info */
                }
                else if (sort == 2)
                {
                    sizeStart = sSourceCode.IndexOf("tm-clear J_TSaleProp");  /* Locating size info */
                }

                if (sizeStart == -1)
                {
                    /* Maybe the product is sold out */
                    throw new Exception("Critical: Cannot locate size start.");
                }

                sizeEnd = sSourceCode.IndexOf("</dl>", sizeStart);
                if (sizeStart >= sizeEnd)
                {
                    log.AppendText("Warning: Cannot locate size end.");
                    sizeEnd = sSourceCode.Length;
                }

                sSizeBlock = sSourceCode.Substring(sizeStart, sizeEnd - sizeStart);
                mSize = Regex.Matches(sSizeBlock, "data-value=");
                sizeLength = mSize.Count;
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function countSize" + "\n");
                throw err;
            }
        }

        private void countColour()
        {
            MatchCollection mColour;

            try
            {
                if (sSourceCode == null || sSourceCode == "")
                {
                    throw new Exception("Empty excel file.");
                }

                if (sort == 1)
                {
                    colourStart = sSourceCode.IndexOf("tb-clearfix J_TSaleProp tb-img");  /* Locating colour info */
                }
                else if (sort == 2)
                {
                    colourStart = sSourceCode.IndexOf("tm-clear J_TSaleProp tb-img");  /* Locating colour info */
                }
                if (colourStart == -1)
                {
                    throw new Exception("Critical: Cannot locate colour start.");
                }

                colourEnd = sSourceCode.IndexOf("</dl>", colourStart);
                if (colourStart >= colourEnd)
                {
                    log.AppendText("Warning: Cannot locate colour end.");
                    colourEnd = sSourceCode.Length;
                }

                sColourBlock = sSourceCode.Substring(colourStart, colourEnd - colourStart);
                mColour = Regex.Matches(sColourBlock, "data-value=");
                colourLength = mColour.Count;
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function countColour" + "\n");
                throw err;
            }
        }

        private void constructSizeArray()
        {
            int datastart = 0;
            int dataend = 0;
            string datablock = "";
            int idstart = -1;
            int idend = -1;
            string idblock = "";
            int idblockend = -1;
            int characterstart = -1;
            int characterend = -1;

            /* e.g. <li data-value="20509:28314"><a href="#"><span>S/2码</span></a><i>已选中</i></li>  */

            try
            {
                countSize();
                if (sizeLength <= 0)
                {
                    throw new Exception("Cannot get size info.");
                }
                sizeCharacters = new string[sizeLength];
                sizeId = new string[sizeLength];

                for (int i = 0; i < sizeLength; i++)
                {
                    datastart = sSizeBlock.IndexOf("data-value=", dataend);
                    dataend = sSizeBlock.IndexOf("</li>", datastart);
                    if (datastart >= dataend)
                    {
                        throw new Exception("Cannot locate size data.");
                    }

                    datablock = sSizeBlock.Substring(datastart, dataend - datastart);
                    idblockend = datablock.IndexOf(">");
                    if (idblockend == -1)
                    {
                        throw new Exception("Cannot locate colour id block end.");
                    }
                    idblock = datablock.Substring(0, idblockend);

                    idstart = idblock.IndexOf("\u0022");
                    if (idstart == -1)
                    {
                        throw new Exception("Cannot locate size id start.");
                    }
                    idstart++;
                    idend = idblock.IndexOf("\u0022", idstart);
                    if (idstart >= idend)
                    {
                        throw new Exception("Cannot locate size id end.");
                    }
                    sizeId[i] = idblock.Substring(idstart, idend - idstart);

                    characterstart = datablock.IndexOf("<span>");
                    if (characterstart == -1)
                    {
                        throw new Exception("Cannot locate size character start.");
                    }
                    characterstart += 6;
                    characterend = datablock.IndexOf("</span>", characterstart);
                    if (characterstart >= characterend)
                    {
                        throw new Exception("Cannot locate size character end.");
                    }

                    sizeCharacters[i] = datablock.Substring(characterstart, characterend - characterstart);
                }
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function constructSizeArray" + "\n");
                throw err;
            }
        }

        private void constructColourArray()
        {
            int datastart = 0;
            int dataend = 0;
            string datablock = "";
            int idstart = -1;
            int idend = -1;
            string idblock = "";
            int idblockend = -1;
            int characterstart = -1;
            int characterend = -1;

            /* <li data-value="1627207:3232484" title="兰花《支持验货》" class="tb-txt">
				<a href="#" >
					<span>兰花《支持验货》</span>
				</a>
				<i>已选中</i>			</li>
            */

            try
            {
                countColour();
                if (colourLength <= 0)
                {
                    throw new Exception("Cannot get colour info.");
                }
                colourCharacters = new string[colourLength];
                colourId = new string[colourLength];

                for (int i = 0; i < colourLength; i++)
                {
                    datastart = sColourBlock.IndexOf("data-value=", dataend);
                    dataend = sColourBlock.IndexOf("</li>", datastart);
                    if (datastart >= dataend)
                    {
                        throw new Exception("Cannot locate colour data.");
                    }

                    datablock = sColourBlock.Substring(datastart, dataend - datastart);
                    idblockend = datablock.IndexOf("title");
                    if (idblockend == -1)
                    {
                        throw new Exception("Cannot locate colour id block end.");
                    }
                    idblock = datablock.Substring(0, idblockend);

                    idstart = idblock.IndexOf("\u0022");
                    if (idstart == -1)
                    {
                        throw new Exception("Cannot locate colour id start.");
                    }
                    idstart++;
                    idend = idblock.IndexOf("\u0022", idstart);
                    if (idstart >= idend)
                    {
                        throw new Exception("Cannot locate colour id end.");
                    }
                    colourId[i] = idblock.Substring(idstart, idend - idstart);

                    characterstart = datablock.IndexOf("<span>");
                    if (characterstart == -1)
                    {
                        throw new Exception("Cannot locate colour character start.");
                    }
                    characterstart += 6;
                    characterend = datablock.IndexOf("</span>", characterstart);
                    if (characterstart >= characterend)
                    {
                        throw new Exception("Cannot locate colour character end.");
                    }

                    colourCharacters[i] = datablock.Substring(characterstart, characterend - characterstart);
                }
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function constructColourArray" + "\n");
                throw err;
            }
        }

        private void getSibSourceCode()
        {
            int start = -1;
            int end = -1;

            try
            {
                start = sSourceCode.IndexOf("var b=");
                if (start == -1)
                {
                    throw new Exception("Cannot locate sib start.");
                }
                start = start + 7;

                end = sSourceCode.IndexOf(",a=", start);
                if (start >= end)
                {
                    throw new Exception("Cannot locate sib end.");
                }
                end --;

                sSibUrl = sSourceCode.Substring(start, end - start);
                WebClient objWebClient = new WebClient();
                byte[] arraybDataBuffer = objWebClient.DownloadData(sSibUrl);
                sSibSourceCode = Encoding.Default.GetString(arraybDataBuffer);

            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function getSibSourceCode" + "\n");
                throw err;
            }

            return;
        }

        private void getSkuMap()
        {
            int start = -1;
            int end = -1;
            string sStart = "";
            string sEnd = "";
            string sCode = "";

            try
            {
                if (sort == 1)
                {
                    getSibSourceCode();
                    sCode = sSibSourceCode;
                    sStart = "sku";
                    sEnd = "g_config.";
                }
                else if (sort == 2)
                {
                    sCode = sSourceCode;
                    sStart = "skuMap";
                    sEnd = "cmCatId";
                }
                else
                {
                    throw new Exception("Undefined url sort! sort = "+sort.ToString());
                }

                start = sCode.IndexOf(sStart);
                if (start == -1)
                {
                    throw new Exception("Cannot locate skuMap start.");
                }


                end = sCode.IndexOf(sEnd, start);
                if (start >= end)
                {
                    throw new Exception("Cannot locate skuMap end.");
                }

                sSkuMap = sCode.Substring(start, end - start);
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function getSkuMap" + "\n");
                throw err;
            }
        }

        public int getArrayIndex(string[] array, string element, int length)
        {
            try
            {
                for (int i = 0; i < length; i++)
                {
                    if (element == array[i])
                    {
                        return i;
                    }
                }
                log.AppendText("length= " + length + "\n");
                log.AppendText("element= " + element + "\n");
                for (int j = 0; j < length; j++)
                {
                    log.AppendText(j + " = " + array[j] + "\n");
                }
                throw new Exception("Cannot find element in array.");
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function getArrayIndex" + "\n");
                throw err;
            }
        }

        private void fillStockMap()
        {
            int start = 0;
            int end = 0;
            string size = "";
            string colour = "";
            string stockBlock = "";
            stock = new int[sizeLength, colourLength];
            Regex regex = new Regex(@"\d+");
            Match match;

            try
            {
                while (end != -1)
                {
                    /* size */
                    start = sSkuMap.IndexOf("\u003B", end);
                    if (start == -1)
                    {
                        /* Console.Write("creating sku map may complete."); */
                        break;
                    }
                    start++;
                    end = sSkuMap.IndexOf("\u003B", start);
                    if (start >= end)
                    {
                        throw new Exception("Cannot locate size end in skuMap.");
                    }
                    size = sSkuMap.Substring(start, end - start);

                    /* colour */
                    start = sSkuMap.IndexOf("\u003B", end);
                    if (start == -1)
                    {
                        throw new Exception("Cannot locate colour start in skuMap.");
                    }
                    start++;
                    end = sSkuMap.IndexOf("\u003B", start);
                    if (start >= end)
                    {
                        throw new Exception("Cannot locate colour end in skuMap.");
                    }
                    colour = sSkuMap.Substring(start, end - start);

                    /* stock block */
                    start = sSkuMap.IndexOf("stock", end);
                    if (start == -1)
                    {
                        throw new Exception("Cannot locate stock block start in skuMap.");
                    }
                    end = sSkuMap.IndexOf(",", start);
                    if (end <= start)
                    {
                        stockBlock = sSkuMap.Substring(start);
                    }
                    else
                    {
                        stockBlock = sSkuMap.Substring(start, end - start);
                    }

                    /* stock */
                    match = regex.Match(stockBlock);
                    stock[getArrayIndex(sizeId, size, sizeLength), getArrayIndex(colourId, colour, colourLength)] = System.Int32.Parse(match.Groups[0].ToString());
                }
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function fillStockMap" + "\n");
                throw err;
            }
        }

        private void printStockMap()
        {
            log.AppendText("colourLength = " + colourLength + "\n");
            log.AppendText("sizeLength = " + sizeLength + "\n");
            for (int i = 0; i < colourLength; i++)
            {
                log.AppendText("                                  ");
                log.AppendText(" " + colourCharacters[i]);
            }
            log.AppendText("\n");
            for (int i = 0; i < sizeLength; i++)
            {
                log.AppendText(sizeCharacters[i] + " ");
                for (int j = 0; j < colourLength; j++)
                {
                    log.AppendText(stock[i, j] + " ");
                }
                log.AppendText("\n");
            }
        }

        public bool createStockMap()
        {
            try
            {
                constructSizeArray();
                constructColourArray();
                getSkuMap();
                fillStockMap();
                printStockMap();
                return false;  /* false means normal */
            }
            catch (Exception err)
            {
                log.AppendText("Creating stock map failure" + "\n");
                if (err.Message.Substring(0, 8) != "Critical")
                {
                    throw err;
                }
                else
                {
                    log.AppendText(err.Message + "\n");
                    return true;   /* true means all stocks should be 0 */
                }
            }
        }
    }

    class Stock
    {
        /* Basic information */
        public TextBox log;
        private string sInputFile;  /* Input file path which contains target URLs */
        private string sReportFile;  /* Report file path */
        private string sURL; /* Recording the URLs acquired from input file, utilizing in cycle */
        private int currentRow;  /* current processing row */
        private int iWarningLimit; /* NO USE ----If the stock is under it that would lead to warning */
        private int iLimitPublish;  /* publish if higher than it */
        private int iLimitUnpublish;  /* unpublish if lower than it */
        private string[] aURLs;  /* Restore URLs for loading */
        private bool scanCompleted;  /* is scanning completed? */

        private string currentChineseName;
        private string currentEnglishName;
        private string currentSKU;

        private List<int> reportRows;  /* Restore the rows which will be reported */
        private List<string> reportChineseNames;
        private List<string> reportEnglishNames;
        private List<string> reportSKUs;
        private List<string> reportURLs;

        /* StreamReader objReader;
        public string sOutputFile;
        StreamWriter objWriter; 
        public string sSKUIDline;  
        public string sSKUID; 
        public string sStock; */

        /* Excel infomation */
        Excel.Application scanExcel;
        Excel.Workbook scanWorkbook;
        Excel.Worksheet scanWorksheet; /* input excel worksheet */

        Excel.Application reportExcel;
        Excel.Workbook reportWorkbook;
        Excel.Worksheet reportWorksheet; /* report excel worksheet */

        //Excel.Application loadExcel;
        //Excel.Workbook loadWorkbook;
        //Excel.Worksheet loadWorksheet; /* input load excel worksheet */

        /* Stock Map */
        StockMap stockMap;

        /* public int start;
        public int end;
        public int length; 

        public void readerTxt()
        {
            objReader = new StreamReader(sInputFile);
        }

        public void writterTxt()
        {
            objWriter = new StreamWriter(sOutputFile);
        }
         * */

        private void clearTxt(string sFile)
        {
            FileStream objFileStream = new FileStream(sFile, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            objFileStream.Seek(0, SeekOrigin.Begin);
            objFileStream.SetLength(0);
            objFileStream.Close();
        }

        private bool checkExcelPath(string sExcelPath)
        {
            /* path OK return true */
            string pathext = "";

            if (sExcelPath == null || sExcelPath == "")
            {
                return false;
            }

            try
            {
                pathext = Path.GetExtension(sExcelPath).ToUpper();
                if (pathext == ".XLS" || pathext == ".XLSX")
                {
                    return true;
                }
                else
                {
                    log.AppendText("Unacquainted file type:" + pathext + "\n");
                    return false;
                }
            }
            catch (Exception err)
            {
                log.AppendText(err.Message + "\n");
                return false;
            }
        }

        private void basicElementsInitial(string sInput, string sReport, int limit, TextBox logBox)
        {
            if (checkExcelPath(sInput))
            {
                sInputFile = sInput;
            }
            else
            {
                log.AppendText("Please input the correct scan path" + "\n");
                StockEnd();
                return;
            }

            if (checkExcelPath(sReport))
            {
                sReportFile = sReport;
            }
            else
            {
                sReportFile = "";
                log.AppendText("Warning: no valid report path." + "\n");
            }

            sURL = "";
            scanCompleted = false;
            currentRow = 2;   /* Start row */
            iWarningLimit = limit;   /* It could be modified */
            iLimitPublish = 0;
            iLimitUnpublish = 0;
            log = logBox;
            reportRows = new List<int>();
            reportChineseNames = new List<string>();
            reportEnglishNames = new List<string>();
            reportSKUs = new List<string>();
            reportURLs = new List<string>();
        }

        private void stockInfoInitial()
        {
            stockMap = new StockMap(log);
        }

        private void initScanExcel()
        {
            scanExcel = new Excel.Application();
            scanWorkbook = scanExcel.Workbooks.Open(sInputFile);
            scanWorksheet = (Excel.Worksheet)scanWorkbook.Worksheets[1];
            log.AppendText("Load scan excel successful." + "\n");
        }

        private void initReportExcel()
        {
            reportExcel = new Excel.Application();
            reportWorkbook = reportExcel.Workbooks.Open(sReportFile);
            reportWorksheet = (Excel.Worksheet)reportWorkbook.Worksheets[1];
            log.AppendText("Load report excel successful." + "\n");
        }

        private void loadExcelFunction()
        {
            try
            {
                initScanExcel();

                if (sReportFile != "")
                {
                    initReportExcel();
                }
            }
            catch (Exception err)
            {
                log.AppendText(err.Message + "\n");
                log.AppendText("Loading excel failure." + "\n");
                StockEnd();
                return;
            }
        }

        /* Construction function */
        public Stock(string sInput, string sReport, int limit, TextBox logBox)
        {
            try
            {
                basicElementsInitial(sInput, sReport, limit, logBox);
                stockInfoInitial();
                loadExcelFunction();
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function Stock" + "\n");
                log.AppendText(err.Message + "\n");
                if (sInput != null)
                {
                    StockEnd();
                }
                return;
            }
        }

        /* Killing excel process */
        private void killExcelProcess()
        {
            System.Diagnostics.Process[] myProcesses;

            try
            {
                myProcesses = System.Diagnostics.Process.GetProcessesByName("Excel");
                foreach (System.Diagnostics.Process myProcess in myProcesses)
                {
                    myProcess.Kill();
                }
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function killExcelProcess" + "\n");
                throw err;
            }
        }

        /* Termination function */
        public void StockEnd()
        {
            try
            {
                log = null;
                sInputFile = null;
                sReportFile = null;
                sURL = null;
                aURLs = null;
                stockMap = null;
                scanWorkbook.Save();
                scanWorkbook.Close();
                scanExcel.Quit();
                scanWorkbook = null;
                scanExcel = null;
                scanWorksheet = null;
                reportWorkbook.Save();
                reportWorkbook.Close();
                reportExcel.Quit();
                reportWorkbook = null;
                reportExcel = null;
                reportWorksheet = null;
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function StockEnd" + "\n");
                log.AppendText(err.Message);
                killExcelProcess();
            }
            GC.Collect();
        }

        /*
        private void LoadExcelEnd()
        {
            try
            {
                log = null;
                sInputFile = null;
                sReportFile = null;
                sURL = null;
                aURLs = null;
                stockMap = null;
                loadWorkbook.Save();
                loadWorkbook.Close();
                loadExcel.Quit();
                loadWorkbook = null;
                loadExcel = null;
                loadWorksheet = null;
            }
            catch (Exception err)
            {
                log.AppendText(err.Message);
                killExcelProcess();
            }
            GC.Collect();
        }
         */

        /* Checking the string is a valid URL or the end. If so, return true. */
        private bool analysisURL(string url)
        {
            try
            {
                if (url == "" || url == null)
                {
                    return true;
                }

                if (url.Substring(0, 4) == "http")
                {
                    if (-1 != url.IndexOf("taobao.com/"))
                    {
                        stockMap.sort = 1;
                    }
                    else if (-1 != url.IndexOf("tmall.com/"))
                    {
                        stockMap.sort = 2;
                    }
                    else
                    {
                        log.AppendText("Invalid URL" + "\n");
                        return true;
                    }
                    return false;
                }
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function analysisURL" + "\n");
                log.AppendText("Maybe we encountered the end of the excel." + "\n");
                return true;
            }
            log.AppendText("Debug" + sURL + "\n");
            log.AppendText("Maybe we encountered the end of the excel." + "\n");
            return true;
        }

        private void readUrl()
        {
            Excel.Range range;

            try
            {
                range = scanWorksheet.Cells[currentRow, 12];  /* [row,column] */
                sURL = range.Value;
                if (analysisURL(sURL))
                {
                    scanCompleted = true;
                    return;
                }
                log.AppendText("Reading  " + sURL + "\n");
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function readUrl" + "\n");
                throw err;
            }
        }

        private void getWebResourceFile(string url)
        {
            try
            {
                WebClient objWebClient = new WebClient();
                byte[] arraybDataBuffer = objWebClient.DownloadData(url);
                stockMap.sSourceCode = Encoding.Default.GetString(arraybDataBuffer);

                /*****************************   Debug   **********************************************/
                clearTxt(@"C:\test\debug.txt");
                StreamWriter writer = new StreamWriter(@"C:\test\debug.txt");
                writer.Write(stockMap.sSourceCode);
                writer.Close();
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function getWebResourceFile" + "\n");
                throw (err);
            }
        }

        /*
        public void analysisStock()
        {
            int stock = System.Int32.Parse(sStock);
            if (stock <= 0)
            {
                sStock += " ***NULL STOCK!!!***";
            }
            else if (stock < iWarningLimit)
            {
                sStock += " *Warning*";
            }
            Console.Write(sStock + "\n");
            objWriter.WriteLine(sStock + "\n");
            return;
        }

        
        public void getSKUID(int i)
        {
            sSKUID = sSKUIDline.Substring(i * 12, 11);
            if (sSKUID == "")
            {
                return;
            }
            string si = i.ToString();
            Console.Write(si + " " + sSKUID + "\n");
            objWriter.WriteLine(si + " " + sSKUID + "\n");
            return;
        }

        public void getStartLocation()
        {
            start = sSourceCode.IndexOf(sSKUID);
            if(-1 == start)
            {             
                return;
            }
            Regex regex = new Regex(@"\d+");
            start = sSourceCode.IndexOf("stock", start);
            if(-1 == start)
            {
                return;
            }
            Match match = regex.Match(sSourceCode,start);
            string sFirstNum = match.Groups[0].ToString();
            start = sSourceCode.IndexOf(sFirstNum,start);
            return;
        }

        public void getEndLocation()
        {
            Regex regex = new Regex(@"\D+");
            Match match = regex.Match(sSourceCode,start);
            string sLastNum = match.Groups[0].ToString();
            end = sSourceCode.IndexOf(sLastNum,start);
        }
        public void getStockLength()
        {
            length = end - start;
            return;
        }

        public void recoverVarinCycle()
        {
            sSKUID = "";
            sStock = "";
            start = -1;
            end = -1;
            length = 0;
            return;
        }
         * */

        private bool isCellHaveContent(Excel.Worksheet worksheet, int row, int column)
        {
            try
            {
                Excel.Range range = worksheet.Cells[row, column];  /* [row,column] */
                if (range.Value == null)
                {
                    return false;
                }
                string content = range.Value.ToString();
                if (content == null || content == "")
                {
                    return false;
                }
                return true;
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function isCellHaveContent" + "\n");
                throw err;
            }
        }

        private int getStockByCharacters(string currentSize, string currentColour)
        {
            try
            {
                int sizeIndex = stockMap.getArrayIndex(stockMap.sizeCharacters, currentSize, stockMap.sizeLength);
                int colourIndex = stockMap.getArrayIndex(stockMap.colourCharacters, currentColour, stockMap.colourLength);
                return stockMap.stock[sizeIndex, colourIndex];
            }
            catch
            {
                /* If cannot being found, then we believe the product is out of sell */
                return -1;
            }
        }

        private void setCellColour(Excel.Worksheet worksheet, int stockValue, int row, int column)
        {
            Excel.Range range;
            int none = -4142; /* nothing */
            int yellow = 6;  /* warning */
            int red = 3;   /* zero  */
            int green = 10; /* cannot find */
            int blue = 5;   /* presell */
            int violet = 13; /* limit  */

            try
            {
                range = worksheet.Cells[row, column];  /* [row,column] */
                range.Interior.ColorIndex = none;

                if (stockValue == 0)
                {
                    range.Interior.ColorIndex = red;
                }
                else if (stockValue < 0)
                {
                    range.Interior.ColorIndex = green;
                }
                else if (stockValue <= iLimitUnpublish)
                {
                    range.Interior.ColorIndex = yellow;
                }
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function setCellColour" + "\n");
                throw err;
            }
        }

        private void setLimitUnpublish()
        {
            Excel.Range range;

            try
            {
                range = scanWorksheet.Cells[1, 19];  /* [1,S] for unpublish limit */
                if (range.Value.ToString() == "")
                {
                    iLimitUnpublish = 0;
                }
                else
                {
                    iLimitUnpublish = (int)range.Value;
                }
                log.AppendText("Unpublish limit is " + iLimitUnpublish + "\n");
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function setLimitUnpublish" + "\n");
                throw err;
            }
        }

        private void setLimitPublish()
        {
            Excel.Range range;

            try
            {
                range = scanWorksheet.Cells[1, 20];  /* [1,T] for publish limit */
                if (range.Value.ToString() == "")
                {
                    iLimitPublish = 10000;
                }
                else
                {
                    iLimitPublish = (int)range.Value;
                }
                log.AppendText("Publish limit is " + iLimitPublish + "\n");
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function setLimitPublish" + "\n");
                throw err;
            }
        }

        private void setLimits()
        {
            try
            {
                setLimitUnpublish();
                setLimitPublish();
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function setLimits" + "\n");
                throw err;
            }
        }

        private void reportRow(int stockValue, string sRowType)
        {
            try
            {
                if ((sRowType == "un" && stockValue > iLimitPublish)
                    || (sRowType == "on" && stockValue < iLimitUnpublish)
                    || (sRowType == "nothing"))
                {
                    reportRows.Add(currentRow);
                    reportChineseNames.Add(currentChineseName);
                    reportEnglishNames.Add(currentEnglishName);
                    reportSKUs.Add(currentSKU);
                    reportURLs.Add(sURL);
                }
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function reportRow" + "\n");
                throw err;
            }
        }

        private string getCurrentColour()
        {
            try
            {
                Excel.Range range = scanWorksheet.Cells[currentRow, 13];  /* [row,column] */
                return range.Value;
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function getCurrentColour" + "\n");
                throw err;
            }
        }

        private string getCurrentSize()
        {
            try
            {
                Excel.Range range = scanWorksheet.Cells[currentRow, 14];  /* [row,column] */
                return range.Value;
            }
            catch(Exception err)
            {
                log.AppendText("An exception catched in function getCurrentSize" + "\n");
                throw err;
            }
        }
        private int getStockInCurrentRow(bool needAllZero, string currentSize, string currentColour)
        {
            try
            {
                if (!needAllZero)
                {
                    return getStockByCharacters(currentSize, currentColour);
                }
                else
                {
                    return 0;
                }
            }
            catch(Exception err)
            {
                log.AppendText("An exception catched in function getStockInCurrentRow" + "\n");
                throw err;
            }
        }

        private void fillStockInCurrentRow(bool needAllZero, string currentSize, string currentColour, string sRowType)
        {
            try
            {
                Excel.Range range = scanWorksheet.Cells[currentRow, 15];
                range.Value = getStockInCurrentRow(needAllZero, currentColour, currentSize);
                reportRow((int)range.Value, sRowType);
                setCellColour(scanWorksheet, (int)range.Value, currentRow, 15);
            }
            catch(Exception err)
            {
                log.AppendText("An exception catched in function fillStockInCurrentRow" + "\n");
                throw err;
            }
        }

        private string getRowType()
        {
            try
            {
                Excel.Range range = scanWorksheet.Cells[currentRow, 8];  /* [row,column] */
                return range.Value;
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function getRowType" + "\n");
                throw err;
            }
        }

        private void fillStockIntoScanExcel(bool needAllZero)
        {
            int protect = 1000;
            int blankRow = 10;   /* set blank rows */
            string sRowType = "";

            try
            {
                do
                {
                    sRowType = getRowType();

                    if (sRowType == "skip")
                    {
                        currentRow++;
                        protect++;
                        continue;
                    }

                    if (isCellHaveContent(scanWorksheet, currentRow, 13) 
                        && isCellHaveContent(scanWorksheet, currentRow, 14))
                    {
                        fillStockInCurrentRow(needAllZero, getCurrentColour(), getCurrentSize(), sRowType);
                    }
                    else
                    {
                        blankRow--;
                    }

                    if (blankRow == 0)
                    {
                        scanCompleted = true;
                        break;
                    }

                    currentRow++;
                    protect++;
                } while (!isCellHaveContent(scanWorksheet, currentRow, 12) || protect <= 1000);
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function fillStockIntoScanExcel" + "\n");
                throw err;
            }
        }

        private void clearReportExcel()
        {
            Excel.Range range;

            try
            {
                range = reportWorksheet.Columns["a:z", System.Type.Missing];
                range.Value = "";
                range.Interior.ColorIndex = -4142;
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function clearReportExcel" + "\n");
                throw err;
            }
        }

        /*
        private void clearLoadExcel()
        {
            Excel.Range range;

            try
            {
                range = loadWorksheet.Columns["b:e", System.Type.Missing];
                range.Value = "";
                range.Interior.ColorIndex = -4142;
            }
            catch (Exception err)
            {
                throw err;
            }
        }
         */

        /*
        private void fillStockIntoLoadExcel(string url, bool needAllZero)
        {
            Excel.Range range;

            try
            {
                range = loadWorksheet.Cells[currentRow, 2];
                range.Value = url;

                if (needAllZero)
                {
                    range = loadWorksheet.Cells[currentRow, 3];
                    range.Value = "Nothing";
                    return;
                }

                for (int i = 0; i < stockMap.colourLength; i++)
                {
                    for (int j = 0; j < stockMap.sizeLength; j++)
                    {
                        range = loadWorksheet.Cells[currentRow, 3];
                        range.Value = stockMap.colourCharacters[i];
                        range = loadWorksheet.Cells[currentRow, 4];
                        range.Value = stockMap.sizeCharacters[j];
                        range = loadWorksheet.Cells[currentRow, 5];
                        range.Value = stockMap.stock[j, i];
                        setCellColour(loadWorksheet, stockMap.stock[j, i], currentRow, 5);
                        currentRow++;
                    }
                }
            }
            catch (Exception err)
            {
                throw err;
            }
        }
        */

        private void zeroStockMap()
        {
            int red = 3;   /* zero  */

            try
            {
                Excel.Range range = scanWorksheet.Cells[currentRow, 8];
                range.Value = "nothing";
                range.Interior.ColorIndex = red;
                reportRow(0, "nothing");
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function zeroStockMap" + "\n");
                throw err;
            }
        }

        private void fillCurrentColour(int i)
        {
            try
            {
                Excel.Range range = scanWorksheet.Cells[currentRow, 13];
                range.Value = stockMap.colourCharacters[i];
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function fillCurrentColour" + "\n");
                throw err;
            }
        }

        private void fillCurrentSize(int j)
        {
            try
            {
                Excel.Range range = scanWorksheet.Cells[currentRow, 14];
                range.Value = stockMap.sizeCharacters[j];
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function fillCurrentSize" + "\n");
                throw err;
            }
        }

        private void changeCurrentRowType(string sNewRowType)
        {
            try
            {
                Excel.Range range = scanWorksheet.Cells[currentRow, 8];
                range.Value = sNewRowType;
            }
            catch(Exception err)
            {
                log.AppendText("An exception catched in function changeCurrentRowType" + "\n");
                throw err;
            }
        }

        private void fillCurrentStock(int i, int j)
        {
            try
            {
                Excel.Range range = scanWorksheet.Cells[currentRow, 15];
                range.Value = stockMap.stock[j, i];
                changeCurrentRowType("on");
                reportRow((int)range.Value, getRowType());
                setCellColour(scanWorksheet, stockMap.stock[j, i], currentRow, 15);
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function fillCurrentStock" + "\n");
                throw err;
            }
        }

        private void insertRow(int i, int j)
        {
            try
            {
                if (i > 0 || j > 0)
                {
                    Excel.Range range = scanWorksheet.Rows[currentRow, Type.Missing];
                    range.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Type.Missing);
                }
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function insertRow" + "\n");
                throw err;
            }
        }

        private void fillStockMap()
        {
            try
            {
                for (int i = 0; i < stockMap.colourLength; i++)
                {
                    for (int j = 0; j < stockMap.sizeLength; j++)
                    {
                        insertRow(i, j);
                        fillCurrentColour(i);
                        fillCurrentSize(j);
                        fillCurrentStock(i, j);
                        currentRow++;
                    }
                }
            }
            catch(Exception err)
            {
                log.AppendText("An exception catched in function fillStockMap" + "\n");
                throw err;
            }
        }

        private void loadStockMapIntoScanExcel(bool needAllZero)
        {
            try
            {
                if (needAllZero)
                {
                    zeroStockMap();
                    return;
                }

                fillStockMap();
            }
            catch(Exception err)
            {
                log.AppendText("An exception catched in function loadStockMapIntoScanExcel" + "\n");
                throw err;
            }
        }

        private void clearStock()
        {
            sURL = "";
            stockMap.sSourceCode = "";
            stockMap.sort = 0;
            currentChineseName = "";
            currentEnglishName = "";
            currentSKU = "";
        }

        private void fillReportRow(int i)
        {
            Excel.Range range;
            Excel.Range rangeScan;

            try
            {
                /* Row index */
                range = reportWorksheet.Cells[i + 2, 1];
                range.Value = reportRows[i];
                /* Chinese Name */
                range = reportWorksheet.Cells[i + 2, 2];
                range.Value = reportChineseNames[i];
                /* English Name */
                range = reportWorksheet.Cells[i + 2, 3];
                range.Value = reportEnglishNames[i];
                /* SKU */
                range = reportWorksheet.Cells[i + 2, 4];
                range.Value = reportSKUs[i];
                /* URL */
                range = reportWorksheet.Cells[i + 2, 5];
                range.Value = reportURLs[i];
                /* Row Type */
                rangeScan = scanWorksheet.Cells[reportRows[i],8];
                range = reportWorksheet.Cells[i + 2, 6];
                range.Value = rangeScan.Value;
                /* Colour */
                rangeScan = scanWorksheet.Cells[reportRows[i], 13];
                range = reportWorksheet.Cells[i + 2, 7];
                range.Value = rangeScan.Value;
                /* Size */
                rangeScan = scanWorksheet.Cells[reportRows[i], 14];
                range = reportWorksheet.Cells[i + 2, 8];
                range.Value = rangeScan.Value;
                /* Stock */
                rangeScan = scanWorksheet.Cells[reportRows[i], 15];
                range = reportWorksheet.Cells[i + 2, 9];
                range.Value = rangeScan.Value;
            }
            catch(Exception err)
            {
                log.AppendText("An exception catched in function fillReportRow" + "\n");
                throw err;
            }
        }

        private void fillReportExcel()
        {
            int reportLength = 0;

            try
            {
                if (sReportFile == "")
                {
                    log.AppendText("No report file path." + "\n");
                    return;
                }

                reportLength = reportRows.Count();
                if (reportLength == 0)
                {
                    log.AppendText("No item need to be reported." + "\n");
                    return;
                }

                log.AppendText("\n");
                log.AppendText("Start to report " + sReportFile + "\n");
                log.AppendText("There are " + reportLength + " record(s) to be reported." + "\n");
                clearReportExcel();

                for (int i = 0; i < reportLength; i++)
                {
                    fillReportRow(i);
                }
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function fillReportExcel" + "\n");
                throw err;
            }
        }

        private void recordReportInfo()
        {
            Excel.Range range;

            try
            {
                if(sReportFile == "")
                {
                    return;
                }

                range = scanWorksheet.Cells[currentRow, 1];
                currentChineseName = range.Value;
                if (currentChineseName == null)
                {
                    currentChineseName = "";
                }
                range = scanWorksheet.Cells[currentRow, 2];
                currentEnglishName = range.Value;
                if (currentEnglishName == null)
                {
                    currentEnglishName = "";
                }
                range = scanWorksheet.Cells[currentRow, 3];
                currentSKU = range.Value;
                if (currentSKU == null)
                {
                    currentSKU = "";
                }
            }
            catch(Exception err)
            {
                log.AppendText("An exception catched in function recordReportInfo" + "\n");
                throw err;
            }
        }

        public void scanStock()
        {
            int protect = 10000;  /* A protectiveness for cycle */

            try
            {
                log.AppendText("\n");
                log.AppendText("Start to scan excel " + sInputFile + "\n");

                setLimits();

                for (int i = 0; i < protect; i++)
                {
                    stockInfoInitial();
                    readUrl();  /* reading target URL  */
                    if(scanCompleted)
                    {
                        break;
                    }
                    getWebResourceFile(sURL);  /* getting source code */
                    recordReportInfo();
                    
                    /* creating stock map by utilizing source code  */
                    if (getRowType() == "new")
                    {
                        loadStockMapIntoScanExcel(stockMap.createStockMap()); /* write stock map into target excel */
                    }
                    else
                    {
                        fillStockIntoScanExcel(stockMap.createStockMap());  /* write stock into target Excel */
                        if (scanCompleted)
                        {
                            break;
                        }
                    }
                    clearStock(); /* clear info */
                }

                fillReportExcel();
            }
            catch (Exception err)
            {
                log.AppendText("An exception catched in function scanStock" + "\n");
                log.AppendText(err.Message + "\n");
            }
            finally
            {
                log.AppendText("Scan has completed.");
                StockEnd();
            }
        }

        /*
        private void countURL()
        {
            int column = 1;
            int row = 2;
            int count = 0;
            int blank = 10;
            int fill = 0;
            Excel.Range range;

            try
            {
                for (int i = 0; i < 10000 && blank > 0; i++)
                {
                    if (isCellHaveContent(loadWorksheet, row + i, column))
                    {
                        count++;
                    }
                    else
                    {
                        blank--;
                    }
                }
                if (count > 0)
                {
                    aURLs = new string[count];
                    for (int i = 0; i < 10000 && fill < count; i++)
                    {
                        if (isCellHaveContent(loadWorksheet, row + i, column))
                        {
                            range = loadWorksheet.Cells[row + i, column];
                            aURLs[fill] = range.Value;
                            fill++;
                        }
                    }
                }
                else
                {
                    throw new Exception("empty loading excel.");
                }
            }
            catch (Exception err)
            {
                throw err;
            }
        }
         */

        /*
        public void loadScan()
        {
            try
            {
                log.AppendText("\n");
                log.AppendText("Start to load excel " + sInputLoadFile + "\n");
                clearLoadExcel();
                countURL();
                for (int i = 0; i < aURLs.Length; i++)
                {
                    if (!analysisURL(aURLs[i]))
                    {
                        getWebResourceFile(aURLs[i]);
                        // creating stock map by utilizing source code 
                        fillStockIntoLoadExcel(aURLs[i],stockMap.createStockMap());
                        clearStock(); // clear info
                    }
                }
                currentRow = 2;
                return;
            }
            catch (Exception err)
            {
                currentRow = 2;
                log.AppendText(err.Message + "\n");
            }
            finally
            {
                log.AppendText("Load has completed.");
                LoadExcelEnd();
            }
        }
         */
    }
}
