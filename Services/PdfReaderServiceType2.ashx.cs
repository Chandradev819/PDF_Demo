using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using org.pdfclown.documents;
using org.pdfclown.documents.contents;
using org.pdfclown.documents.contents.objects;
using org.pdfclown.files;
using PDF_Demo.Helper;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Threading;

namespace PDF_Demo.Service
{
    /// <summary>
    /// Summary description for PdfReaderServiceType2
    /// </summary>
    public class PdfReaderServiceType2 : IHttpHandler
    {
        private List<string> _contentList;
        public void ProcessRequest(HttpContext context)
        {
           // string path = GetFilePath();
            string blank = "";
            PdfTextExtractor pdfTextExtractor = new PdfTextExtractor();
            if (context.Request.Files.Count > 0)
            {
                HttpFileCollection files = context.Request.Files;
                foreach (string key in files)
                {
                    HttpPostedFile file = files[key];
                    string fileName = file.FileName;
                    string extension = System.IO.Path.GetExtension(fileName);
                    if (extension == ".pdf")
                    {
                        fileName = context.Server.MapPath("~/Pdf/" + fileName);
                        // string fileName = System.IO.Path.Combine(context.Server.MapPath("~/Pdf"), fileName);
                        file.SaveAs(fileName);
                        _contentList = new List<string>();
                        CreatePdfContent(fileName);
                        var proprietor = _contentList.FindIndex(m => m == "Name of proprietor") + 210;
                        string proprietorName = _contentList[proprietor];

                        var sSN = _contentList.FindIndex(m => m == "number (SSN)");
                        string _ssN1 = _contentList[sSN + 1];
                        string _ssN2 = _contentList[sSN + 2];
                        string _ssN3 = _contentList[sSN + 3];
                        string _ssN4 = _contentList[sSN + 4];
                        string _ssN5 = _contentList[sSN + 5];
                        string _ssN = string.Concat(_ssN1 + _ssN2 + _ssN3 + _ssN4 + _ssN5);

                        var Principal_Crop = _contentList.FindIndex(m => m == "Principal crop or activity") + 220;
                        string PrincipalCrop = _contentList[Principal_Crop];

                        var Code_from_PartIV = _contentList.FindIndex(m => m == "Code_from_PartIV") + 3;
                        string CodefromPartIV = _contentList[Principal_Crop + 1];
                        var Acconting_Method = _contentList.FindIndex(m => m == "Accounting method:");
                        string AccontingMethod = _contentList[Principal_Crop + 2];
                        var EIN = _contentList.FindIndex(m => m == "Employer ID number (EIN) ") + 3;
                        string _EIN = _contentList[Principal_Crop + 3];
                        var a1 = _contentList.FindIndex(m => m == "1a");

                        string E = _contentList[Principal_Crop + 3];
                        string F = _contentList[Principal_Crop + 4];
                        string G = _contentList[Principal_Crop + 5];
                        //string OneA = _contentList[a1+0];
                        //string OneB = _contentList[a1+0];
                        //string OneC = _contentList[a1+0];
                        string OneA = blank;
                        string OneB = blank;
                        string OneC = blank;
                        string two = _contentList[Principal_Crop + 6];
                        string threeA = _contentList[Principal_Crop + 7];
                        string threeB = _contentList[Principal_Crop + 8];
                        string fourA = _contentList[Principal_Crop + 9];
                        string fourB = _contentList[Principal_Crop + 10];
                        var fiveIndex = _contentList.FindIndex(m => m == "5a");
                        string fiveA = blank;//_contentList[fiveIndex + 0];
                        string fiveB = blank;//_contentList[fiveIndex + 0];
                        string fiveC = blank;//_contentList[fiveIndex + 0];
                        var sixIndex = _contentList.FindIndex(m => m == "6a");
                        string sixA = _contentList[sixIndex + 55];
                        string sixB = _contentList[sixIndex + 56];
                        string sixC = blank;// _contentList[sixIndex+0];
                        string sixD = blank;// _contentList[sixIndex+0];
                        string seven = blank;// _contentList[sixIndex+0];
                        string eight = _contentList[sixIndex + 58];
                        string nine = _contentList[sixIndex + 59];
                        string ten = _contentList[sixIndex];
                        string eleven = _contentList[sixIndex + 60];
                        string twelve = blank; // _contentList[a1+0]; 
                        string thirteen = _contentList[sixIndex + 61];
                        string fourteen = _contentList[sixIndex + 62];
                        string fiveteen = blank; //_contentList[a1+0];
                        string sixteen = blank;// _contentList[a1+0];
                        string seventeen = blank;// _contentList[sixIndex + 63];
                        string eighteen = _contentList[sixIndex + 64];
                        string ninteen = _contentList[sixIndex + 65];
                        string twenty = _contentList[sixIndex + 66];
                        var twentyOneAIndex = _contentList.FindIndex(m => m == "21a");
                        string twentyOneA = blank;// _contentList[a1+0];
                        string twentyOneB = _contentList[twentyOneAIndex + 122];
                        string twentyTwo = blank;//_contentList[a1+0];
                        string twentyThree = blank;//_contentList[a1+0];
                        var twentyFourAIndex = _contentList.FindIndex(m => m == "24a");
                        string twentyFourA = blank;//_contentList[a1+0];
                        string twentyFourB = _contentList[twentyFourAIndex + 87];
                        string twentyFive = _contentList[twentyFourAIndex + 88];
                        string twentySix = _contentList[twentyFourAIndex + 89];
                        string twentySeven = blank;// _contentList[a1+0];
                        string twentyEight = _contentList[twentyFourAIndex + 90];
                        string twentyNine = _contentList[twentyFourAIndex + 91];
                        var thirtyIndex = _contentList.FindIndex(m => m == "30");
                        string thirty = _contentList[twentyFourAIndex + 92];
                        string thirtyOne = blank;//_contentList[a1+0] ;
                        var thirtyTwoAIndex = _contentList.FindIndex(m => m == "32a");
                        string thirtyTwoA = _contentList[thirtyTwoAIndex + 311];
                        string thirtyTwoB = _contentList[thirtyTwoAIndex + 313];
                        string thirtyTwoC = _contentList[thirtyTwoAIndex + 316];
                        string thirtyTwoD = _contentList[thirtyTwoAIndex + 318];
                        string thirtyTwoE = _contentList[thirtyTwoAIndex + 320];
                        string thirtyTwoF = _contentList[thirtyTwoAIndex + 323];
                        string thirtyThree = _contentList[thirtyTwoAIndex + 324];
                        string thirtyFour = _contentList[thirtyTwoAIndex + 325];


                        //dumping data in excel file
                        Excel.Application xlApp = new Excel.Application();

                        if (xlApp == null)
                        {
                            context.Response.Write("Excel is not properly installed!!");
                            return;
                        }

                        Excel.Workbook xlWorkBook;
                        Excel.Worksheet xlWorkSheet;
                        object misValue = System.Reflection.Missing.Value;

                        xlWorkBook = xlApp.Workbooks.Add(misValue);
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        xlWorkSheet.Cells[1, 1] = "Proprietor";
                        xlWorkSheet.Cells[1, 2] = "SSN";
                        xlWorkSheet.Cells[1, 3] = "A_Principal_Crop";
                        xlWorkSheet.Cells[1, 4] = "B_Code_from_PartIV";
                        xlWorkSheet.Cells[1, 5] = "C_Acconting Method";
                        xlWorkSheet.Cells[1, 6] = "D_EIN";
                        xlWorkSheet.Cells[1, 7] = "E";
                        xlWorkSheet.Cells[1, 8] = "F";

                        xlWorkSheet.Cells[1, 9] = "G";
                        xlWorkSheet.Cells[1, 10] = "1a";
                        xlWorkSheet.Cells[1, 11] = "1b";
                        xlWorkSheet.Cells[1, 12] = "1c";
                        xlWorkSheet.Cells[1, 13] = "2";
                        xlWorkSheet.Cells[1, 14] = "3a";
                        xlWorkSheet.Cells[1, 15] = "3b";
                        xlWorkSheet.Cells[1, 16] = "4a";
                        xlWorkSheet.Cells[1, 17] = "4b";
                        xlWorkSheet.Cells[1, 18] = "5a";
                        xlWorkSheet.Cells[1, 19] = "5b";
                        xlWorkSheet.Cells[1, 20] = "5c";

                        xlWorkSheet.Cells[1, 21] = "6a";
                        xlWorkSheet.Cells[1, 22] = "6b";
                        xlWorkSheet.Cells[1, 23] = "6c";
                        xlWorkSheet.Cells[1, 24] = "6d";
                        xlWorkSheet.Cells[1, 25] = "7";
                        xlWorkSheet.Cells[1, 26] = "8";

                        xlWorkSheet.Cells[1, 27] = "9";
                        xlWorkSheet.Cells[1, 28] = "10";
                        xlWorkSheet.Cells[1, 29] = "11";
                        xlWorkSheet.Cells[1, 30] = "12";
                        xlWorkSheet.Cells[1, 31] = "13";
                        xlWorkSheet.Cells[1, 32] = "14";

                        xlWorkSheet.Cells[1, 33] = "15";
                        xlWorkSheet.Cells[1, 34] = "16";
                        xlWorkSheet.Cells[1, 35] = "17";

                        xlWorkSheet.Cells[1, 36] = "18";
                        xlWorkSheet.Cells[1, 37] = "19";
                        xlWorkSheet.Cells[1, 38] = "20";
                        xlWorkSheet.Cells[1, 39] = "21a";
                        xlWorkSheet.Cells[1, 40] = "21b";
                        xlWorkSheet.Cells[1, 41] = "22";
                        xlWorkSheet.Cells[1, 42] = "23";
                        xlWorkSheet.Cells[1, 43] = "24a";
                        xlWorkSheet.Cells[1, 44] = "24b";
                        xlWorkSheet.Cells[1, 45] = "25";
                        xlWorkSheet.Cells[1, 46] = "26";
                        xlWorkSheet.Cells[1, 47] = "27";
                        xlWorkSheet.Cells[1, 48] = "28";
                        xlWorkSheet.Cells[1, 49] = "29";
                        xlWorkSheet.Cells[1, 50] = "30";
                        xlWorkSheet.Cells[1, 51] = "31";
                        xlWorkSheet.Cells[1, 52] = "32a";
                        xlWorkSheet.Cells[1, 53] = "32b";
                        xlWorkSheet.Cells[1, 54] = "32c";
                        xlWorkSheet.Cells[1, 55] = "32d";
                        xlWorkSheet.Cells[1, 56] = "32e";
                        xlWorkSheet.Cells[1, 57] = "32f";
                        xlWorkSheet.Cells[1, 58] = "33";
                        xlWorkSheet.Cells[1, 59] = "34";

                        //Filling on Cell

                        xlWorkSheet.Cells[2, 1] = proprietorName;
                        xlWorkSheet.Cells[2, 2] = _ssN;
                        xlWorkSheet.Cells[2, 3] = PrincipalCrop;
                        xlWorkSheet.Cells[2, 4] = CodefromPartIV;
                        xlWorkSheet.Cells[2, 5] = AccontingMethod;
                        xlWorkSheet.Cells[2, 6] = _EIN;
                        xlWorkSheet.Cells[2, 7] = E;
                        xlWorkSheet.Cells[2, 8] = F;

                        //Comodity value in excel
                        xlWorkSheet.Cells[2, 9] = G;
                        xlWorkSheet.Cells[2, 10] = OneA;
                        xlWorkSheet.Cells[2, 11] = OneB;
                        xlWorkSheet.Cells[2, 12] = OneC;
                        xlWorkSheet.Cells[2, 13] = two;
                        xlWorkSheet.Cells[2, 14] = threeA;

                        //Program Elected
                        xlWorkSheet.Cells[2, 15] = threeB;
                        xlWorkSheet.Cells[2, 16] = fourA;
                        xlWorkSheet.Cells[2, 17] = fourB;
                        xlWorkSheet.Cells[2, 18] = fiveA;
                        xlWorkSheet.Cells[2, 19] = fiveB;
                        xlWorkSheet.Cells[2, 20] = fiveC;
                        //Base Acres
                        xlWorkSheet.Cells[2, 21] = sixA;
                        xlWorkSheet.Cells[2, 22] = sixB;
                        xlWorkSheet.Cells[2, 23] = sixC;
                        xlWorkSheet.Cells[2, 24] = sixD;
                        xlWorkSheet.Cells[2, 25] = seven;
                        xlWorkSheet.Cells[2, 26] = eight;

                        //PLC Yield
                        xlWorkSheet.Cells[2, 27] = nine;
                        xlWorkSheet.Cells[2, 28] = ten;
                        xlWorkSheet.Cells[2, 29] = eleven;
                        xlWorkSheet.Cells[2, 30] = twelve;
                        xlWorkSheet.Cells[2, 31] = thirteen;
                        xlWorkSheet.Cells[2, 32] = fourteen;

                        xlWorkSheet.Cells[2, 33] = fiveteen;
                        xlWorkSheet.Cells[2, 34] = sixteen;
                        xlWorkSheet.Cells[2, 35] = seventeen;

                        xlWorkSheet.Cells[2, 36] = eighteen;
                        xlWorkSheet.Cells[2, 37] = ninteen;
                        xlWorkSheet.Cells[2, 38] = twenty;
                        xlWorkSheet.Cells[2, 39] = twentyOneA;
                        xlWorkSheet.Cells[2, 40] = twentyOneB;
                        xlWorkSheet.Cells[2, 41] = twentyTwo;
                        xlWorkSheet.Cells[2, 42] = twentyThree;
                        xlWorkSheet.Cells[2, 43] = twentyFourA;
                        xlWorkSheet.Cells[2, 44] = twentyFourB;
                        xlWorkSheet.Cells[2, 45] = twentyFive;
                        xlWorkSheet.Cells[2, 46] = twentySix;
                        xlWorkSheet.Cells[2, 47] = twentySeven;
                        xlWorkSheet.Cells[2, 48] = twentyEight;
                        xlWorkSheet.Cells[2, 49] = twentyNine;
                        xlWorkSheet.Cells[2, 50] = thirty;
                        xlWorkSheet.Cells[2, 51] = thirtyOne;
                        xlWorkSheet.Cells[2, 52] = thirtyTwoA;
                        xlWorkSheet.Cells[2, 53] = thirtyTwoB;
                        xlWorkSheet.Cells[2, 54] = thirtyTwoC;
                        xlWorkSheet.Cells[2, 55] = thirtyTwoD;
                        xlWorkSheet.Cells[2, 56] = thirtyTwoE;
                        xlWorkSheet.Cells[2, 57] = thirtyTwoF;
                        xlWorkSheet.Cells[2, 58] = thirtyThree;
                        xlWorkSheet.Cells[2, 59] = thirtyFour;

                        xlWorkBook.SaveAs(@"C:\Users\chandradev_ps\Desktop\Input\SchecduleExcelOutput.xlsx", 
                            Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, 
                            Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, 
                            misValue);
                        xlWorkBook.Close(true, misValue, misValue);
                        xlApp.Quit();

                        Marshal.ReleaseComObject(xlWorkSheet);
                        Marshal.ReleaseComObject(xlWorkBook);
                        Marshal.ReleaseComObject(xlApp);
                        //context.Response.ContentType = "text/plain";
                        //context.Response.Write("Excel file created in c drive");
                    }

                    else
                    {
                        context.Response.Write("Please select file to upload");
                    }
                }

                context.Response.ContentType = "text/plain";
                context.Response.Write("Excel file created");

                //context.Response.ContentType = "text/plain";
                //context.Response.Write("File(s) uploaded successfully!");
            }
        }

        public string GetFilePath()
        {
            string selectedPath = "";

            Thread t = new Thread((ThreadStart)(() => {
                FolderBrowserDialog folderDialog = new FolderBrowserDialog();
                folderDialog.RootFolder = System.Environment.SpecialFolder.MyComputer;
                folderDialog.ShowNewFolderButton = true;
                if (folderDialog.ShowDialog() == DialogResult.Cancel)
                    return;

                selectedPath = folderDialog.SelectedPath;
            }));

            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
            Console.WriteLine(selectedPath);
            return selectedPath;

        }
        private void Extract(ContentScanner level)
        {
            if (level == null)
                return;

            while (level.MoveNext())
            {
                var content = level.Current;
                switch (content)
                {
                    case ShowText text:
                        {
                            var font = level.State.Font;
                            _contentList.Add(font.Decode(text.Text));
                            break;
                        }
                    case Text _:
                    case ContainerObject _:
                        Extract(level.ChildLevel);
                        break;
                }
            }
        }
        public void CreatePdfContent(string filePath)
        {
            using (var file = new File(filePath))
            {
                Document document = file.Document;
                foreach (var page in document.Pages)
                {
                    Extract(new ContentScanner(page));
                }
            }
        }
        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}