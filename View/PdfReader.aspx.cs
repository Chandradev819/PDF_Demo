using org.pdfclown.documents;
using org.pdfclown.documents.contents;
using org.pdfclown.documents.contents.objects;
using org.pdfclown.files;
using PDF_Demo.Helper;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace PDF_Demo.View
{
    public partial class PdfReader : System.Web.UI.Page
    {
        private List<string> _contentList;
        protected void Page_Load(object sender, EventArgs e)
        {
        }
        protected void Button1_Click(object sender, EventArgs e)
        {
            PdfTextExtractor pdfTextExtractor = new PdfTextExtractor();
            if (FileUpload1.HasFile)
            {
                string fileName = System.IO.Path.Combine(Server.MapPath("~/Pdf"), FileUpload1.FileName);
                FileUpload1.SaveAs(fileName);
                _contentList = new List<string>();
                CreatePdfContent(fileName);
                var indexProg = _contentList.FindIndex(m => m == "1. Program Year: ") + 1;
                int.TryParse(_contentList[indexProg], out var ProgValue);

                var indexState = _contentList.FindIndex(m => m == "2. State Code") + 3;
                int.TryParse(_contentList[indexState], out var StateValue);

                var indexCountry = _contentList.FindIndex(m => m == "3. County Code") + 3;
                int.TryParse(_contentList[indexCountry], out var CountryValue);

                var indexFarm = _contentList.FindIndex(m => m == "4. Farm Number") + 3;
                int.TryParse(_contentList[indexFarm], out var FarmValue);

                var indexFSAOffice = _contentList.FindIndex(m => m == "5A. County FSA Office Name and Address") + 1;
                string FSAOfficeValue1 = _contentList[indexFSAOffice];
                string FSAOfficeValue2 = _contentList[indexFSAOffice + 1];
                string FSAOfficeValue3 = _contentList[indexFSAOffice + 2];
                string FSAOfficeValue = string.Concat(FSAOfficeValue1, FSAOfficeValue2, FSAOfficeValue3);

                var indexCountryOffice = _contentList.FindIndex(m => m == "5B. County Office Telephone No") + 4;
                string CountryOfficeValue = _contentList[indexCountryOffice];

                var indexCountryFax = _contentList.FindIndex(m => m == "5C. County Office Fax No") + 3;
                string CountryFaxValue = _contentList[indexCountryFax];

                var indexMultiYearContract = _contentList.FindIndex(m => m == "6.  Multi-year Contract ");
                //string MultiYearContractValue = _contentList[indexMultiYearContract];
                string MultiYearContractValue = string.Empty;

                var indexOwnerProducer1 = _contentList.FindIndex(m => m == "12A. Owner or Producer's Name and Address") + 1;
                string ownerProducerValue1 = _contentList[indexOwnerProducer1];
                string ownerProducerValue2 = _contentList[indexOwnerProducer1 + 1];
                string ownerProducerValue3 = _contentList[indexOwnerProducer1 + 2];
                string ownerProducerValue = string.Concat(ownerProducerValue1, ownerProducerValue2, ownerProducerValue3);

                var indexEmailId = _contentList.FindIndex(m => m == "12B. Email Address") + 1;
                //string emailvalue = _contentList[indexEmailId];
                string emailvalue = string.Empty;

                var indexTelephoneNum = _contentList.FindIndex(m => m == "12C. Telephone No. ") + 1;
                //string telePhoneNum= _contentList[indexTelephoneNum];
                string telePhoneNum = string.Empty;

                //For Comodity
                var indexComodity = _contentList.FindIndex(m => m == "Commodity");
                string cornValue = _contentList[indexComodity + 25];
                string ricelongGrainValue = _contentList[indexComodity + 69];
                string seedcottonValue = _contentList[indexComodity + 43];
                string grainsorghumValue = _contentList[indexComodity + 68];
                string soyabeansValue = _contentList[indexComodity + 71];
                string wheatValue = _contentList[indexComodity + 74];

                //For Program Elected
                var indexProgElected = _contentList.FindIndex(m => m == "Elected");
                string plcValue = _contentList[indexProgElected + 23];
                string arcCountyValue = _contentList[indexProgElected + 37];

                //Base Acres
                var indexBaseAcres = _contentList.FindIndex(m => m == "Base Acres");
                string value_643 = _contentList[indexBaseAcres + 22];
                string value_336 = _contentList[indexBaseAcres + 32];
                string value_1052 = _contentList[indexBaseAcres + 40];
                string value_27 = _contentList[indexBaseAcres + 27];
                string value_1853 = _contentList[indexBaseAcres + 36];
                string value_387 = _contentList[indexBaseAcres + 44];

                //PLC Yield
                var indexPLCYield = _contentList.FindIndex(m => m == "PLC Yield");
                string value_185 = _contentList[indexPLCYield + 21];
                string value_6558 = _contentList[indexPLCYield + 31];
                string value_2626 = _contentList[indexPLCYield + 39];
                string value_59 = _contentList[indexPLCYield + 26];
                string value_37 = _contentList[indexPLCYield + 35];
                string value_40 = _contentList[indexPLCYield + 43];

                //dumping data in excel file
                Excel.Application xlApp = new Excel.Application();

                if (xlApp == null)
                {
                    Response.Write("Excel is not properly installed!!");
                    return;
                }

                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet.Cells[1, 1] = "1.Program_Year";
                xlWorkSheet.Cells[1, 2] = "2.State_Code";
                xlWorkSheet.Cells[1, 3] = "3.Country_Code";
                xlWorkSheet.Cells[1, 4] = "4.Fram_Number";
                xlWorkSheet.Cells[1, 5] = "5A.County FSA Office Name and Addres";
                xlWorkSheet.Cells[1, 6] = "5B.County Office Telephone No";
                xlWorkSheet.Cells[1, 7] = "5C.County Office Fax No";
                xlWorkSheet.Cells[1, 8] = "6.Multi-year Contract (2019 - 2023)";

                xlWorkSheet.Cells[1, 9] = "7. Comodity";
                xlWorkSheet.Cells[1, 10] = "8. Program Elected";
                xlWorkSheet.Cells[1, 11] = "Base Acres";
                xlWorkSheet.Cells[1, 12] = "PLC Yield";

                xlWorkSheet.Cells[1, 13] = "12A.. Owner or Producer's Name and Address";
                xlWorkSheet.Cells[1, 14] = "12B. Email Address";
                xlWorkSheet.Cells[1, 15] = "12C. Telephone No";
                xlWorkSheet.Cells[2, 1] = ProgValue;
                xlWorkSheet.Cells[2, 2] = StateValue;
                xlWorkSheet.Cells[2, 3] = CountryValue;
                xlWorkSheet.Cells[2, 4] = FarmValue;
                xlWorkSheet.Cells[2, 5] = FSAOfficeValue;
                xlWorkSheet.Cells[2, 6] = CountryOfficeValue;
                xlWorkSheet.Cells[2, 7] = CountryFaxValue;
                xlWorkSheet.Cells[2, 8] = MultiYearContractValue;
                //Comodity value in excel
                xlWorkSheet.Cells[2, 9] = cornValue;
                xlWorkSheet.Cells[3, 9] = ricelongGrainValue;
                xlWorkSheet.Cells[4, 9] = seedcottonValue;
                xlWorkSheet.Cells[5, 9] = grainsorghumValue;
                xlWorkSheet.Cells[6, 9] = soyabeansValue;
                xlWorkSheet.Cells[7, 9] = wheatValue;
                //Program Elected
                xlWorkSheet.Cells[2, 10] = plcValue;
                xlWorkSheet.Cells[3, 10] = plcValue;
                xlWorkSheet.Cells[4, 10] = plcValue;
                xlWorkSheet.Cells[5, 10] = plcValue;
                xlWorkSheet.Cells[6, 10] = arcCountyValue;
                xlWorkSheet.Cells[7, 10] = plcValue;
                //Base Acres
                xlWorkSheet.Cells[2, 11] = value_643;
                xlWorkSheet.Cells[3, 11] = value_336;
                xlWorkSheet.Cells[4, 11] = value_1052;
                xlWorkSheet.Cells[5, 11] = value_27;
                xlWorkSheet.Cells[6, 11] = value_1853;
                xlWorkSheet.Cells[7, 11] = value_387;
                //PLC Yield
                xlWorkSheet.Cells[2, 12] = value_185;
                xlWorkSheet.Cells[3, 12] = value_6558;
                xlWorkSheet.Cells[4, 12] = value_2626;
                xlWorkSheet.Cells[5, 12] = value_59;
                xlWorkSheet.Cells[6, 12] = value_37;
                xlWorkSheet.Cells[7, 12] = value_40;

                xlWorkSheet.Cells[2, 13] = ownerProducerValue;
                xlWorkSheet.Cells[2, 14] = emailvalue;
                xlWorkSheet.Cells[2, 15] = telePhoneNum;

                xlWorkBook.SaveAs(@"C:\PDFExcel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
                Response.Write("Excel file created in c drive");
            }
            else
            {
                Response.Write("Please select file to upload");
            }

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

    }
}