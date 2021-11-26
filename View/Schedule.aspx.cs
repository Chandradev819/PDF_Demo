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
    public partial class Schedule : System.Web.UI.Page
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
                string extension = System.IO.Path.GetExtension(FileUpload1.PostedFile.FileName);
                if (extension == ".pdf")
                {
                    string fileName = System.IO.Path.Combine(Server.MapPath("~/Pdf"), FileUpload1.FileName);
                    FileUpload1.SaveAs(fileName);
                    _contentList = new List<string>();
                    CreatePdfContent(fileName);
                    //For Propritetor

                    var indexPropri = _contentList.FindIndex(m => m == "Name of proprietor") + 1;
                    var indexProg = _contentList.FindIndex(m => m == "Test");
                    //int.TryParse(_contentList[indexProg], out var ProgValue);

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

                    var paymentshare = _contentList.FindIndex(m => m == "Payment Share");
                    string valuepaymentshare_8 = _contentList[paymentshare + 65];
                    string valuepaymentshare_100 = _contentList[paymentshare + 67];
                    string valuepaymentshare_empty = string.Empty;
                    string valuepaymentshare_15 = _contentList[paymentshare + 70];
                    string valuepaymentshare_90 = _contentList[paymentshare + 74];

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
                    xlWorkSheet.Cells[1, 10] = "7.2 Comodity";
                    xlWorkSheet.Cells[1, 11] = "7.3 Comodity";
                    xlWorkSheet.Cells[1, 12] = "7.4 Comodity";
                    xlWorkSheet.Cells[1, 13] = "7.5 Comodity";
                    xlWorkSheet.Cells[1, 14] = "7.6 Comodity";
                    xlWorkSheet.Cells[1, 15] = "8. Program Elected";
                    xlWorkSheet.Cells[1, 16] = "8.2 Program Elected";
                    xlWorkSheet.Cells[1, 17] = "8.3 Program Elected";
                    xlWorkSheet.Cells[1, 18] = "8.4 Program Elected";
                    xlWorkSheet.Cells[1, 19] = "8.5 Program Elected";
                    xlWorkSheet.Cells[1, 20] = "8.6 Program Elected";

                    xlWorkSheet.Cells[1, 21] = "9. Base Acres";
                    xlWorkSheet.Cells[1, 22] = "9.2 Base Acres";
                    xlWorkSheet.Cells[1, 23] = "9.3 Base Acres";
                    xlWorkSheet.Cells[1, 24] = "9.4 Base Acres";
                    xlWorkSheet.Cells[1, 25] = "9.5 Base Acres";
                    xlWorkSheet.Cells[1, 26] = "9.6 Base Acres";

                    xlWorkSheet.Cells[1, 27] = "10. PLC Yield";
                    xlWorkSheet.Cells[1, 28] = "10.2 PLC Yield";
                    xlWorkSheet.Cells[1, 29] = "10.3 PLC Yield";
                    xlWorkSheet.Cells[1, 30] = "10.4 PLC Yield";
                    xlWorkSheet.Cells[1, 31] = "10.5 PLC Yield";
                    xlWorkSheet.Cells[1, 32] = "10.6 PLC Yield";

                    xlWorkSheet.Cells[1, 33] = "12A.. Owner or Producer's Name and Address";
                    xlWorkSheet.Cells[1, 34] = "12B. Email Address";
                    xlWorkSheet.Cells[1, 35] = "12C. Telephone No";

                    xlWorkSheet.Cells[1, 36] = "P2.14 PAYMENT SHARE";
                    xlWorkSheet.Cells[1, 37] = "P2.14.2 PAYMENT SHARE";
                    xlWorkSheet.Cells[1, 38] = "P2.14.3 PAYMENT SHARE";
                    xlWorkSheet.Cells[1, 39] = "P2.14.4 PAYMENT SHARE";
                    xlWorkSheet.Cells[1, 40] = "P2.14.5 PAYMENT SHARE";
                    xlWorkSheet.Cells[1, 41] = "P2.14.6 PAYMENT SHARE";
                    //Filling on Cell

                   // xlWorkSheet.Cells[2, 1] = ProgValue;
                    xlWorkSheet.Cells[2, 2] = StateValue;
                    xlWorkSheet.Cells[2, 3] = CountryValue;
                    xlWorkSheet.Cells[2, 4] = FarmValue;
                    xlWorkSheet.Cells[2, 5] = FSAOfficeValue;
                    xlWorkSheet.Cells[2, 6] = CountryOfficeValue;
                    xlWorkSheet.Cells[2, 7] = CountryFaxValue;
                    xlWorkSheet.Cells[2, 8] = MultiYearContractValue;

                    //Comodity value in excel
                    xlWorkSheet.Cells[2, 9] = cornValue;
                    xlWorkSheet.Cells[2, 10] = ricelongGrainValue;
                    xlWorkSheet.Cells[2, 11] = seedcottonValue;
                    xlWorkSheet.Cells[2, 12] = grainsorghumValue;
                    xlWorkSheet.Cells[2, 13] = soyabeansValue;
                    xlWorkSheet.Cells[2, 14] = wheatValue;

                    //Program Elected
                    xlWorkSheet.Cells[2, 15] = plcValue;
                    xlWorkSheet.Cells[2, 16] = plcValue;
                    xlWorkSheet.Cells[2, 17] = plcValue;
                    xlWorkSheet.Cells[2, 18] = plcValue;
                    xlWorkSheet.Cells[2, 19] = arcCountyValue;
                    xlWorkSheet.Cells[2, 20] = plcValue;
                    //Base Acres
                    xlWorkSheet.Cells[2, 21] = value_643;
                    xlWorkSheet.Cells[2, 22] = value_336;
                    xlWorkSheet.Cells[2, 23] = value_1052;
                    xlWorkSheet.Cells[2, 24] = value_27;
                    xlWorkSheet.Cells[2, 25] = value_1853;
                    xlWorkSheet.Cells[2, 26] = value_387;

                    //PLC Yield
                    xlWorkSheet.Cells[2, 27] = value_185;
                    xlWorkSheet.Cells[2, 28] = value_6558;
                    xlWorkSheet.Cells[2, 29] = value_2626;
                    xlWorkSheet.Cells[2, 30] = value_59;
                    xlWorkSheet.Cells[2, 31] = value_37;
                    xlWorkSheet.Cells[2, 32] = value_40;

                    xlWorkSheet.Cells[2, 33] = ownerProducerValue;
                    xlWorkSheet.Cells[2, 34] = emailvalue;
                    xlWorkSheet.Cells[2, 35] = telePhoneNum;

                    xlWorkSheet.Cells[2, 36] = valuepaymentshare_8;
                    xlWorkSheet.Cells[2, 37] = valuepaymentshare_empty;
                    xlWorkSheet.Cells[2, 38] = valuepaymentshare_100;
                    xlWorkSheet.Cells[2, 39] = valuepaymentshare_100;
                    xlWorkSheet.Cells[2, 40] = valuepaymentshare_15;
                    xlWorkSheet.Cells[2, 41] = valuepaymentshare_90;

                    xlWorkBook.SaveAs(@"C:\PDFExcel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                    Response.Write("Excel file created in c drive");
                }
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