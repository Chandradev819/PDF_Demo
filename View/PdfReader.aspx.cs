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
                string extension = System.IO.Path.GetExtension(FileUpload1.PostedFile.FileName);
                if (extension == ".pdf")
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
                   
                    var paymentshare = _contentList.FindIndex(m => m == "Payment Share");
                    string valuepaymentshare_8 = _contentList[paymentshare+ 65];
                    string valuepaymentshare_100 = _contentList[paymentshare + 67];
                    string valuepaymentshare_empty = string.Empty;
                    string valuepaymentshare_15= _contentList[paymentshare + 70];
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
                    xlWorkSheet.Cells[1, 5] = "5A.County_FSA_Office_Name_and_Address";
                    xlWorkSheet.Cells[1, 6] = "5B.County_Office_Telephone_No";
                    xlWorkSheet.Cells[1, 7] = "5C.County_Office_Fax_No";
                    xlWorkSheet.Cells[1, 8] = "6.Multi-year_Contract_(2019-2023)";

                    xlWorkSheet.Cells[1, 9] = "7.Comodity";
                    xlWorkSheet.Cells[1, 10] = "7.2Comodity";
                    xlWorkSheet.Cells[1, 11] = "7.3Comodity";
                    xlWorkSheet.Cells[1, 12] = "7.4Comodity";
                    xlWorkSheet.Cells[1, 13] = "7.5Comodity";
                    xlWorkSheet.Cells[1, 14] = "7.6Comodity";
                    xlWorkSheet.Cells[1, 15] = "8.Program_Elected";
                    xlWorkSheet.Cells[1, 16] = "8.2Program_Elected";
                    xlWorkSheet.Cells[1, 17] = "8.3Program_Elected";
                    xlWorkSheet.Cells[1, 18] = "8.4Program_Elected";
                    xlWorkSheet.Cells[1, 19] = "8.5Program_Elected";
                    xlWorkSheet.Cells[1, 20] = "8.6Program_Elected";

                    xlWorkSheet.Cells[1, 21] = "9.Base_Acres";
                    xlWorkSheet.Cells[1, 22] = "9.2Base_Acres";
                    xlWorkSheet.Cells[1, 23] = "9.3Base_Acres";
                    xlWorkSheet.Cells[1, 24] = "9.4Base_Acres";
                    xlWorkSheet.Cells[1, 25] = "9.5Base_Acres";
                    xlWorkSheet.Cells[1, 26] = "9.6Base_Acres";

                    xlWorkSheet.Cells[1, 27] = "10.PLC_Yield";
                    xlWorkSheet.Cells[1, 28] = "10.2PLC_Yield";
                    xlWorkSheet.Cells[1, 29] = "10.3PLC_Yield";
                    xlWorkSheet.Cells[1, 30] = "10.4PLC_Yield";
                    xlWorkSheet.Cells[1, 31] = "10.5PLC_Yield";
                    xlWorkSheet.Cells[1, 32] = "10.6PLC_Yield";

                    xlWorkSheet.Cells[1, 33] = "11.Participating";
                    xlWorkSheet.Cells[1, 34] = "11.2Participating";
                    xlWorkSheet.Cells[1, 35] = "11.3Participating";
                    xlWorkSheet.Cells[1, 36] = "11.4Participating";
                    xlWorkSheet.Cells[1, 37] = "11.5Participating";
                    xlWorkSheet.Cells[1, 38] = "11.6Participating";

                    xlWorkSheet.Cells[1, 39] = "12A.Owner_or_Producer's_Name_and_Address";
                    xlWorkSheet.Cells[1, 40] = "12B.Email_Address";
                    xlWorkSheet.Cells[1, 41] = "12C.Telephone_No";

                    xlWorkSheet.Cells[1, 42] = "13.Commodity";
                    xlWorkSheet.Cells[1, 43] = "13.2Commodity";
                    xlWorkSheet.Cells[1, 44] = "13.3Commodity";
                    xlWorkSheet.Cells[1, 45] = "13.4Commodity";
                    xlWorkSheet.Cells[1, 46] = "13.5Commodity";
                    xlWorkSheet.Cells[1, 47] = "13.6Commodity";
                    xlWorkSheet.Cells[1, 48] = "13.7Commodity";
                    xlWorkSheet.Cells[1, 49] = "13.8Commodity";

                    xlWorkSheet.Cells[1, 50] = "14.Payment_Share";
                    xlWorkSheet.Cells[1, 51] = "14.2Payment_Share";
                    xlWorkSheet.Cells[1, 52] = "14.3Payment_Share";
                    xlWorkSheet.Cells[1, 53] = "14.4Payment_Share";
                    xlWorkSheet.Cells[1, 54] = "14.5Payment_Share";
                    xlWorkSheet.Cells[1, 55] = "14.6Payment_Share";
                    xlWorkSheet.Cells[1, 56] = "14.7Payment_Share";
                    xlWorkSheet.Cells[1, 57] = "14.8Payment_Share";

                    xlWorkSheet.Cells[1, 58] = "P2.1.Program_Year";
                    xlWorkSheet.Cells[1, 59] = "P2.2.State_Code";
                    xlWorkSheet.Cells[1, 60] = "P2.3._County_Code";
                    xlWorkSheet.Cells[1, 61] = "P2.4.Farm_Number";

                    xlWorkSheet.Cells[1, 62] = "12A.Owner_or_Producer's_Name_and_Address";
                    xlWorkSheet.Cells[1, 63] = "P2.12B.Email_Address";
                    xlWorkSheet.Cells[1, 64] = "P2.12C._Telephone_No";

                    xlWorkSheet.Cells[1, 65] = "P2.13.Commodity";
                    xlWorkSheet.Cells[1, 66] = "P2.13.2Commodity";
                    xlWorkSheet.Cells[1, 67] = "P2.13.3Commodity";
                    xlWorkSheet.Cells[1, 68] = "P2.13.4Commodity";
                    xlWorkSheet.Cells[1, 69] = "P2.13.5Commodity";
                    xlWorkSheet.Cells[1, 70] = "P2.13.6Commodity";
                    xlWorkSheet.Cells[1, 71] = "P2.13.7Commodity";
                    xlWorkSheet.Cells[1, 72] = "P2.13.8Commodity";

                    xlWorkSheet.Cells[1, 73] = "P2.14.Payment_Share";
                    xlWorkSheet.Cells[1, 74] = "P2.14.2Payment_Share";
                    xlWorkSheet.Cells[1, 75] = "P2.14.3Payment_Share";
                    xlWorkSheet.Cells[1, 76] = "P2.14.4Payment_Share";
                    xlWorkSheet.Cells[1, 77] = "P2.14.5Payment_Share";
                    xlWorkSheet.Cells[1, 78] = "P2.14.6Payment_Share";

                    xlWorkSheet.Cells[1, 79] = "P2.14.7Payment_Share";
                    xlWorkSheet.Cells[1, 80] = "P2.14.8Payment_Share";
                    xlWorkSheet.Cells[1, 81] = "P2.15A.Refused_Payment_Information";
                    xlWorkSheet.Cells[1, 82] = "P2.15B.Producer's_Initials";
                    xlWorkSheet.Cells[1, 83] = "P2.15C.Date_Initialed_MM-DD-YYYY";
                    xlWorkSheet.Cells[1, 84] = "P2.16A.Producer's_Signature_By";
                    xlWorkSheet.Cells[1, 85] = "P2.16B.Title/Relationship_of_the_Individual_Signing_in_the_Representative_Capacity";
                    xlWorkSheet.Cells[1, 86] = "P2.16C.Date_MM-DD-YYYY";
                    xlWorkSheet.Cells[1, 87] = "12A.Owner_or_Producer's_Name_and_Address";
                    xlWorkSheet.Cells[1, 88] = "12A.Owner_or_Producer's_Name_and_Address";
                    xlWorkSheet.Cells[1, 89] = "12A.Owner_or_Producer's_Name_and_Address";
                    xlWorkSheet.Cells[1, 90] = "12A.Owner_or_Producer's_Name_and_Address";
                    xlWorkSheet.Cells[1, 91] = "12A.Owner_or_Producer's_Name_and_Address";

                    //Filling on Cell

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