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
    public partial class CCCPdf : System.Web.UI.Page
    {
        private List<string> _contentList;
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
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

                    //Need to correct the value
                    string plcValue1 = _contentList[indexProgElected + 23];
                    string plcValue2 = _contentList[indexProgElected + 33];
                    string plcValue3 = _contentList[indexProgElected + 41];
                    string plcValue4 = _contentList[indexProgElected + 28];
                    string plcValue5 = _contentList[indexProgElected + 45];

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


                    string participating = string.Empty;

                    //12A. Owner or Producer's Name and Address
                    var indexOwner = _contentList.FindIndex(m => m == "12A. Owner or Producer's Name and Address");
                    string owner_producerName_Address1 = _contentList[indexOwner + 1];
                    string owner_producerName_Address2 = _contentList[indexOwner + 2];
                    string owner_producerName_Address3 = _contentList[indexOwner + 3];

                    string owner_producerName_Address = string.Concat(owner_producerName_Address1 + owner_producerName_Address2 + owner_producerName_Address3);

                    string email_Address = string.Empty;

                    string telephone_no = string.Empty;

                    //var indexCommodity = _contentList.FindIndex(m => m == "Commodity");
                    string commodity_13 = _contentList[188];
                    string commodity_13_2 = _contentList[192];
                    string commodity_13_3 = _contentList[195];
                    string commodity_13_4 = string.Empty;
                    string commodity_13_5 = _contentList[190];
                    string commodity_13_6 = _contentList[193];
                    string commodity_13_7 = _contentList[197];
                    string commodity_13_8 = string.Empty;

                    //For Payment Share
                    var indexPaymentShare = _contentList.FindIndex(m => m == "Payment Share");
                    string paymentshare_14 = _contentList[indexPaymentShare + 6];
                    string paymentshare_14_2 = _contentList[indexPaymentShare + 72];
                    string paymentshare_14_3 = string.Empty;
                    string paymentshare_14_4 = string.Empty;
                    string paymentshare_14_5 = string.Empty;
                    string paymentshare_14_6 = _contentList[indexPaymentShare + 11];
                    string paymentshare_14_7 = _contentList[indexPaymentShare + 14];
                    string paymentshare_14_8 = string.Empty;

                    //12A. Owner or Producer's Name and Address
                    var indexOwnerProducer = _contentList.FindIndex(m => m == "12A. Owner or Producer's Name and Address");
                    string address1 = _contentList[indexOwnerProducer + 59];
                    string address2 = _contentList[indexOwnerProducer + 60];
                    string address3 = _contentList[indexOwnerProducer + 61];
                    string owner_producer_name_address = string.Concat(address1, address2, address3);

                    string emailAddress = string.Empty;
                    //telephone
                    var indexTelephone = _contentList.FindIndex(m => m == "12C. Telephone No. ");
                    string telephoneNum = _contentList[indexTelephone + 60];

                    var indexcommudity = _contentList.FindIndex(m => m == "Commodity");
                    string commudity_13 = _contentList[indexcommudity + 125];
                    string commudity_13_2 = _contentList[indexcommudity + 129];
                    string commudity_13_3 = _contentList[indexcommudity + 132];
                    string commudity_13_4 = string.Empty;
                    string commudity_13_5 = _contentList[indexcommudity + 127];
                    string commudity_13_6 = _contentList[indexcommudity + 130];
                    string commudity_13_7 = _contentList[indexcommudity + 134];
                    string commudity_13_8 = string.Empty;

                    var indexpaymenetshare = _contentList.FindIndex(m => m == "Payment Share");
                    string paymentshare_P2_14 = _contentList[indexpaymenetshare + 65];
                    string paymentshare_P2_14_2 = string.Empty;
                    string paymentshare_P2_14_3 = _contentList[indexpaymenetshare + 72];
                    string paymentshare_P2_14_4 = string.Empty;
                    string paymentshare_P2_14_5 = _contentList[indexpaymenetshare + 67];
                    string paymentshare_P2_14_6 = _contentList[indexpaymenetshare + 70];
                    string paymentshare_P2_14_7 = _contentList[indexpaymenetshare + 74];
                    string paymentshare_P2_14_8 = string.Empty;
                    //15A. Refused Payment Information:
                    var indexRefusedPayment = _contentList.FindIndex(m => m == "15A. Refused Payment Information:");
                    string refused_Payment_Information = string.Concat(_contentList[indexRefusedPayment + 62], " " + _contentList[indexRefusedPayment + 63]);
                    string Producer_Initials = string.Empty;
                    string date_Initialed_MM_DD_yyyy = string.Empty;
                    string Producer_Signature_By = string.Empty;
                    string Relationship_of_the_Individual_Signing_in_the_Representative_Capacity = string.Empty;
                    string Date_MM_DD_YYYY = string.Empty;

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
                    xlWorkSheet.Cells[1, 4] = "4.Farm_Number";
                    xlWorkSheet.Cells[1, 5] = "5A.County_FSA_Office_Name_and_Address";
                    xlWorkSheet.Cells[1, 6] = "5B.County_Office_Telephone_No";
                    xlWorkSheet.Cells[1, 7] = "5C.County_Office_Fax_No";
                    xlWorkSheet.Cells[1, 8] = "6.Multi-year_Contract_(2019-2023)";

                    xlWorkSheet.Cells[1, 9] = "7.Commodity";
                    xlWorkSheet.Cells[1, 10] = "7.2Commodity";
                    xlWorkSheet.Cells[1, 11] = "7.3Commodity";
                    xlWorkSheet.Cells[1, 12] = "7.4Commodity";
                    xlWorkSheet.Cells[1, 13] = "7.5Commodity";
                    xlWorkSheet.Cells[1, 14] = "7.6Commodity";
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

                    xlWorkSheet.Cells[1, 62] = "P2.12A.Owner_or_Producer's_Name_and_Address";
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

                    //xlWorkSheet.Cells[1, 81] = "P2.15A.Refused_Payment_Information";
                    //xlWorkSheet.Cells[1, 82] = "P2.15B.Producer's_Initials";
                    //xlWorkSheet.Cells[1, 83] = "P2.15C.Date_Initialed_MM-DD-YYYY";
                    //xlWorkSheet.Cells[1, 84] = "P2.16A.Producer's_Signature_By";
                    //xlWorkSheet.Cells[1, 85] = "P2.16B.Title/Relationship_of_the_Individual_Signing_in_the_Representative_Capacity";
                    //xlWorkSheet.Cells[1, 86] = "P2.16C.Date_MM-DD-YYYY";

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
                    xlWorkSheet.Cells[2, 15] = plcValue1;
                    xlWorkSheet.Cells[2, 16] = plcValue2;
                    xlWorkSheet.Cells[2, 17] = plcValue3;
                    xlWorkSheet.Cells[2, 18] = plcValue4;
                    xlWorkSheet.Cells[2, 19] = arcCountyValue;
                    xlWorkSheet.Cells[2, 20] = plcValue5;

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

                    xlWorkSheet.Cells[2, 33] = participating;
                    xlWorkSheet.Cells[2, 34] = participating;
                    xlWorkSheet.Cells[2, 35] = participating;
                    xlWorkSheet.Cells[2, 36] = participating;
                    xlWorkSheet.Cells[2, 37] = participating;
                    xlWorkSheet.Cells[2, 38] = participating;

                    //To be fill
                    xlWorkSheet.Cells[2, 39] = owner_producerName_Address;
                    xlWorkSheet.Cells[2, 40] = email_Address;
                    xlWorkSheet.Cells[2, 41] = telephone_no;

                    xlWorkSheet.Cells[2, 42] = commodity_13;
                    xlWorkSheet.Cells[2, 43] = commodity_13_2;
                    xlWorkSheet.Cells[2, 44] = commodity_13_3;
                    xlWorkSheet.Cells[2, 45] = commodity_13_4;
                    xlWorkSheet.Cells[2, 46] = commodity_13_5;
                    xlWorkSheet.Cells[2, 47] = commodity_13_6;
                    xlWorkSheet.Cells[2, 48] = commodity_13_7;
                    xlWorkSheet.Cells[2, 49] = commodity_13_8;

                    xlWorkSheet.Cells[2, 50] = paymentshare_14;
                    xlWorkSheet.Cells[2, 51] = paymentshare_14_2;
                    xlWorkSheet.Cells[2, 52] = paymentshare_14_3;
                    xlWorkSheet.Cells[2, 53] = paymentshare_14_4;
                    xlWorkSheet.Cells[2, 54] = paymentshare_14_5;
                    xlWorkSheet.Cells[2, 55] = paymentshare_14_6;
                    xlWorkSheet.Cells[2, 56] = paymentshare_14_7;
                    xlWorkSheet.Cells[2, 57] = paymentshare_14_8;

                    xlWorkSheet.Cells[2, 58] = ProgValue;
                    xlWorkSheet.Cells[2, 59] = StateValue;
                    xlWorkSheet.Cells[2, 60] = CountryValue;
                    xlWorkSheet.Cells[2, 61] = FarmValue;

                    xlWorkSheet.Cells[2, 62] = owner_producer_name_address;
                    xlWorkSheet.Cells[2, 63] = emailAddress;
                    xlWorkSheet.Cells[2, 64] = telePhoneNum;

                    xlWorkSheet.Cells[2, 65] = commodity_13;
                    xlWorkSheet.Cells[2, 66] = commodity_13_2;
                    xlWorkSheet.Cells[2, 67] = commodity_13_3;
                    xlWorkSheet.Cells[2, 68] = commodity_13_4;
                    xlWorkSheet.Cells[2, 69] = commodity_13_5;
                    xlWorkSheet.Cells[2, 70] = commodity_13_6;
                    xlWorkSheet.Cells[2, 71] = commodity_13_7;
                    xlWorkSheet.Cells[2, 72] = commodity_13_8;

                    xlWorkSheet.Cells[2, 73] = paymentshare_P2_14;
                    xlWorkSheet.Cells[2, 74] = paymentshare_P2_14_2;
                    xlWorkSheet.Cells[2, 75] = paymentshare_P2_14_3;
                    xlWorkSheet.Cells[2, 76] = paymentshare_P2_14_4;
                    xlWorkSheet.Cells[2, 77] = paymentshare_P2_14_5;
                    xlWorkSheet.Cells[2, 78] = paymentshare_P2_14_6;
                    xlWorkSheet.Cells[2, 79] = paymentshare_P2_14_7;
                    xlWorkSheet.Cells[2, 80] = paymentshare_P2_14_8;

                    //xlWorkSheet.Cells[2, 81] = refused_Payment_Information;
                    //xlWorkSheet.Cells[2, 82] = Producer_Initials;
                    //xlWorkSheet.Cells[2, 83] = date_Initialed_MM_DD_yyyy;
                    //xlWorkSheet.Cells[2, 84] = Producer_Signature_By;
                    //xlWorkSheet.Cells[2, 85] = Relationship_of_the_Individual_Signing_in_the_Representative_Capacity;
                    //xlWorkSheet.Cells[2, 86] = Date_MM_DD_YYYY;

                    xlWorkBook.SaveAs(@"C:\Users\chandradev_ps\Desktop\Input\ExcelOutput.xlsx", 
                        Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, 
                        Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
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