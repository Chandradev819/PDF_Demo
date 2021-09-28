using org.pdfclown.documents;
using org.pdfclown.documents.contents;
using org.pdfclown.documents.contents.objects;
using org.pdfclown.files;
using System;
using System.Collections.Generic;

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
            _contentList = new List<string>();
            CreatePdfContent(@"C:\Users\chandradev_ps\Desktop\Chandradev\Demo.pdf");

            var indexProg = _contentList.FindIndex(m => m == "1. Program Year: ") + 1;
            int.TryParse(_contentList[indexProg], out var ProgValue);

            var indexState = _contentList.FindIndex(m => m == "2. State Code") + 3;
            int.TryParse(_contentList[indexState], out var StateValue);

            var indexCountry = _contentList.FindIndex(m => m == "3. County Code") + 3;
            int.TryParse(_contentList[indexCountry], out var CountryValue);

            var indexFarm = _contentList.FindIndex(m => m == "4. Farm Number") + 3;
            int.TryParse(_contentList[indexFarm], out var FarmValue);

            var indexFSAOffice = _contentList.FindIndex(m => m == "5A. County FSA Office Name and Address") + 1;
            int.TryParse(_contentList[indexFSAOffice], out var FSAOfficeValue);

            var indexCountryOffice = _contentList.FindIndex(m => m == "5B. County Office Telephone No") + 4;
            int.TryParse(_contentList[indexCountryOffice], out var CountryOfficeValue);

            var indexCountryFax = _contentList.FindIndex(m => m == "5C. County Office Fax No") + 3;
            int.TryParse(_contentList[indexCountryFax], out var CountryFaxValue);

            var indexMultiYearContract = _contentList.FindIndex(m => m == "6.  Multi-year Contract ") + 1;
            int.TryParse(_contentList[indexMultiYearContract], out var MultiYearContractValue);

            var indexOwnerProducer = _contentList.FindIndex(m => m == "12A. Owner or Producer's Name and Address") + 1;
            int.TryParse(_contentList[indexOwnerProducer], out var OwnerProducerValue);

            var indexEmailId = _contentList.FindIndex(m => m == "12B. Email Address") + 1;
            int.TryParse(_contentList[indexEmailId], out var EmailIdValue);

            var indexTelephoneNum = _contentList.FindIndex(m => m == "12C. Telephone No. ") + 1;
            int.TryParse(_contentList[indexTelephoneNum], out var TelephoneNumValue);

            Response.Write(ProgValue);
            Response.Write(StateValue);
            Response.Write(CountryValue);
            Response.Write(FarmValue);
            Response.Write(FSAOfficeValue);
            Response.Write(CountryOfficeValue);
            Response.Write(CountryFaxValue);
            Response.Write(MultiYearContractValue);
            Response.Write(OwnerProducerValue);
            Response.Write(EmailIdValue);
            Response.Write(TelephoneNumValue);
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
    }
}