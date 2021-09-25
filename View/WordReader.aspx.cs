using System;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace PDF_Demo.View
{
    public partial class WordReader : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            using (PdfDocument document = PdfDocument.Open(@"C:\Users\chandradev_ps\Desktop\Chandradev\Demo.pdf"))
            {
                foreach (Page page in document.GetPages())
                {
                    string pageText = page.Text;

                    foreach (Word word in page.GetWords())
                    {
                        if (word.Text== "State")
                        {
                            Response.Write(word.Text);
                        }
                        
                    }
                }
            }
        }
    }
}