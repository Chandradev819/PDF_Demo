using System;
using System.Threading;
using System.Web;
using System.Windows.Forms;

namespace PDF_Demo.Services
{
    /// <summary>
    /// Summary description for UtilityHandler
    /// </summary>
    public class UtilityHandler : IHttpHandler
    {
        [MTAThread]
        public void ProcessRequest(HttpContext context)
        {
            string selectedPath = "";

            Thread t = new Thread((ThreadStart)(() =>
            {
                FolderBrowserDialog folderDialog = new FolderBrowserDialog();
                folderDialog.RootFolder = System.Environment.SpecialFolder.MyComputer;
                folderDialog.ShowNewFolderButton = true;
                if (folderDialog.ShowDialog() == DialogResult.Cancel)
                    context.Response.Write("");

                selectedPath = folderDialog.SelectedPath;
            }));

            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
            Console.WriteLine(selectedPath);
            HttpCookie cookie = new HttpCookie("path");
            cookie["path"] = selectedPath;

            context.Response.ContentType = "text/plain";
            context.Response.Write(selectedPath);
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