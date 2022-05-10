using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;

namespace ClipboardConverter
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
			/*Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);*/

			ListenForClipboardChanges();
        }

		private static void ListenForClipboardChanges()
        {
			String lastClipboard = "";
			while(true)
            {
				try
                {
					string currentClipboard = Clipboard.GetText(TextDataFormat.Html);

					if (currentClipboard != null && currentClipboard != lastClipboard && currentClipboard.Contains("<tr class=\"gridHeader\">"))
                    {
						//Clipboard.SetText("", TextDataFormat.Rtf);
						Thread.Sleep(50);
						ConvertClipboardToRTF(currentClipboard.Substring(currentClipboard.IndexOf("<table>")));
					}
					lastClipboard = currentClipboard;
					Thread.Sleep(50);
				}
				catch(Exception e)
                {

                }
            }
        }

		public static void ConvertClipboardToRTF(string html)
		{
			RichTextBox rtbTemp = new RichTextBox();
			WebBrowser wb = new WebBrowser();
			wb.Navigate("about:blank");

			wb.Document.Write(html);
			wb.Document.ExecCommand("SelectAll", false, null);
			wb.Document.ExecCommand("Copy", false, null);

			rtbTemp.SelectAll();
			rtbTemp.Paste();
			rtbTemp.SelectAll();
			rtbTemp.Copy();

			string borderInsert = "\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrb\\brdrs\\clbrdrr\\brdrs";

			string[] split = rtbTemp.Rtf.Split(@"\cellx");

			string endRtf = "";

			//Insert borders
			foreach (string entry in split)
			{
				endRtf += entry + borderInsert + "\\cellx";
			}

			Clipboard.SetText(endRtf, TextDataFormat.Rtf);
		}
	}
}
