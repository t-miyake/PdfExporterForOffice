using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using CustomControlLibrary;

namespace PdfExporterForWord
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        #region IRibbonExtensibility のメンバー

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PdfExporterForWord.Ribbon.xml");
        }

        #endregion

        #region リボンのコールバック

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }

        public void ExportPdf(Office.IRibbonControl control)
        {
            var document = Globals.ThisAddIn.Application.ActiveDocument;

                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "pdf|*.pdf",
                    CreatePrompt = false,
                    FileName = document.Name.Replace(".docx", ".pdf")
                };

                var tagetPath = saveFileDialog.ShowDialog() == DialogResult.OK ? saveFileDialog.FileName : null;
                saveFileDialog.Dispose();

                if (tagetPath != null)
                {
                    document.ExportAsFixedFormat(tagetPath, WdExportFormat.wdExportFormatPDF, true,
                        WdExportOptimizeFor.wdExportOptimizeForPrint, WdExportRange.wdExportAllDocument, 0, 0,
                        WdExportItem.wdExportDocumentContent, true, true,
                        WdExportCreateBookmarks.wdExportCreateNoBookmarks, true, true, false, Missing.Value);
                }
        }

        public void ExportPdfWithPass(Office.IRibbonControl control)
        {
            var passwordInput = new PasswordInputWindow();
            var dialogResult = passwordInput.ShowDialog();

            if (dialogResult == true)
            {
                var password = passwordInput.Password;
                passwordInput.Close();

                if (!string.IsNullOrEmpty(password))
                {
                    var document = Globals.ThisAddIn.Application.ActiveDocument;
                    var tempFolder = Path.GetTempPath();
                    var tempPath = tempFolder + "pdf.temp";

                    var saveFileDialog = new SaveFileDialog
                    {
                        Filter = "pdf|*.pdf",
                        CreatePrompt = false,
                        FileName = document.FullName.Replace(".docx", ".pdf")
                    };

                    var tagetPath = saveFileDialog.ShowDialog() == DialogResult.OK ? saveFileDialog.FileName : null;
                    saveFileDialog.Dispose();

                    if (tagetPath != null)
                    {
                        document.ExportAsFixedFormat(tempPath, WdExportFormat.wdExportFormatPDF,false,
                            WdExportOptimizeFor.wdExportOptimizeForPrint, WdExportRange.wdExportAllDocument, 0, 0,
                            WdExportItem.wdExportDocumentContent, true, true,
                            WdExportCreateBookmarks.wdExportCreateNoBookmarks, true, true, false, Missing.Value);

                        var reader = new PdfReader(tempPath);
                        using (var output = new FileStream(tagetPath, FileMode.Create, FileAccess.Write,
                            FileShare.None))
                        {
                            PdfEncryptor.Encrypt(reader, output, PdfWriter.STRENGTH128BITS, password, password,
                                PdfWriter.ALLOW_COPY | PdfWriter.ALLOW_PRINTING);
                        }
                        reader.Close();

                        File.Delete(tempPath);

                        Process.Start(tagetPath);
                    }
                }
            }
        }

        #endregion

        #region ヘルパー

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
