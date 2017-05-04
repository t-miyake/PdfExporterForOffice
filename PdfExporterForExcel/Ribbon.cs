using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using CustomControlLibrary;

namespace PdfExporterForExcel
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility のメンバー

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PdfExporterForExcel.Ribbon.xml");
        }

        #endregion

        #region リボンのコールバック
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void ExportPdf(Office.IRibbonControl control)
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            var activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "pdf|*.pdf",
                    CreatePrompt = false,
                    FileName = activeWorkbook.Name.Replace(".xlsx", ".pdf")
                };

                var tagetPath = saveFileDialog.ShowDialog() == DialogResult.OK ? saveFileDialog.FileName : null;
                saveFileDialog.Dispose();

                if (tagetPath != null)
                {
                    activeSheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, tagetPath, XlFixedFormatQuality.xlQualityStandard,true,false, Missing.Value, Missing.Value, true, Missing.Value);
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
                    Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
                    var activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

                    var tempFolder = Path.GetTempPath();
                    var tempPath = tempFolder + "pdf.temp.pdf";

                    if (!string.IsNullOrEmpty(password))
                    {
                        var saveFileDialog = new SaveFileDialog
                        {
                            Filter = "pdf|*.pdf",
                            CreatePrompt = false,
                            FileName = activeWorkbook.FullName.Replace(".xlsx", ".pdf")
                        };

                        var tagetPath = saveFileDialog.ShowDialog() == DialogResult.OK ? saveFileDialog.FileName : null;
                        saveFileDialog.Dispose();

                        if (tagetPath != null)
                        {
                            activeSheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, tempPath,
                                XlFixedFormatQuality.xlQualityStandard, true, false, Missing.Value, Missing.Value,
                                false,
                                Missing.Value);

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
