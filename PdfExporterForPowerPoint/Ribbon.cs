using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using CustomControlLibrary;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PdfExporterForPowerPoint
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
            return GetResourceText("PdfExporterForPowerPoint.Ribbon.xml");
        }

        #endregion

        #region リボンのコールバック

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void ExportPdf(Office.IRibbonControl control)
        {
            var document = Globals.ThisAddIn.Application.ActivePresentation;
            
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "pdf|*.pdf",
                CreatePrompt = false,
                FileName = document.Name.Replace(".pptx", ".pdf")
            };

            var tagetPath = saveFileDialog.ShowDialog() == DialogResult.OK ? saveFileDialog.FileName : null;
            saveFileDialog.Dispose();

            if (tagetPath != null)
            {
                document.ExportAsFixedFormat(tagetPath, PpFixedFormatType.ppFixedFormatTypePDF,
                    PpFixedFormatIntent.ppFixedFormatIntentPrint, Office.MsoTriState.msoFalse,
                    PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                    PpPrintOutputType.ppPrintOutputSlides, Office.MsoTriState.msoFalse, null,
                    PpPrintRangeType.ppPrintAll, "", true, true, true, true, false, Missing.Value);

                Process.Start(tagetPath);
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
                    var document = Globals.ThisAddIn.Application.ActivePresentation;
                    var tempFolder = Path.GetTempPath();
                    var tempPath = tempFolder + "pdf.temp";

                    var saveFileDialog = new SaveFileDialog
                    {
                        Filter = "pdf|*.pdf",
                        CreatePrompt = false,
                        FileName = document.Name.Replace(".pptx", ".pdf")
                    };

                    var tagetPath = saveFileDialog.ShowDialog() == DialogResult.OK ? saveFileDialog.FileName : null;
                    saveFileDialog.Dispose();

                    if (tagetPath != null)
                    {
                        document.ExportAsFixedFormat(tempPath, PpFixedFormatType.ppFixedFormatTypePDF,
                            PpFixedFormatIntent.ppFixedFormatIntentPrint, Office.MsoTriState.msoFalse,
                            PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                            PpPrintOutputType.ppPrintOutputSlides, Office.MsoTriState.msoFalse, null,
                            PpPrintRangeType.ppPrintAll, "", true, true, true, true, false, Missing.Value);

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
