using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace XlsToPDFExporter
{
    /// <summary>
    /// Logika interakcji dla klasy MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void FileExportBtn_Click(object sender, RoutedEventArgs e)
        {
            ChangeStateOfControlsToDocsInProgress();
            try
            {

                var dialog = new OpenFileDialog()
                {
                    Multiselect = true,
                    DefaultExt = ".xls",
                    Filter = "Excel files (.xls,.xlsx, .xlsm)|*.xls;*.xlsx;*.xlsm",
                    CheckPathExists = true,
                    CheckFileExists = true
                };
                if (dialog.ShowDialog() == true)
                {
                   await ExportFiles(dialog.FileNames.Where(fileName => xlsFileExtension.IsMatch(fileName)).ToArray<string>());
                }

                FinishApplication();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

         
        }

        private void FinishApplication()
        {
            if (MessageBox.Show("The files has been successfully exported.", "End of task", MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
            {
                this.Close();
            }

        }

        private async Task<bool> ExportFiles(string[] fileNames)
        {
            Microsoft.Office.Interop.Excel.Application app = null;
            try
            {
                string dirToExport = CreatePDFFolderForExportedDocs(fileNames[0]);
                app = new Microsoft.Office.Interop.Excel.Application();

                IProgress<int> progress = new Progress<int>(report =>
                {
                    progressLabel.Content = string.Format("{0}/{1}", report, fileNames.Length);
                });

                await Task.Factory.StartNew(() => ExportEachFile(fileNames, app, dirToExport, progress));

                
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Some problem with files export. See details: {0}", ex.Message));
            }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                }
               
            }
            return true;
        }

        private void ExportEachFile(string[] fileNames, Microsoft.Office.Interop.Excel.Application app, string dirToExport, IProgress<int> progress)
        {
            
            for (int i=1; i<= fileNames.Length; i++)
            {
                ExportToPDF(app, fileNames[i-1], dirToExport);
                progress.Report(i);
            }
        }

        private string CreatePDFFolderForExportedDocs(string fullFilePathToExport)
        {
            string dirOfFile = System.IO.Path.GetDirectoryName(fullFilePathToExport);
            DirectoryInfo pdfExportDir = new DirectoryInfo(string.Format("{0}\\PDF", dirOfFile));

            if (!pdfExportDir.Exists)
            {
                pdfExportDir.Create();
            }

            return pdfExportDir.FullName;
        }

        private Regex xlsFileExtension = new Regex(@"\.xls(x)?");
        private void ExportToPDF(Microsoft.Office.Interop.Excel.Application app, string filePath, string dirToExport)
        {


            string pdfFileName = xlsFileExtension.Replace(System.IO.Path.GetFileName(filePath), @".pdf");


            try
            {
                Workbook theWorkbook = app.Workbooks.Open(
                    filePath, true, false,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                string fullPDFFilePath = System.IO.Path.Combine(dirToExport, pdfFileName);

                theWorkbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, fullPDFFilePath, IgnorePrintAreas: false);

            }
            catch (Exception ex)
            {
                throw new PdfExportException(string.Format("There is a problem with the file: {0}. Msg: {1}", filePath, ex.Message));
            }

        }

        private async void dirExportBtn_Click(object sender, RoutedEventArgs e)
        {
            ChangeStateOfControlsToDocsInProgress();

            try
            {

                var dialog = new CommonOpenFileDialog()
                {
                    IsFolderPicker = true,
                    Title = "Wybierz folder"
                };
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    string directoryPath = dialog.FileName;
                    string[] filesInDirectory = Directory.GetFiles(directoryPath);

                    List<string> filesToExport = new List<string>();

                    foreach (string file in filesInDirectory)
                    {
                        if (File.Exists(file) && xlsFileExtension.IsMatch(file))
                        {
                            filesToExport.Add(file);
                        }
                    }

                    await ExportFiles(filesToExport.ToArray());
                }

                FinishApplication();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ChangeStateOfControlsToDocsInProgress()
        {
            fileExportBtn.IsEnabled = false;
            dirExportBtn.IsEnabled = false;
            chooseToExportLabel.IsEnabled = false;  

            progressLabel.Visibility = Visibility.Visible;
            progressTitleLabel.Visibility = Visibility.Visible;
        }
    }
}
