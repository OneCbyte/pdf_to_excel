using Adobe.PDFServicesSDK.auth;
using Adobe.PDFServicesSDK.io;
using Adobe.PDFServicesSDK.options.exportpdf;
using Adobe.PDFServicesSDK.pdfops;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.IO;

namespace pdf_to_excel
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Handling button clicks
            Button button = (Button)sender;
            DoComplexTask(button);
        }
        async void DoComplexTask(object sender)
        {
            Button button = (Button)sender;
            string filename = Text2.Text;
            string path = Text1.Text;

            if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\" + filename) == true)
            {
                MessageBox.Show("Attention! There is already a file with the same name. Rename it!");
            }
            else if (filename.Contains(".xlsx") && File.Exists(path))
            {
                button.IsEnabled = false;
                button.Content = "CONVERTING...";
                await Task.Run(() => Do(filename, path));
                
                MessageBox.Show("The process is complete! The file was saved to the desktop.");
                button.IsEnabled = true;
                button.Content = "PRESS TO START...";
            }
            else if (File.Exists(path) == false)
            {
                Text1.Background = new SolidColorBrush(Color.FromRgb(247, 4, 4));
                MessageBox.Show("Pdf file not found!");
            }
            else if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\output.xlsx") == true)
            {
                MessageBox.Show("Attention! There is already a file with the same name. Rename it!");
            }
            else if (filename.Contains(".xlsx") == false && File.Exists(path))
            {
                button.IsEnabled = false;
                button.Content = "CONVERTING...";
                MessageBox.Show("The file will be saved as output.xlsx");
                await Task.Run(() => Do("output.xlsx", path));
                
                MessageBox.Show("The process is complete! The file was saved to the desktop.");
                button.IsEnabled = true;
                button.Content = "PRESS TO START...";
            }
        }
        private void Do(string file, string path)
        {
            // Here Adobe converts .pdf to .xlsx using your credentials
            Credentials credentials = Credentials.ServiceAccountCredentialsBuilder()
            .FromFile("file.json")
                                     .Build();
            Adobe.PDFServicesSDK.ExecutionContext executionContext = Adobe.PDFServicesSDK.ExecutionContext.Create(credentials);
            ExportPDFOperation exportPDFOperation = ExportPDFOperation.CreateNew(ExportPDFTargetFormat.XLSX);
            FileRef fileRef = FileRef.CreateFromLocalFile(path);
            exportPDFOperation.SetInput(fileRef);
            FileRef result = exportPDFOperation.Execute(executionContext);
            result.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\" + file);
        }
        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Change the color of the input field
            Text1.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
        }
    }
}