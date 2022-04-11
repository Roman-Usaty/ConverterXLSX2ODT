using System.Windows;
using GroupDocs.Conversion;
using GroupDocs.Conversion.Options.Convert;
using Microsoft.Win32;
using GroupDocs.Conversion.FileTypes;

namespace ConverterXLSX2ODT
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string _filePathXLSX;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenDocument_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();

            openFile.DefaultExt = "*.xlsx";
            openFile.Multiselect = false;
            openFile.Filter = "файл Excel(*.xlsx) | *.xlsx";
            openFile.Title = "Выберите файл таблицы";

            if (!(bool)openFile.ShowDialog())
            {
                return;
            }

            if (openFile == null || openFile.FileNames.Length <= 0) return;
            
            _filePathXLSX = openFile.FileName;
        }

        private void Convert_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.DefaultExt = "*.odt";
            saveFileDialog.Filter = "файл ODT (*.odt) | *.odt";
            saveFileDialog.RestoreDirectory = true;

            if (!(bool)saveFileDialog.ShowDialog()) return;
            if (saveFileDialog.FileNames.Length <= 0) return;

            using Converter converter = new Converter(_filePathXLSX);

            var options = new WordProcessingConvertOptions();
            options.Format = WordProcessingFileType.Odt;
            converter.Convert(saveFileDialog.FileName, options);
        }
    }
}
