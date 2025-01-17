using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
//using System.Windows.Shapes;
using Novacode;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO.Compression;
using System.Drawing.Imaging;
using System.ComponentModel;

namespace BCFReader
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private List<Note.BCFNote> notes = new List<Note.BCFNote>();
        private SelectedStyles _styles;
        private string _templatePath;
        private string _bcfFilePath;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void loadTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".dotx";
            dlg.Filter = "Word Template (*.dotx)|*.dotx|Word Macro-Enabled Template (*.dotm)|*.dotm";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                this.templatePathTextBox.Text = filename;
                LoadTemplate(filename);
            }

        }

        private void LoadTemplate(string filename)
        {
            //Check if the file exist
            if (File.Exists(filename))
            {
                //Get Styles from the template
                List<DocumentFormat.OpenXml.Wordprocessing.Style> styles = ExtractStyles(filename);

                this.contentStyleCombo.ItemsSource = styles;
                this.contentStyleCombo.SelectedIndex = 0;

                this.dateStyleCombo.ItemsSource = styles;
                this.dateStyleCombo.SelectedIndex = 0;

                this.titleStyleCombo.ItemsSource = styles;
                this.titleStyleCombo.SelectedIndex = 0;

                this.StylesGroupBox.IsEnabled = true;
                _templatePath = filename;
            }
            else
            {
                this.StylesGroupBox.IsEnabled = false;
                _templatePath = "";
            }
        }

        private void loadBCFButton_Click(object sender, RoutedEventArgs e)
        {
            //Fill the list with notes
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".bcfzip";
            dlg.Filter = "Building Collaboration Format(*.bcfzip)|*.bcfzip";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                _bcfFilePath = dlg.FileName;
                this.BCFPathTextBox.Text = _bcfFilePath;

                //Fill the list with notes
                ReadBCF();
            }
        }

        private void SaveAsWordClick(object sender, RoutedEventArgs e)
        {
            //Open a saveFileDialog
            Microsoft.Win32.SaveFileDialog sfd = new Microsoft.Win32.SaveFileDialog();

            // Set filter for file extension and default file extension 
            sfd.DefaultExt = ".docx";
            if (File.Exists(_bcfFilePath))
            {
                sfd.InitialDirectory = Path.GetDirectoryName(_bcfFilePath);
                sfd.FileName = Path.GetFileNameWithoutExtension(_bcfFilePath) + ".docx";
            }

            sfd.Filter = "Word Document (*.docx)|*.docx|Word Macro-Enabled Document (*.docm)|*.docm|Word 97-2003 Document (*.doc)|*.doc";

            // Display SaveFileDialog by calling ShowDialog method 
            Nullable<bool> result = sfd.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = sfd.FileName;
                //this.templatePathTextBox.Text = filename;
                CreateWord(filename);
            }
        }

        private string GetSelectedStyle(ComboBox combo)
        {
            DocumentFormat.OpenXml.Wordprocessing.Style selectedStyle = combo.SelectedItem as DocumentFormat.OpenXml.Wordprocessing.Style;
            if (selectedStyle != null)
            {
                return selectedStyle.StyleId;
            }
            else
            {
                return "";
            }
        }

        private void CreateWord(string fileName)
        {
            pbStatus.Maximum = notes.Count;

            //Retrive styles
            if (_styles == null)
            {
                _styles = new SelectedStyles
    (
    GetSelectedStyle(this.titleStyleCombo),
    GetSelectedStyle(this.dateStyleCombo),
    GetSelectedStyle(this.contentStyleCombo)
    );
            }


            //Create the backgroud Worker
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += worker_WriteBCF;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;


            worker.RunWorkerAsync(fileName);
        }

        void worker_WriteBCF(object sender, DoWorkEventArgs e)
        {
            string fileName = (string)e.Argument;

            // Create a document in memory:
            DocX doc = DocX.Create(fileName);
            if (_templatePath != null)
            {
                doc.ApplyTemplate(_templatePath);
            }

            int i = 1;

            foreach (Note.BCFNote BCFNote in notes)
            {
                //Add note to the report
                AddNoteToReport(BCFNote, doc);
                (sender as BackgroundWorker).ReportProgress(i);
                i++;
            }

            // Save to the output directory:
            doc.Save();
        }

        private void AddNoteToReport(Note.BCFNote BCFNote, DocX doc)
        {
            Novacode.Paragraph p;

            // Insert a paragrpah:
            p = doc.InsertParagraph(BCFNote.Markup.Topic.Title);
            p.StyleName = _styles.TitleStyle;

            doc.InsertParagraph("");

            //Insert the date of the note
            if (BCFNote.Markup.Comment != null)
            {
                p = doc.InsertParagraph("Note created on " + BCFNote.Markup.Comment[0].Date.ToString() + " by " + BCFNote.Markup.Comment[0].Author);
                p.StyleName = _styles.DateStyle;

                p = doc.InsertParagraph("Status : " + BCFNote.Markup.Comment[0].Status + " - " + BCFNote.Markup.Comment[0].VerbalStatus);
                p.StyleName = _styles.DateStyle;

                p = doc.InsertParagraph(BCFNote.Markup.Comment[0].Comment1);
                p.StyleName = _styles.ContentStyle;
            }


            if (BCFNote.Picture != null)
            {
                using (Stream stream = ToStream(BCFNote.Picture, ImageFormat.Png))
                {
                    // Add an Image to the docx file
                    Novacode.Image img = doc.AddImage(stream);
                    Novacode.Picture pic = img.CreatePicture(450, 600);

                    p = doc.InsertParagraph("", false);
                    p.InsertPicture(pic);
                }
            }


            if (BCFNote.Markup.Comment != null)
            {
                int commentCount = BCFNote.Markup.Comment.Count();
                if (commentCount > 1)
                {
                    for (int j = 1; j < commentCount; j++)
                    {
                        p = doc.InsertParagraph("Note created on " + BCFNote.Markup.Comment[j].Date.ToString() + " by " + BCFNote.Markup.Comment[0].Author);
                        p.StyleName = _styles.DateStyle;

                        p = doc.InsertParagraph(BCFNote.Markup.Comment[j].Comment1);
                        p.StyleName = _styles.ContentStyle;
                    }
                }
            }

            p.InsertPageBreakAfterSelf();
        }

        public Stream ToStream(System.Drawing.Image image, ImageFormat formaw)
        {
            var stream = new System.IO.MemoryStream();
            image.Save(stream, formaw);
            stream.Position = 0;
            return stream;
        }

        private List<DocumentFormat.OpenXml.Wordprocessing.Style> ExtractStyles(string fileName)
        {
            List<DocumentFormat.OpenXml.Wordprocessing.Style> styles = new List<DocumentFormat.OpenXml.Wordprocessing.Style>();

            using (var document = WordprocessingDocument.Open(fileName, false))
            {
                var docPart = document.MainDocumentPart;

                StyleDefinitionsPart stylesPart = null;

                stylesPart = docPart.StyleDefinitionsPart;

                if (stylesPart != null)
                {
                    Styles stylesList = stylesPart.Styles;

                    foreach (OpenXmlElement element in stylesList)
                    {
                        if (element.GetType() == typeof(DocumentFormat.OpenXml.Wordprocessing.Style))
                        {
                            DocumentFormat.OpenXml.Wordprocessing.Style style = element as DocumentFormat.OpenXml.Wordprocessing.Style;

                            if (style.Type == "paragraph")
                            {
                                styles.Add(style);
                            }
                        }
                    }
                }

            }

            return styles;
        }

        private void ReadBCF()
        {
            //Get the directory
            DirectoryInfo dirInfo = new DirectoryInfo(System.IO.Path.GetDirectoryName(_bcfFilePath));
            string folderName = System.IO.Path.GetFileNameWithoutExtension(_bcfFilePath);
            DirectoryInfo bcfMainFolder = dirInfo.CreateSubdirectory(folderName);

            System.IO.Compression.ZipFile.ExtractToDirectory(_bcfFilePath, bcfMainFolder.FullName);
            System.Threading.Thread.Sleep(100);
            int count = bcfMainFolder.EnumerateDirectories().Count();
            pbStatus.Maximum = count;
            notes.Clear();

            //Create the backgroud Worker
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += worker_ReadBCF;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;

            worker.RunWorkerAsync(bcfMainFolder);
        }

        void worker_ReadBCF(object sender, DoWorkEventArgs e)
        {

            DirectoryInfo bcfMainFolder = (DirectoryInfo)e.Argument;

            int i = 1;

            foreach (DirectoryInfo dI in bcfMainFolder.EnumerateDirectories())
            {
                notes.Add(new Note.BCFNote(dI));
                (sender as BackgroundWorker).ReportProgress(i);
                i++;
            }

            try
            {
                bcfMainFolder.Delete(true);
            }
            catch
            {
                try
                {
                    bcfMainFolder.Delete(true);
                }
                catch
                {

                }
            }
        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //show the progress bar
            pbStatus.Visibility = System.Windows.Visibility.Visible;

            //update the progress
            pbStatus.Value = e.ProgressPercentage;

            //update the text
            string pbText = string.Format("({0} / {1})", e.ProgressPercentage, notes.Count);

        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //hide the progress bar
            pbStatus.Visibility = System.Windows.Visibility.Hidden;

            int i = notes.Count;

            this.BCFTitle.Text = string.Format("{0} notes have been succesfully loaded.", i);
        }

        private void Button_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            this.Close();
        }

    }

    public class SelectedStyles
    {
        public SelectedStyles(string titleStyle, string dateStyle, string contentStyle)
        {
            this.ContentStyle = contentStyle;
            this.DateStyle = dateStyle;
            this.TitleStyle = titleStyle;
        }

        public string TitleStyle { get; set; }
        public string DateStyle { get; set; }
        public string ContentStyle { get; set; }
    }


}
