using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

using Microsoft.Win32;
using System.IO;
using System.Diagnostics;
using System.Net;



namespace LocalChecker_1._0
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        String toolDirectory = AppDomain.CurrentDomain.BaseDirectory;
       
        public MainWindow()
        {
            InitializeComponent();
            queriesTextBox.AddHandler(RichTextBox.DragOverEvent, new DragEventHandler(queriesRichTextBox_DragOver), true);
            queriesTextBox.AddHandler(RichTextBox.DropEvent, new DragEventHandler(queriesRichTextBox_Drop), true);
        }

         

        
        private void queriesImportButton_Click(object sender, RoutedEventArgs e)
        {
            
            queriesImportCheckBox.IsChecked = false;
            OpenFileDialog theDataImportDialogBox = new OpenFileDialog();
            theDataImportDialogBox.Title = "Open Text File";
            theDataImportDialogBox.DefaultExt = ".txt";
            theDataImportDialogBox.Filter = "TXT files|*.txt";
            theDataImportDialogBox.InitialDirectory = @"C:\";

            try
            {


                if (theDataImportDialogBox.ShowDialog() == true)
                {
                    String dataImportDialogBoxFilename = theDataImportDialogBox.FileName;
                    TextRange t = new TextRange(queriesTextBox.Document.ContentStart, queriesTextBox.Document.ContentEnd);
                    using (FileStream importedFile = new FileStream(dataImportDialogBoxFilename, FileMode.Open))
                    {
                        t.Load(importedFile, System.Windows.DataFormats.Text);
                    }
                    queriesImportCheckBox.IsChecked = true;

                    String richTextBoxContent = new TextRange(queriesTextBox.Document.ContentStart, queriesTextBox.Document.ContentEnd).Text;
                    if (richTextBoxContent == "")
                    {
                        MessageBox.Show("oops");
                        // queriesImportCheckBox.IsEnabled = false;
                    }
                }
                else
                {
                    MessageBox.Show("You have cancelled the operation");
                }
            }
            catch (UnauthorizedAccessException)
            {
                MessageBox.Show("The selected file is protected and can't be open by this tool! Select another file!");
            }

        }

        private String[] checkFeatureOnBing(string[] aFile, string aMarket, string aFeatureCode, string aFileName, string aLocation)
        {
 
            string searchEngineStart = "http://www.bing.";
            string searchEngineEnd = "/search?q=";
            string[] linesResults;
            string fileLocation = toolDirectory + aFileName + aMarket + ".txt";

            using (StreamWriter queriesCheckedResutsText = new StreamWriter(fileLocation))
            {
                foreach (string aLine in aFile)
                {
                    if (aLine.Length > 0)
                    {
                        string queryPath = searchEngineStart + aMarket + searchEngineEnd + aLine + aLocation;
                        WebClient myWebClient = new WebClient();
                        string htmlCode = myWebClient.DownloadString(queryPath);
                        if (htmlCode.Contains(aFeatureCode))
                        {
                            queriesCheckedResutsText.WriteLine("\"=HYPERLINK(\"\"" + queryPath + "\"\",\"\"" + aLine + "\"\")\"" + "," + "yes");
                        }
                        else
                        {
                            queriesCheckedResutsText.WriteLine("\"=HYPERLINK(\"\"" + queryPath + "\"\",\"\"" + aLine + "\"\")\"" + "," + "no");
                        }

                    }
                    
                }
            }
            return linesResults = File.ReadAllLines(fileLocation);
        }



        private void createExcelWithHeadersAndSave(string[] aFile, string aMarket, string aFileName, string aFeatureName)
        {
            SaveFileDialog theSaveExcelFileDialogBox = new SaveFileDialog();
            theSaveExcelFileDialogBox.InitialDirectory = @"C:\";
            theSaveExcelFileDialogBox.Title = "Save results";
            theSaveExcelFileDialogBox.DefaultExt = ".csv";
            theSaveExcelFileDialogBox.Filter = "CSV Files (.csv)|*.csv";
            theSaveExcelFileDialogBox.FileName = aFileName + aMarket;
            if (theSaveExcelFileDialogBox.ShowDialog() == true)
            {
                try
                {
                    using (StreamWriter headers = File.CreateText(theSaveExcelFileDialogBox.FileName))
                    {
                        headers.WriteLine(string.Format("{0}, {1}", "Query", aFeatureName));
                    }
                    using (StreamWriter queriesCheckedResultsCSV = new StreamWriter(theSaveExcelFileDialogBox.FileName, true, Encoding.Unicode))
                    {
                        foreach (string aLine in aFile)
                        {
                            queriesCheckedResultsCSV.WriteLine(aLine);
                        }
                    }

                    MessageBoxResult openExcelFile = MessageBox.Show("Do you want to open the saved file?", "Option to open the saved file", MessageBoxButton.YesNo);
                    if (openExcelFile == MessageBoxResult.Yes)
                    {
                        Process.Start(theSaveExcelFileDialogBox.FileName);
                    }
                    else
                    {
                        MessageBox.Show("The results have been saved in the Excel file");
                    }

                    //Application.Current.Shutdown();
                }
                catch (IOException)
                {
                    MessageBox.Show("A file with the same name is already opened on your computer! Close the file and press run again.");
                }
                
             
            }
            else
            {
                MessageBox.Show("You have cancelled the save operation.");
            }

            File.Delete(toolDirectory + aFileName + aMarket + ".txt");

        }

        private void runButton_Click(object sender, RoutedEventArgs e)
        {

            String marketCode = marketSelector.SelectionBoxItem.ToString();
            String featureName = featureSelector.SelectionBoxItem.ToString();
            string locationName = locationSelector.SelectionBoxItem.ToString();
            string marketPrefix = "";
            string featureCode = "";
            string savedFileName = "";
            string locationCode = "";

             if (marketCode == "fr-fr")
                {
                    marketPrefix = "fr";
                }
                else if (marketCode == "de-de")
                {
                    marketPrefix = "de";
                }
                else if (marketCode == "en-gb")
                {
                    marketPrefix = "co.uk";
                }
                else if (marketCode == "it-it")
                {
                    marketPrefix = "it";
                }
                else if (marketCode == "ru-ru")
                {
                 marketPrefix = "ru";
                }
             else if (marketCode == "es-es")
             {
                 marketPrefix = "es";
             }


             if (featureName == "Local Answer")
             {
                 featureCode = "local?lid";
                 savedFileName = "localQueriesChecked_";
             }
             else if (featureName == "Snapshot (Taskpane)")
             {
                 featureCode = "TaskPane";
                 savedFileName = "snapshotChecked_";
             }
             else if (featureName == "Weather Answer")
             {
                 featureCode = "WeatherAnswer+Provider";
                 savedFileName = "weatherAnswerQueriesChecked_";
             }
             else if (featureName == "News Answer")
             {
                 featureCode = "bing.com/news";
                 savedFileName = "newsAnswerQueriesChecked_";
             }

             if (locationName == "Paris")
             {
                 //the location used the long and lat of the cities (in decimal degré)
                 locationCode = "&location=lat:48.85%3blong:2.34";
             }
             else if (locationName == "Munich")
             {
                 locationCode = "&location=lat:48.13%3blong:11.57";
             }
             else if (locationName == "London")
             {
                 locationCode = "&location=lat:51.50%3blong:-0.12";
             }
             else if (locationName == "Rome")
             {
                 locationCode = "&location=lat:41.89%3blong:12.51";
             }
             else if (locationName == "Moscow")
             {
                 locationCode = "&location=lat:55.75%3blong:37.61";
             }
             else if (locationName == "Madrid")
             {
                 locationCode = "&location=lat:40.41%3blong:-3.70";
             }


             if (locationCode == "")
             {
                 MessageBoxResult selectLocation = MessageBox.Show("You haven't selected a location! Do you want to select a location? Click YES to select a location or NO to go on with the current ip in Bing or any flight you are on.","Location selection", MessageBoxButton.YesNo);
                 if (selectLocation == MessageBoxResult.Yes)
                 {
                     MessageBox.Show("Select a location!");
                 }
             }
            

                if ((marketPrefix == "") || (featureCode == ""))
                {
                    MessageBoxOptions.RightAlign.ToString();
                    MessageBox.Show("You haven't selected a market or a feature to test! Please selected a market and a feature to test!");
                }

            else
                 {

                TextRange t = new TextRange(queriesTextBox.Document.ContentStart, queriesTextBox.Document.ContentEnd);
                //splitting the content of the RicthTextBox in order to extrect the lines.
                string[] queriesTextBoxLines = t.Text.Split('\n');
                //the lines are coming with the '\n', which is affecting the final results in the Excel file. We will run a foreach to try to get rid of the '\n'.
                List<string> queriesRichTextBoxLines = new List<string>();
                foreach (string aLine in queriesTextBoxLines)
                {
                    queriesRichTextBoxLines.Add(aLine.TrimEnd('\r', '\n'));
                }

                string[] queriesRichTextBoxLinesTrimed = queriesRichTextBoxLines.ToArray();
                string[] queriesCheckedOnBing = checkFeatureOnBing(queriesRichTextBoxLinesTrimed, marketPrefix, featureCode, savedFileName, locationCode);
                createExcelWithHeadersAndSave(queriesCheckedOnBing, marketPrefix, savedFileName, featureName);

                 }
            
        }

        private void queriesTextBox_TextChanged(object sender, TextChangedEventArgs e)
        { 

        }

        private void queriesRichTextBox_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] doctPath = (string[])e.Data.GetData(DataFormats.FileDrop);
                var dataFormat = DataFormats.Rtf;

                if (e.KeyStates == DragDropKeyStates.ShiftKey)
                {
                    dataFormat = DataFormats.Text;
                }

                if (File.Exists(doctPath[0]))
                {
                    try
                    {
                        TextRange tDrop = new TextRange(queriesTextBox.Document.ContentStart, queriesTextBox.Document.ContentEnd);
                        using (Stream streamDrop = new FileStream(doctPath[0], FileMode.OpenOrCreate))
                        {
                            tDrop.Load(streamDrop, DataFormats.Text);
                        }
                        queriesImportCheckBox.IsChecked = true;

                    }

                    catch (SystemException)
                    {
                        MessageBox.Show("File could not be opened. Make sure the file is a text file");
                    }
                }
            }
        }

        private void queriesRichTextBox_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.All;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }

            e.Handled = false;
        }

        private void marketSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

    }
}
