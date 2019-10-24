using CsvHelper;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using static Excel_Converter.MainWindowDataContext;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Converter
{
    public partial class MainWindow : System.Windows.Window
    {
        public string fileName = string.Empty;
        public bool fileFormattedCorrectly;

        MainWindowDataContext context = new MainWindowDataContext();


        public MainWindow()
        {
            InitializeComponent();

            DataContext = context;
        }


        private void ImportAFile(object sender, RoutedEventArgs e)
        {
            var DialogBox = new OpenFileDialog
            {
                //InitialDirectory = "C:\\Users\\John Scholey\\Downloads\\",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer),
                Filter = "csv file (*.csv)|*.csv",
                FilterIndex = 2,
                RestoreDirectory = true
            };

            if (DialogBox.ShowDialog() == true)
            {
                Excel.Application xlApp = null;
                Worksheet xlWorkSheet = null;

                fileName = DialogBox.FileName;

                try
                {
                    // Create an instance of Excel
                    xlApp = new Excel.Application();
                    xlApp.Workbooks.OpenText(fileName, Comma: true);

                    xlApp.Visible = false;
                    xlApp.DisplayAlerts = false;

                    xlWorkSheet = (Worksheet)xlApp.Worksheets.get_Item(1);
                    Range range = (Range)(xlWorkSheet.UsedRange.Columns[1, Type.Missing]);


                    // Inject headers into CSV file
                    xlWorkSheet.Rows["1:9"].Delete();

                    xlWorkSheet.Cells[1, 1] = "name";
                    xlWorkSheet.Rows["2:3"].Delete();

                    xlWorkSheet.Cells[1, 13] = "au1";
                    xlWorkSheet.Cells[1, 16] = "au2";

                    xlWorkSheet.Cells[1, 19] = "sp1";
                    xlWorkSheet.Cells[1, 23] = "sp2";

                    xlWorkSheet.Cells[1, 26] = "su1";
                    xlWorkSheet.Cells[1, 30] = "su2";

                    xlWorkSheet.SaveAs(fileName);
                }
                catch (Exception error)
                {
                    MessageBox.Show("Error: " + error.Message);
                }
                finally
                {
                    //Completely kill the csv file so it cannot interfer with further modifications.
                    xlApp.Workbooks.Close();
                    xlApp.Quit();

                    if (xlWorkSheet != null)
                    {
                        Marshal.FinalReleaseComObject(xlWorkSheet);
                        xlWorkSheet = null;
                    }

                    if (xlApp != null)
                    {
                        Marshal.FinalReleaseComObject(xlApp);
                        xlApp = null;
                    }

                    //Add pupil details into Observable Collection so it can be displayed

                    try
                    {
                        using (var reader = new StreamReader(File.OpenRead(fileName)))
                        using (var csv = new CsvReader(reader))
                        {
                            var records = csv.GetRecords<dynamic>();

                            csv.Read();
                            csv.ReadHeader();

                            while (csv.Read())
                            {
                                var record = new PupilInfo
                                {
                                    name = csv.GetField<string>("name"),
                                    au1 = csv.GetField<string>("au1"),
                                    au2 = csv.GetField<string>("au2"),
                                    sp1 = csv.GetField<string>("sp1"),
                                    sp2 = csv.GetField<string>("sp2"),
                                    su1 = csv.GetField<string>("su1"),
                                    su2 = csv.GetField<string>("su2")
                                };

                                if (!string.IsNullOrWhiteSpace(record.name))
                                {
                                    CSVFileContents.Add(new FileContents
                                    {
                                        pupilName = record.name.ToString(),
                                        au1 = record.au1.ToString(),
                                        au2 = record.au2.ToString(),

                                        sp1 = record.sp1.ToString(),
                                        sp2 = record.sp2.ToString(),

                                        su1 = record.su1.ToString(),
                                        su2 = record.su2.ToString()
                                    });
                                }
                            }

                            foreach (var pupil in CSVFileContents)
                            {
                                if (pupil.au1.Equals("-"))
                                {
                                    pupil.au1 = " ";
                                }
                                if (pupil.au2.Equals("-"))
                                {
                                    pupil.au2 = " ";
                                }

                                if (pupil.sp1.Equals("-"))
                                {
                                    pupil.sp1 = " ";
                                }
                                if (pupil.sp2.Equals("-"))
                                {
                                    pupil.sp2 = " ";
                                }

                                if (pupil.su1.Equals("-"))
                                {
                                    pupil.su1 = " ";
                                }
                                if (pupil.su2.Equals("-"))
                                {
                                    pupil.su2 = " ";
                                }
                            }

                            // Split O'Track data into chunks to display back correct subject
                            var subjectPicker = DataSetPicker.SelectedIndex;
                            var counter = 0;
                            var key = CSVFileContents[0].pupilName;

                            try
                            {
                                Subject.Clear();

                                switch (subjectPicker)
                                {
                                    case 0:

                                        foreach (var pupil in CSVFileContents)
                                        {
                                            if (pupil.pupilName.Equals(key))
                                            {
                                                counter++;

                                                if (counter > 1)
                                                {
                                                    break;
                                                }
                                            }

                                            Subject.Add(pupil);
                                        }

                                        break;

                                    case 1:
                                        // TODO Add logic for adding a kid FROM second key

                                        foreach (var pupil in CSVFileContents)
                                        {
                                            if (pupil.pupilName.Equals(key) && counter > 1)
                                            {
                                                counter++;

                                                if (counter > 2)
                                                {
                                                    break;
                                                }
                                            }

                                            Subject.Add(pupil);
                                        }

                                        break;

                                    case 2:

                                        break;

                                    case 3:
                                        // TODO add logic for adding a kid FROM third key

                                        break;
                                }
                            }

                            catch (Exception error)
                            {
                                MessageBox.Show(error.Message);
                            }
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Unable to open the file - "
                            + "please ensure  you do not have it open in Excel and that the headers have been added to the file.");
                    }
                }
            }
            else
            {
                MessageBox.Show("Please choose a file");
            }
        }


        private void ExportToMag(object sender, RoutedEventArgs e)
        {
            var cellRow = 2;
            var pupilCounter = 0;

            string au1Cell = string.Empty;
            string au2Cell = string.Empty;
            string sp1Cell = string.Empty;
            string sp2Cell = string.Empty;
            string su1Cell = string.Empty;
            string su2Cell = string.Empty;

            var YGValue = YGPicker.SelectedIndex;

            var DialogBox = new OpenFileDialog
            {
                InitialDirectory = "C:\\Users\\John Scholey\\Downloads\\",
                Filter = "xls file (*.xls)|*.xls",
                FilterIndex = 2,
                RestoreDirectory = true
            };

            if (DialogBox.ShowDialog() == true)
            {
                Excel.Application xlApp = null;
                Worksheet xlWorkSheet = null;

                fileName = DialogBox.FileName;

                try
                {
                    //Set up new Excel instance
                    xlApp = new Excel.Application();
                    xlApp.Workbooks.OpenText(fileName, Comma: true);

                    xlApp.Visible = true;
                    xlApp.DisplayAlerts = false;

                    xlWorkSheet = (Worksheet)xlApp.Worksheets.get_Item(1);



                    switch(YGValue)
                    {
                        case 0:

                            au1Cell = "G";
                            au2Cell = "H";
                            sp1Cell = "I";
                            sp2Cell = "J";
                            su1Cell = "K";
                            su2Cell = "L";
                            
                            break;

                        case 1:

                            au1Cell = "N";
                            au2Cell = "O";
                            sp1Cell = "P";
                            sp2Cell = "Q";
                            su1Cell = "R";
                            su2Cell = "S";

                            break;

                        case 2:

                            au1Cell = "U";
                            au2Cell = "V";
                            sp1Cell = "W";
                            sp2Cell = "X";
                            su1Cell = "Y";
                            su2Cell = "Z";

                            break;

                        case 3:

                            au1Cell = "AB";
                            au2Cell = "AC";
                            sp1Cell = "AD";
                            sp2Cell = "AE";
                            su1Cell = "AF";
                            su2Cell = "AG";

                            break;

                        case 4:

                            au1Cell = "AI";
                            au2Cell = "AJ";
                            sp1Cell = "AK";
                            sp2Cell = "AL";
                            su1Cell = "AM";
                            su2Cell = "AN";

                            break;

                        case 5:

                            au1Cell = "AP";
                            au2Cell = "AQ";
                            sp1Cell = "AR";
                            sp2Cell = "AS";
                            su1Cell = "AT";
                            su2Cell = "AU";

                            break;
                    }

                    foreach (var pupil in Subject)
                    {
                        xlWorkSheet.Range[au1Cell + cellRow].Value = Subject[pupilCounter].au1.ToString();
                        xlWorkSheet.Range[au2Cell + cellRow].Value = Subject[pupilCounter].au2.ToString();

                        xlWorkSheet.Range[sp1Cell + cellRow].Value = Subject[pupilCounter].sp1.ToString();
                        xlWorkSheet.Range[sp2Cell + cellRow].Value = Subject[pupilCounter].sp2.ToString();

                        xlWorkSheet.Range[su1Cell + cellRow].Value = Subject[pupilCounter].su1.ToString();
                        xlWorkSheet.Range[su2Cell + cellRow].Value = Subject[pupilCounter].su2.ToString();

                        cellRow++;
                        pupilCounter++;
                    }
                }
                catch (Exception error)
                {
                    MessageBox.Show(error.Message);
                }
            }
        }


        private void Check(object sender, RoutedEventArgs e)
        {
            var sum = new Summatives();
            var yg = YGPicker.Text;

            if (!String.IsNullOrWhiteSpace(tbEm.Text) && !String.IsNullOrWhiteSpace(tbDev.Text)
                && !String.IsNullOrWhiteSpace(tbSec.Text) && !String.IsNullOrWhiteSpace(tbGD.Text))
            {
                sum.emerging = yg + " " + tbEm.Text;
                sum.developing = yg + " " + tbDev.Text;
                sum.secure = yg + " " + tbSec.Text;
                sum.greaterDepth = yg + " " + tbGD.Text;

                Console.WriteLine(sum.emerging + "\n" + sum.developing + "\n" + sum.secure + "\n" + sum.greaterDepth);

                foreach (var item in CSVFileContents)
                {
                    if (!item.Equals(sum.emerging) || !item.Equals(sum.developing) || !item.Equals(sum.secure) || !item.Equals(sum.greaterDepth))
                    {
                        Console.WriteLine(item.ToString());
                    }
                }
            }
            else
            {
                MessageBox.Show("Please ensure all the Summatives are filled in the Preferences tab.");
            }
        }
    }
}
