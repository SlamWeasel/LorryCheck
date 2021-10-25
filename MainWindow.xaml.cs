using System;
using System.Windows;
using System.Data.SqlClient;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Lastwagen_Abfrage
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        const string connectionString = @"-";
        List<string> LKWs = new List<string>();
        private string TableWhitelist;
        SaveFileDialog excelLocation = new SaveFileDialog();
        private ObservableCollection<LKWLocation> collection;

        public MainWindow()
        {
            InitializeComponent();
            RefreshTable(null);
        }

        private void Refresh_Click_1(object sender, RoutedEventArgs e)
        {
            RefreshTable(TableWhitelist);
        }

        private void Export_Click_1(object sender, RoutedEventArgs e)
        {
            excelLocation.Filter = "Excel File .xls|*.xls";

            if(excelLocation.ShowDialog() == true)
            {
                writeXLS(excelLocation.FileName);
                MessageBox.Show($"Export unter {excelLocation.FileName} gespeichert");
            }
        }

        private void RefreshTable(string whitelist)
        {
            collection = null;
            collection = new ObservableCollection<LKWLocation>();

            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                SqlCommand sqlCommand = whitelist == null ? new SqlCommand(
                    @"" +
                    "SELECT	" +
                    "One.LKW, " +
                    "MAX(Datum) AS LastUpdate," +
                    "(SELECT TOP 1 Ort FROM XXALkwMsg WHERE Datum = (SELECT MAX(One.Datum)) AND LKW = One.LKW) AS Position," +
                    "FORMAT(((SELECT TOP 1 LX FROM XXALkwMsg WHERE Datum = (SELECT MAX(One.Datum)) AND LKW = One.LKW) / 10000.0), 'N4') AS CoordX," +
                    "FORMAT(((SELECT TOP 1 LY FROM XXALkwMsg WHERE Datum = (SELECT MAX(One.Datum)) AND LKW = One.LKW) / 10000.0), 'N4') AS CoordY " +
                    "FROM XXALkwMsg One WHERE LX <> 0 AND LY <> 0 " +
                    "GROUP BY LKW ORDER BY LKW", cn) :
                    new SqlCommand(
                    @"" +
                    "SELECT	TOP 20 " +
                    "One.LKW, " +
                    "Datum AS LastUpdate," +
                    "(SELECT TOP 1 Ort FROM XXALkwMsg WHERE Datum = (SELECT MAX(One.Datum)) AND LKW = One.LKW) AS Position," +
                    "FORMAT(((SELECT TOP 1 LX FROM XXALkwMsg WHERE Datum = (SELECT MAX(One.Datum)) AND LKW = One.LKW) / 10000.0), 'N4') AS CoordX," +
                    "FORMAT(((SELECT TOP 1 LY FROM XXALkwMsg WHERE Datum = (SELECT MAX(One.Datum)) AND LKW = One.LKW) / 10000.0), 'N4') AS CoordY " +
                    $"FROM XXALkwMsg One WHERE LKW = '{whitelist}' " +
                    "GROUP BY LKW, Datum ORDER BY Datum DESC", cn);
                sqlCommand.CommandTimeout = 600;
                SqlDataReader read = sqlCommand.ExecuteReader();

                bool isLKWsEmpty = LKWs.Count == 0;

                while (read.Read())
                {
                    LKWLocation loc = new LKWLocation();

                    string _A = read.IsDBNull(0) ? "" : read.GetString(0);
                    loc.LKW = _A;
                    if (!_A.StartsWith("50") && !_A.StartsWith("60"))
                        continue;
                    if (isLKWsEmpty) LKWs.Add(_A);
                    DateTime _B = read.IsDBNull(1) ? DateTime.MinValue : read.GetDateTime(1);
                    loc.LastUpdatedUtc = _B;
                    var _C = read.IsDBNull(2) ? "" : read.GetValue(2);
                    loc.Position = (string)_C;
                    float _D = read.IsDBNull(3) ? 0 : (float)Convert.ToDouble(read.GetString(3));
                    loc.Latitude = _D;
                    float _E = read.IsDBNull(4) ? 0 : (float)Convert.ToDouble(read.GetString(4));
                    loc.Longitude = _E;

                    collection.Add(loc);

                    Console.WriteLine(loc.ToString());
                }

                cn.Close();
            }

            Table.ItemsSource = collection;
            LKWList.ItemsSource = LKWs;
        }

        private void LKWList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if(LKWList.SelectedItem != null)
                TableWhitelist = LKWList.SelectedItem.ToString();
        }

        private void ClearSelection_Click(object sender, RoutedEventArgs e)
        {
            LKWList.SelectedItem = null;
            TableWhitelist = null;
            RefreshTable(null);
        }

        private void writeXLS(string path)
        {
            Excel.Application xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;

            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }
            xlApp.DisplayAlerts = false;
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;

            xlWorkBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            xlWorkSheet = (Worksheet)xlWorkBook.Sheets.Item[1];

            xlWorkSheet.Cells[1, 1] = "LKW";
            xlWorkSheet.Cells[1, 2] = "LastUpdate";
            xlWorkSheet.Cells[1, 3] = "Position";
            xlWorkSheet.Cells[1, 4] = "Latitude";
            xlWorkSheet.Cells[1, 5] = "Longitude";

            Range RowA = (Range)xlWorkSheet.Columns[1];
            RowA.NumberFormat = "@";
            Range RowB = (Range)xlWorkSheet.Columns[2];
            RowB.NumberFormat = "DD.MM.YYYY hh:mm:ss";

            int cursor = 2;
            foreach(LKWLocation lc in collection)
            {
                xlWorkSheet.Cells[cursor, 1] = lc.LKW.ToString();
                xlWorkSheet.Cells[cursor, 2] = lc.LastUpdatedUtc;
                xlWorkSheet.Cells[cursor, 3] = lc.Position;
                xlWorkSheet.Cells[cursor, 4] = lc.Latitude;
                xlWorkSheet.Cells[cursor, 5] = lc.Longitude;
                cursor++;

                Title = $"{cursor - 1} von {collection.Count} geschrieben";
            }

            Title = "Lastwagen Abfrage";

            xlWorkSheet.Columns.AutoFit();

            xlWorkBook.SaveAs(path, XlFileFormat.xlWorkbookNormal, null, null, null, null, XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
            xlWorkBook.Close(false, null, null);
            xlApp.Quit();
            while (xlApp.Quitting) { }
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
