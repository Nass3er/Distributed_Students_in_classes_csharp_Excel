using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Data.OleDb;

namespace DistributedStudents
{
    public partial class FinalForm : Form
    {

        string connectionString;
        public FinalForm()
        {
            InitializeComponent();
            openExcelFile();
            //// Get the JSON string from settings
            //string json = Properties.Settings.Default.finalReport;

            //// Deserialize the JSON string back to a List<DepartmentInfo>
            //List<DepartmentInfo> departmentList = JsonConvert.DeserializeObject<List<DepartmentInfo>>(json);

            //// Set the DataSource of your DataGridView
            //dataGridView1.DataSource = departmentList;


        }
        //public class DepartmentInfo
        //{
        //    public DateTime التاريخ { get; set; }
        //    public string اليوم { get; set; }
        //    public string م { get; set; }
        //    public string التخصصات { get; set; }
        //    public string المادة { get; set; }
        //    public int عدد_الطلاب { get; set; }
        //    public int عدد_الحضور { get; set; }
        //    public int عدد_الغياب { get; set; }
        //    public string ملاحظات { get; set; }
        //}
        private void openExcelFile()
        {
            string sheetName2 = "sheet1";
            string filePath = Properties.Settings.Default.finalFilePath;

            try
            {
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("الملف غير موجود او انه لم يتم إضافته في شاشة تهيئة متغيرات النظام, الرجاء إضافة مسار الملف النهائي!!");
                    return;
                }

                string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    DataTable dataTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    

                    bool sheetExists = false;
                    foreach (DataRow row in dataTable.Rows)
                    {
                        string sheetName = row["TABLE_NAME"].ToString();
                        if (sheetName == sheetName2)
                        {
                            sheetExists = true;
                            break;
                        }
                    }

                    if (!sheetExists)
                    {
                        // Create the sheet
                        OleDbCommand createCommand = connection.CreateCommand();
                        createCommand.CommandText = $"CREATE TABLE [{sheetName2}] ([Column1] TEXT, [Column2] TEXT, [Column3] TEXT)";
                        createCommand.ExecuteNonQuery();

                        //MessageBox.Show($"The sheet '{sheetName2}' has been added to the file.");
                    }

                    string query = $"SELECT * FROM [{sheetName2}$]";
                    OleDbCommand command = new OleDbCommand(query, connection);
                    OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                    DataTable data = new DataTable();
                    adapter.Fill(data);

                    dataGridView1.DataSource = data;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"حدث خظأ: {ex.Message}");
            }
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void ExportToExcel(DataGridView dataGridView)
        {
            // Create a new Excel application
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            // Create a new workbook
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add();

            // Create a new worksheet
            Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.ActiveSheet;

            try
            {
                // Export the column headers
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1] = dataGridView.Columns[i].HeaderText;
                }

                // Export the data rows
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView.Rows[i].Cells[j].Value?.ToString();
                    }
                }

                // Display the SaveFileDialog to the user
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Files|*.xlsx";
                saveFileDialog.Title = "Save Excel File";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string outputPath = saveFileDialog.FileName;

                    // Save the workbook to the selected output path
                    workbook.SaveAs(outputPath);
                    workbook.Close();

                    MessageBox.Show("Data exported to Excel successfully!");
                }
                else
                {
                    // User canceled the save operation
                    workbook.Close(false);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Clean up Excel objects
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            button1.BackColor = Color.White;
            button1.ForeColor = Color.Black;
             
            ExportToExcel(dataGridView1);
           
            button1.BackColor = Color.Navy;
            button1.ForeColor = Color.White;
        }

        private void button4_Click(object sender, EventArgs e)
        {
        }
        private void UpdateExcelFileFromDataGridView(string filePath, DataGridView dataGridView)
        {
            // Load the existing Excel file
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // Get the data from the DataGridView
                DataTable dataGridViewData = GetDataGridViewData(dataGridView);

                // Update the corresponding cells in the worksheet
                UpdateWorksheetWithData(worksheet, dataGridViewData);

                // Save the changes to the Excel file
                package.Save();
            }
        }

        private DataTable GetDataGridViewData(DataGridView dataGridView)
        {
            DataTable dataTable = new DataTable();

            // Create columns in the DataTable based on the DataGridView columns
            foreach (DataGridViewColumn dgvColumn in dataGridView.Columns)
            {
                dataTable.Columns.Add(dgvColumn.Name);
            }

            // Add rows to the DataTable based on the DataGridView rows
            foreach (DataGridViewRow dgvRow in dataGridView.Rows)
            {
                DataRow dataRow = dataTable.NewRow();

                // Populate the DataRow with cell values from the DataGridView
                foreach (DataGridViewCell dgvCell in dgvRow.Cells)
                {
                    object cellValue = dgvCell.Value ?? DBNull.Value;
                    dataRow[dgvCell.ColumnIndex] = cellValue;
                }

                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }

        private void UpdateWorksheetWithData(ExcelWorksheet worksheet, DataTable dataTable)
        {
            // Clear existing data in the worksheet
            worksheet.Cells.Clear();

            // Write the new data from the dataTable starting at cell A1
            worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);
        }

        

        private void button2_Click(object sender, EventArgs e)
        {
            button2.BackColor = Color.White;
            button2.ForeColor = Color.Black;
  
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial; // or LicenseContext.NonCommercial
            UpdateExcelFileFromDataGridView(Properties.Settings.Default.finalFilePath, dataGridView1);
            MessageBox.Show("تم تحديث الملف بنجاااح");

            
            button2.BackColor = Color.Navy;
            button2.ForeColor = Color.White;
            
        }

        private void FinalForm_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Assuming you have a DataGridView named dataGridView
            int[] columnIndices = { 5, 6, 7 }; // Specify the column indices you want to calculate the sum for

            decimal[] sums = new decimal[columnIndices.Length];

            // Iterate over the selected rows
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                // Iterate over the specified column indices
                for (int i = 0; i < columnIndices.Length; i++)
                {
                    int columnIndex = columnIndices[i];

                    // Retrieve the value from the specified column and convert it to a decimal
                    if (row.Cells[columnIndex].Value != null && decimal.TryParse(row.Cells[columnIndex].Value.ToString(), out decimal cellValue))
                    {
                        // Add the cell value to the sum for the current column index
                        sums[i] += cellValue;
                    }
                }
            }

            int rowIndexToInsert = dataGridView1.SelectedRows[dataGridView1.SelectedRows.Count - 1].Index + 1;

            // Shift rows below the selected rows to create an empty row
            DataTable dataTable = (DataTable)dataGridView1.DataSource;

            // Create a new empty row
            DataRow newRow = dataTable.NewRow();

            // Insert the empty row at the desired index
            dataTable.Rows.InsertAt(newRow, rowIndexToInsert);

            // Display the sums in the newly created empty row
            for (int i = 0; i < columnIndices.Length; i++)
            {
                int columnIndex = columnIndices[i];
                decimal sum = sums[i];

                newRow[columnIndex] = sum;
            }

            // Scroll to the newly inserted row
            dataGridView1.FirstDisplayedScrollingRowIndex = rowIndexToInsert;

            // Refresh the DataGridView to reflect the changes
            dataGridView1.Refresh();
            //\\\\\\\\\\\\\\\\\\
 

        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            
                int columnIndex5 = 5;
                int columnIndex6 = 6;
                int columnIndex7 = 7;

                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];

                if (row.Index >= 0 && row.Cells.Count > columnIndex5 && row.Cells.Count > columnIndex6 && row.Cells.Count > columnIndex7)
                {
                    // Get the values from column index 5 and column index 6
                    if (int.TryParse(row.Cells[columnIndex5].Value?.ToString(), out int value5) &&
                        int.TryParse(row.Cells[columnIndex6].Value?.ToString(), out int value6))
                    {
                        if (value6 <= value5)
                        {
                            // Unsubscribe from the CellValueChanged event temporarily
                            dataGridView1.CellValueChanged -= dataGridView1_CellValueChanged;

                            // Calculate the value for column index 7
                            int value7 = value5 - value6;

                            // Set the calculated value in column index 7
                            row.Cells[columnIndex7].Value = value7.ToString();

                            // Subscribe back to the CellValueChanged event
                            dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
                        }
                        else
                        {
                            row.Cells[columnIndex6].Value = value5.ToString();
                        }
                    }
                }
            

        }
    }

}
