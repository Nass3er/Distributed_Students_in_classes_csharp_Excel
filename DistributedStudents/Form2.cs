using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace DistributedStudents
{
    public partial class Form2 : Form
    {
        DataTable dataTable1;
        

        public Form2()
        {
            InitializeComponent();
            dataGridView1.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells); // Adjust row heights based on cell contents
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None; // Disable automatic row height adjustment

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.Height = 50; // Set the desired height for each row
            }
        }
        string connectionString;
        string sheetName = "sheet1";



        private void btnUpdateClasses_Click(object sender, EventArgs e)
        {

        }


        private void btnUpdate_Click(object sender, EventArgs e)
        {


        }


        private void Form2_Load(object sender, EventArgs e)
        {

        }
        string filePath;
        string filePath2;
        private void btnUpdateClasses_Click_1(object sender, EventArgs e)
        {
            // Assuming the 'رقم الجلوس' column index is 2
             //dataGridView1.Columns.Add("SerialNumber", "م");
           // int selectedNumber = (int)numericUpDown1.Value; // Retrieve the numeric value from the NumericUpDown control
            //Display the OpenFileDialog to the user
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx";
            openFileDialog.Title = "Select Excel File";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                Properties.Settings.Default.FilePath = filePath;
                Properties.Settings.Default.Save();
                 
                connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    try
                    {
                        connection.Open();

                        string query = $"SELECT  `اسم الطالب`,`رقم الجلوس`,`اسم المادة`,`التخصص`,`توقيع الحضور`,`توقيع التسليم` FROM [{sheetName}$]";  //  WHERE اللجنة = ?
                        OleDbCommand command = new OleDbCommand(query, connection);
                        //command.Parameters.AddWithValue("@NumericValue", selectedNumber);
                         command = new OleDbCommand(query, connection);

                        OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                        DataTable dataTable1 = new DataTable();
                        adapter.Fill(dataTable1);

                        dataGridView1.DataSource = dataTable1;

                        //Loop through each row in the DataGridView and set the serial number
                        //for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        //{
                        //    dataGridView1.Rows[i].Cells["SerialNumber"].Value = (i + 1).ToString();
                        //}
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"An error occurred: {ex.Message}");
                    }
                    finally
                    {
                        connection.Close();
                    }
                     
                }
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
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

        private void ExportToWord()
        {

            try
            {
                // Create a new file dialog to choose the save location
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Word Document (*.docx)|*.docx";
                saveFileDialog.Title = "Save Word Document";
                saveFileDialog.ShowDialog();

                if (saveFileDialog.FileName != "")
                {
                    string savePath = saveFileDialog.FileName;

                    // Create a new Word document
                    Word.Application wordApp = new Word.Application();
                    Word.Document doc = wordApp.Documents.Add();

                    // Set the layout margin to narrow
                    doc.PageSetup.LeftMargin = wordApp.InchesToPoints(0.5f);
                    doc.PageSetup.RightMargin = wordApp.InchesToPoints(0.5f);
                    doc.PageSetup.TopMargin = wordApp.InchesToPoints(0.5f);
                    doc.PageSetup.BottomMargin = wordApp.InchesToPoints(0.5f);

                    // Get the DataGridView data
                    DataTable dataTable = (DataTable)dataGridView1.DataSource;

                    // Set the title in the header
                    Word.Section section = doc.Sections.First;
                    Word.HeaderFooter header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    Word.Range headerRange = header.Range;

                    // Add the image to the header
                    string imagePath = Properties.Settings.Default.ImageHeaderPath;  // replace with your image path
                    headerRange.InlineShapes.AddPicture(imagePath);
                     
                    // Add the label after the image in the center
                    Word.Paragraph labelParagraph = headerRange.Paragraphs.Add();
                    Word.Range labelRange = labelParagraph.Range;
                    labelRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    labelRange.Text = "رقم اللجنة: " + (int)numericUpDown1.Value;
                    labelRange.Font.Bold = -1;
                    labelRange.Font.Size = 30;
                    labelRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    // Create a Word table
                    Word.Table table = doc.Tables.Add(doc.Range(), dataTable.Rows.Count + 1, dataTable.Columns.Count);

                    // Set table style and borders
                    table.set_Style("Table Grid");
                    table.Borders.Enable = 1; // Enable table borders
                     
                    // Set table headers
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        table.Cell(1, i + 1).Range.Text = dataTable.Columns[i].ColumnName;
                    }
                   // Set the width of the column you want to decrease
                    float targetColumnWidth = 1.1f; // Adjust this value as needed
                    float targetColumnWidthunmber = 0.8f; // Adjust this value as needed
                    float targetColumnWidthName = 2.0f; // Adjust this value as needed
                    float targetColumnWidthdept = 3.0f;

                    // Get the column index of the "رقم الجلوس" column
                    int targetColumnIndex = 3; // Assuming it's the second column, adjust as needed
                    int targetColumnIndexNumber = 1;
                    int targetColumnIndexName = 2;
                    int targetColumnIndexdept = 5;
                    // Set the width of the target column
                    table.Columns[targetColumnIndex].Width = targetColumnWidth * 62; // Multiply by 72 to convert from inches to points
                    table.Columns[targetColumnIndexNumber].Width = targetColumnWidthunmber * 38;
                    table.Columns[targetColumnIndexName].Width = targetColumnWidthName * 68;
                    table.Columns[targetColumnIndexdept].Width = targetColumnWidthdept * 30;

                    // Calculate the total number of cells to fill
                    int totalCells = dataTable.Rows.Count * dataTable.Columns.Count;

                    // Create a counter variable to track the number of filled cells
                    int filledCells = 0;

                    // Fill table with data
                    for (int row = 0; row < dataTable.Rows.Count; row++)
                    {
                        for (int col = 0; col < dataTable.Columns.Count; col++)
                        {
                            Word.Cell cell = table.Cell(row + 2, col + 1);
                            cell.Range.Text = dataTable.Rows[row][col].ToString();
                            cell.Range.Font.Bold = -1;

                            // Increment the filled cells counter
                            filledCells++;

                            // Calculate the progress as a percentage
                            double progressPercentage = (double)filledCells / totalCells * 100;

                            // Update a progress label or display the progress percentage in any other way
                            // For example, if you have a label named progressLabel:
                            progressLabel.Text = $"Progress: {progressPercentage:F2}%";

                            // Allow the UI to update by invoking Application.DoEvents()
                            Application.DoEvents();
                        }
                    }

                    table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);

                    //Get te end of the document range
                    Word.Range endRange = doc.Range(doc.Content.End - 1);

                    // Array of signature labels
                    string[] signatureLabels = { "الملاحظ 1", "الملاحظ 2" };
                    string ignaturelable2 = "نائب رئيس المركز الاختباري";
                    string ignaturelable3 = "رئيس المركز الاختباري";

                    // Add labels for signatures
                    endRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

                     
                        endRange.Text += $"{signatureLabels[0]}: ____________________ \t\t\t\t" + $"{signatureLabels[1]}: ____________________ \n";
                     
                    endRange.Text += $"{ignaturelable2}: ____________________ \t\t" + $"{ignaturelable3}: ____________________ " + "  ";
                     
                    endRange.Font.Bold = 5;
                    endRange.Font.Size = 26;

                    // Select all content in the document
                    Word.Range range = doc.Content;
                    range.Select();

                    // Make the selection bold
                    Word.Selection selection = wordApp.Selection;
                    selection.Font.Bold = 1;

                    // Save the Word document
                    doc.SaveAs2(savePath);

                    // Close the Word document and application
                    doc.Close();
                    wordApp.Quit();

                    MessageBox.Show("Data exported to Word successfully!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }

 
        }
         

        private static string GetEmbeddedImage(string resourceName)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string[] resourceNames = assembly.GetManifestResourceNames();

            string matchingResourceName = resourceNames.FirstOrDefault(name => name.EndsWith(resourceName));
            return matchingResourceName;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView1);
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button2.BackColor = Color.White;
            button2.ForeColor = Color.Black;
            
            ExportToWord();
 
            button2.BackColor = Color.Navy;
            button2.ForeColor = Color.White;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // if (textBox1.Text != "")
            //{
            //    string columnName = textBox1.Text.Trim();

            //    if (!string.IsNullOrWhiteSpace(columnName))
            //    {
            //        // Create a new column with the provided name
            //        DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
            //        newColumn.HeaderText = columnName;
            //        newColumn.Name = columnName;

            //        // Add the column to the DataGridView control
            //        dataGridView1.Columns.Add(newColumn);
            //    }
            //    else
            //    {
            //        MessageBox.Show("Please enter a valid column name.", "Invalid Column Name", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("يرجى كتابة اسم العمود اولا ");
            //}
        }

        private void ExportToWord2()
        {

            try
            {
                // Create a new file dialog to choose the save location
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Word Document (*.docx)|*.docx";
                saveFileDialog.Title = "Save Word Document";
                saveFileDialog.ShowDialog();

                if (saveFileDialog.FileName != "")
                {
                    string savePath = saveFileDialog.FileName;

                    // Create a new Word document
                    Word.Application wordApp = new Word.Application();
                    Word.Document doc = wordApp.Documents.Add();

                    // Set the layout margin to narrow
                    doc.PageSetup.LeftMargin = wordApp.InchesToPoints(0.5f);
                    doc.PageSetup.RightMargin = wordApp.InchesToPoints(0.5f);
                    doc.PageSetup.TopMargin = wordApp.InchesToPoints(0.5f);
                    doc.PageSetup.BottomMargin = wordApp.InchesToPoints(0.5f);

                    // Get the DataGridView data
                    DataTable dataTable = (DataTable)dataGridView1.DataSource;

                    // Set the title in the header
                    Word.Section section = doc.Sections.First;
                    Word.HeaderFooter header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    Word.Range headerRange = header.Range;

                    // Add the image to the header
                    string imagePath = Properties.Settings.Default.ImageHeaderPath;  // replace with your image path
                    headerRange.InlineShapes.AddPicture(imagePath);

                    // Add the label after the image in the center
                    Word.Paragraph labelParagraph = headerRange.Paragraphs.Add();
                    Word.Range labelRange = labelParagraph.Range;
                    labelRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    labelRange.Text = "رقم اللجنة: " + (int)numericUpDown1.Value;
                    labelRange.Font.Bold = -1;
                    labelRange.Font.Size = 30;
                    labelRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    // Create a Word table
                    Word.Table table = doc.Tables.Add(doc.Range(), dataTable.Rows.Count + 1, dataTable.Columns.Count);

                    // Set table style and borders
                    table.set_Style("Table Grid");
                    table.Borders.Enable = 1; // Enable table borders

                    // Set table headers
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        table.Cell(1, i + 1).Range.Text = dataTable.Columns[i].ColumnName;
                    }
                    // Set the width of the column you want to decrease
                    float targetColumnWidth = 1.2f; // Adjust this value as needed
                    float targetColumnWidthunmber = 0.9f; // Adjust this value as needed
                    float targetColumnWidthName = 2.3f; // Adjust this value as needed
                    float targetColumnWidthdept = 3.2f;

                    // Get the column index of the "رقم الجلوس" column
                    int targetColumnIndex = 3; // Assuming it's the second column, adjust as needed
                    int targetColumnIndexNumber = 1;
                    int targetColumnIndexName = 2;
                    int targetColumnIndexdept = 5;
                    // Set the width of the target column
                    table.Columns[targetColumnIndex].Width = targetColumnWidth * 60; // Multiply by 72 to convert from inches to points
                    table.Columns[targetColumnIndexNumber].Width = targetColumnWidthunmber * 40;
                    table.Columns[targetColumnIndexName].Width = targetColumnWidthName * 70;
                    table.Columns[targetColumnIndexdept].Width = targetColumnWidthdept * 30;

                    //int[] targetColumnIndices = { 1, 2, 3, 4, 5 }; // Specify the indices of the columns you want to fill

                    // Calculate the total number of cells to fill for the specific columns
                    // Calculate the total number of cells to fill
                    int totalCells = dataTable.Rows.Count * dataTable.Columns.Count;

                    // Create a counter variable to track the number of filled cells
                    int filledCells = 0;

                    // Fill table with data for all columns
                    for (int row = 0; row < dataTable.Rows.Count; row++)
                    {
                        for (int col = 0; col < dataTable.Columns.Count; col++)
                        {
                            Word.Cell cell = table.Cell(row + 2, col + 1);
                            cell.Range.Text = dataTable.Rows[row][col].ToString();
                            cell.Range.Font.Bold = -1;

                            // Increment the filled cells counter
                            filledCells++;

                            // Calculate the progress as a percentage
                            double progressPercentage = (double)filledCells / totalCells * 100;

                            // Update a progress label or display the progress percentage in any other way
                            // For example, if you have a label named progressLabel:
                            progressLabel.Text = $"Progress: {progressPercentage:F2}%";

                            // Allow the UI to update by invoking Application.DoEvents()
                            Application.DoEvents();
                        }
                    }
                   table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);

                    //Get te end of the document range
                    Word.Range endRange = doc.Range(doc.Content.End - 1);

                    // Array of signature labels
                    //string[] signatureLabels = { "الملاحظ 1", "الملاحظ 2" };
                    //string ignaturelable2 = "نائب رئيس المركز الاختباري";
                    //string ignaturelable3 = "رئيس المركز الاختباري";

                    // Add labels for signatures
                    //endRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;


                    //endRange.Text += $"{signatureLabels[0]}: ____________________ \t\t\t\t" + $"{signatureLabels[1]}: ____________________ \n";

                    //endRange.Text += $"{ignaturelable2}: ____________________ \t\t" + $"{ignaturelable3}: ____________________ " + "  ";

                    //endRange.Font.Bold = 5;
                    //endRange.Font.Size = 26;

                    // Select all content in the document
                    Word.Range range = doc.Content;
                    range.Select();

                    // Make the selection bold
                    Word.Selection selection = wordApp.Selection;
                    selection.Font.Bold = 1;

                    // Save the Word document
                    doc.SaveAs2(savePath);

                    // Close the Word document and application
                    doc.Close();
                    wordApp.Quit();

                    MessageBox.Show("Data exported to Word successfully!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }



        }
        ///function
        ///
        private void FillDgvByCulumns()
        {
            if (numericUpDown1.Value == 0)
            {
                // Handle the case when the numericUpDown1 value is 0
                // You can clear the DataGridView or handle it as needed
            }
            else if (numericUpDown1.Value > 0)
            {
                dataGridView1.DataSource = null;
                int selectedNumber = (int)numericUpDown1.Value;
                connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={Properties.Settings.Default.FilePath};Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    try
                    {
                        connection.Open();

                        string query = $"SELECT `اسم الطالب`,`رقم الجلوس`,`اسم المادة`,`التخصص`  FROM [{sheetName}$] WHERE اللجنة = ?";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        command.Parameters.AddWithValue("@NumericValue", selectedNumber);

                        OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                        DataTable dataTable = new DataTable();
                        adapter.Fill(dataTable);

                        // Add a column for the serial number to the DataTable
                        dataTable.Columns.Add("م", typeof(string));

                        // Loop through each row in the DataTable and set the serial number
                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            dataTable.Rows[i]["م"] = (i + 1).ToString();
                        }

                        // Reorder the columns in the DataTable to move the SerialNumber column to the first position
                        dataTable.Columns["م"].SetOrdinal(0);

                        dataGridView1.DataSource = dataTable;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"An error occurred: {ex.Message}");
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            FillDgvByCulumns();

            ExportToWord2();


        }

        private void numericUpDown1_KeyUp(object sender, KeyEventArgs e)
        {
 

        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {

        }

        private void numericUpDown1_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            button5.BackColor = Color.White;
            button5.ForeColor = Color.Black;
           
            if (numericUpDown1.Value == 0)
            {
                // Handle the case when the numericUpDown1 value is 0
                // You can clear the DataGridView or handle it as needed
            }
            else if (numericUpDown1.Value > 0)
            {
                    dataGridView1.DataSource = null;
                    int selectedNumber = (int)numericUpDown1.Value;
                    connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={Properties.Settings.Default.FilePath};Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                    using (OleDbConnection connection = new OleDbConnection(connectionString))
                    {
                        try
                        {
                            connection.Open();

                            string query = $"SELECT `اسم الطالب`,`رقم الجلوس`,`اسم المادة`,`التخصص`,`توقيع الحضور`,`توقيع التسليم`  FROM [{sheetName}$] WHERE اللجنة = ?";
                            OleDbCommand command = new OleDbCommand(query, connection);
                            command.Parameters.AddWithValue("@NumericValue", selectedNumber);

                            OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);

                        // Add a column for the serial number to the DataTable
                        dataTable.Columns.Add("م", typeof(string));

                        // Loop through each row in the DataTable and set the serial number
                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            dataTable.Rows[i]["م"] = (i + 1).ToString();
                        }

                        // Reorder the columns in the DataTable to move the SerialNumber column to the first position
                        dataTable.Columns["م"].SetOrdinal(0);

                        dataGridView1.DataSource = dataTable;
                    }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"An error occurred: {ex.Message}");
                        }
                        finally
                        {
                            connection.Close();
                        }
                    }
                

                }
             
            button5.BackColor = Color.Navy;
            button5.ForeColor = Color.White;
        }
         
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.FilePath = null;
            Properties.Settings.Default.Save();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
 
            // Create an instance of the OpenFileDialog
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title = "Select an Excel File";

            // Show the OpenFileDialog and wait for the user's selection
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Get the selected file path
                string filePath = openFileDialog.FileName;

                // Create an Excel application object
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = true;

                // Open the workbook
                Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);

                // Get the first sheet of the workbook
                Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Sheets[1];

                // Rename the sheet to "Sheet1"
                worksheet.Name = "Sheet1";

                // Save and close the workbook
                workbook.Save();
                workbook.Close();

                // Quit Excel application
                excelApp.Quit();
             }
        }
    }

}
