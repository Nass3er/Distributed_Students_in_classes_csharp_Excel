using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace DistributedStudents
{
    public partial class Form1 : Form
    {

        string connectionString;
        string sheetName = "sheet1";
        //Dictionary<int, Button> buttons = new Dictionary<int, Button>(); // Add this line to create a dictionary for buttons
        List<Button> buttons = new List<Button>(); // Create a list for buttons


        public Form1()
        {

            InitializeComponent();
             
          

            createButtonsAndLables();
  
        }

        private void createButtonsAndLables()
        {
            buttons.Clear();
            // Assuming you have a FlowLayoutPanel named "flowLayoutPanel1" on your form:
            flowLayoutPanel1.Controls.Clear();

            string[] classNumbers = Properties.Settings.Default.classNumbers.Split(',');
            string[] classCapacities = Properties.Settings.Default.classCapacities.Split(',');

            flowLayoutPanel1.FlowDirection = FlowDirection.LeftToRight; // Set the flow direction to horizontal



            for (int i = 0; i < classNumbers.Length; i++)
            {
                // Create a container panel for each button and label
                Panel containerPanel = new Panel();
                containerPanel.AutoSize = true; // Adjust panel size to fit contents

                // Create a button for each classCapacity
                Button button = new Button();
                button.Text = classCapacities[i];
                button.Name = "nbutton" + i; // Assign a unique name to each button

                // Set button properties
                button.BackColor = Color.Navy;
                button.ForeColor = Color.White;
                button.Font = new Font("Tahoma", 10);
                button.Width = 63;
                button.Height = 49;

                buttons.Add(button); // Add the button to the dictionary

                // Create a label for each classNumber
                Label label = new Label();
                label.Text = classNumbers[i];
                label.Name = "nlabel" + i; // Assign a unique name to each label

                // Set label properties
                label.Font = new Font("Tahoma", 9);
                label.BackColor = Color.White;
                label.ForeColor = Color.Navy;
                label.Width = 19;
                label.Height = 21;

                // Set the label under the corresponding button
                label.Top = button.Bottom;
                label.Left = button.Left + 22;

                // Add the button and label to the container panel
                containerPanel.Controls.Add(button);
                containerPanel.Controls.Add(label);

                // Add the container panel to the FlowLayoutPanel
                flowLayoutPanel1.Controls.Add(containerPanel);

            }

            // Assign the hover event handler to each button in the list
            foreach (Button button in buttons)
            {
                button.MouseHover += ButtonHover;
            }
        }

        // Assuming you have a list of materials
        List<string> materials = new List<string>();
        private void ButtonHover(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            int buttonIndex = buttons.IndexOf(button);
            int matchingIndex = buttonIndex + 1;

            // Find the matching rows in the DataGridView
            var matchingRows = dataGridView1.Rows.Cast<DataGridViewRow>()
                .Where(row => row.Cells["اللجنة"].Value?.ToString() == matchingIndex.ToString());

            // Get the materials from the matching rows without duplicates
            List<string> matchingMaterials = matchingRows
                .Select(row => row.Cells["اسم المادة"].Value?.ToString())
                .Distinct()
                .ToList();

            // Get the departments from the matching rows without duplicates
            List<string> matchingDepartments = matchingRows
                .Select(row => row.Cells["التخصص"].Value?.ToString())
                .Distinct()
                .ToList();

            // Create a tooltip for materials
            ToolTip materialsTooltip = new ToolTip();
            materialsTooltip.Show(string.Join(", ", matchingMaterials), button, 0, -30, 2000); // Delay the tooltip display by 2000 milliseconds (2 seconds)

            // Create a tooltip for departments
            ToolTip departmentsTooltip = new ToolTip();
            departmentsTooltip.Show(string.Join(", ", matchingDepartments), button, 0, 0, 2000); // Delay the tooltip display by 2000 milliseconds (2 seconds)

            // Hook the MouseLeave event to dispose of the tooltips
            button.MouseLeave += (s, ev) =>
            {
                materialsTooltip.Dispose();
                departmentsTooltip.Dispose();
            };
        }



        private void calculateCapacityOfClassesAndShowInButtonText()
        {
            flowLayoutPanel1.Controls.Clear();

            string[] classNumbers = Properties.Settings.Default.classNumbers.Split(',');
            string[] classCapacities = Properties.Settings.Default.classCapacities.Split(',');

            flowLayoutPanel1.FlowDirection = FlowDirection.LeftToRight; // Set the flow direction to horizontal

            // Calculate the remaining capacity for each class
            int[] remainingCapacities = new int[classCapacities.Length];
            for (int i = 0; i < classCapacities.Length; i++)
            {
                remainingCapacities[i] = int.Parse(classCapacities[i]);
            }

            // Subtract the used capacity from the remaining capacity
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["اللجنة"].Value != null)
                {
                    string committeeCapacity = row.Cells["اللجنة"].Value.ToString();
                    int index = Array.IndexOf(classNumbers, committeeCapacity);
                    if (index >= 0)
                    {
                        remainingCapacities[index] -= 1; // Assuming each committee uses 1 capacity
                    }
                }
            }

            for (int i = 0; i < classNumbers.Length; i++)
            {
                // Create a container panel for each button and label
                Panel containerPanel = new Panel();
                containerPanel.AutoSize = true; // Adjust panel size to fit contents

                // Create a button for each classCapacity
                Button button = new Button();
                button.Text = $"{classCapacities[i]} ({remainingCapacities[i]} )";
                button.Name = "nbutton" + i; // Assign a unique name to each button

                if (remainingCapacities[i] < int.Parse(classCapacities[i]))
                {
                    button.BackColor = Color.Green;
                }
                else if (remainingCapacities[i] == 0)
                {
                    button.BackColor = Color.Red;
                }

                // Set button properties
                button.BackColor = Color.Navy;
                button.ForeColor = Color.White;
                button.Font = new Font("Tahoma", 10);
                button.Width = 66;
                button.Height = 51;

                buttons.Add(button); // Add the button to the dictionary

                // Create a label for each classNumber
                Label label = new Label();
                label.Text = classNumbers[i];
                label.Name = "nlabel" + i; // Assign a unique name to each label

                // Set label properties
                label.Font = new Font("Tahoma", 9);
                label.BackColor = Color.White;
                label.ForeColor = Color.Navy;
                label.Width = 19;
                label.Height = 21;

                // Set the label under the corresponding button
                label.Top = button.Bottom;
                label.Left = button.Left + 22;

                // Add the button and label to the container panel
                containerPanel.Controls.Add(button);
                containerPanel.Controls.Add(label);

                // Add the container panel to the FlowLayoutPanel
                flowLayoutPanel1.Controls.Add(containerPanel);
            }
        }
        private void PopulateDepartmentsComboBox()
        {
            HashSet<string> uniqueDepartments = new HashSet<string>();

            // Iterate through each row in the DataGridView
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // Retrieve the department value from the desired column (e.g., "departments")
                string department = row.Cells["التخصص"].Value?.ToString();

                // Add the department name to the HashSet to eliminate duplicates
                if (!string.IsNullOrEmpty(department))
                {
                    uniqueDepartments.Add(department);
                }
            }

            // Clear the existing items in the ComboBox
            comboboxDepartements.Items.Clear();

            // Add the unique department names to the ComboBox
            comboboxDepartements.Items.AddRange(uniqueDepartments.ToArray());
        }

        private void GetStudentsByDay()
        {
            DateTime selectedDate = dateTimePicker1.Value.Date; // Get the selected date from the DateTimePicker control

            string filePath = Properties.Settings.Default.mainFilePath;
            if (!File.Exists(filePath) || filePath == null)
            {
                //Display the OpenFileDialog to the user
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel Files|*.xlsx";
                openFileDialog.Title = "Select Excel File";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;
                    Properties.Settings.Default.mainFilePath = filePath;
                }
            }
               

            connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR=YES;'";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    string query = $"SELECT * FROM [{sheetName}$] WHERE التاريخ = ?";
                    OleDbCommand command = new OleDbCommand(query, connection);
                    command.Parameters.AddWithValue("@Day", selectedDate.ToString("yyyy/MM/dd")); // Format the selected date as "yyyy/MM/dd"

                    OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);

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
 
        private void ExportToExcel(DataGridView dataGridView)
        {
            try
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial; // Set the license context

                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1"); // Create a worksheet

                    // Export the column headers
                    for (int i = 0; i < dataGridView.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dataGridView.Columns[i].HeaderText;
                    }

                    // Export the data rows
                    for (int i = 0; i < dataGridView.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 2, j + 1].Value = dataGridView.Rows[i].Cells[j].Value?.ToString();
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
                        package.SaveAs(new FileInfo(outputPath));

                        MessageBox.Show("Data exported to Excel successfully!");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }

        }
        private void appendDataToFinalFile()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial; // or LicenseContext.NonCommercial

            // Get the JSON string from settings
            string json = Properties.Settings.Default.finalReport;

            // Deserialize the JSON string back to a List<DepartmentInfo>
            List<DepartmentInfo> departmentList = JsonConvert.DeserializeObject<List<DepartmentInfo>>(json);

            // Specify the path of the existing Excel file
            string filePath = Properties.Settings.Default.finalFilePath;

            // Check if the file exists
            if (!File.Exists(filePath))
            {
                MessageBox.Show("The file does not exist or has not been added in the system variable configuration screen. Please add the final file path!");
                return;
            }

            // Load the existing Excel file using EPPlus
            FileInfo fileInfo = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet;

                // Check if the workbook contains any worksheets
                if (package.Workbook.Worksheets.Count == 0)
                {
                    // Add a new worksheet to the workbook
                    worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    // Write the headers
                    worksheet.Cells[1, 1].Value = "التاريخ";
                    worksheet.Cells[1, 2].Value = "اليوم";
                    worksheet.Cells[1, 3].Value = "م";
                    worksheet.Cells[1, 4].Value = "التخصصات";
                    worksheet.Cells[1, 5].Value = "المادة";
                    worksheet.Cells[1, 6].Value = "عدد الطلاب";
                    worksheet.Cells[1, 7].Value = "عدد الحضور";
                    worksheet.Cells[1, 8].Value = "عدد الغياب";
                    worksheet.Cells[1, 9].Value = "غش";
                    worksheet.Cells[1, 10].Value = "ملاحظات";

                    // Apply formatting to the header row if needed
                    // For example, to make it bold:
                    using (ExcelRange headerRange = worksheet.Cells[1, 1, 1, 10])
                    {
                        headerRange.Style.Font.Bold = true;
                    }
                }
                else
                {
                    // Get the first worksheet in the Excel file
                    worksheet = package.Workbook.Worksheets[0];

                    // Check if the worksheet is empty
                    if (worksheet.Dimension == null)
                    {
                        // Write the headers
                        worksheet.Cells[1, 1].Value = "التاريخ";
                        worksheet.Cells[1, 2].Value = "اليوم";
                        worksheet.Cells[1, 3].Value = "م";
                        worksheet.Cells[1, 4].Value = "التخصصات";
                        worksheet.Cells[1, 5].Value = "المادة";
                        worksheet.Cells[1, 6].Value = "عدد الطلاب";
                        worksheet.Cells[1, 7].Value = "عدد الحضور";
                        worksheet.Cells[1, 8].Value = "عدد الغياب";
                        worksheet.Cells[1, 9].Value = "غش";
                        worksheet.Cells[1, 10].Value = "ملاحظات";

                        // Apply formatting to the header row if needed
                        // For example, to make it bold:
                        using (ExcelRange headerRange = worksheet.Cells[1, 1, 1, 10])
                        {
                            headerRange.Style.Font.Bold = true;
                        }
                    }
                }

                // Find the last used row in the worksheet
                int lastUsedRow = worksheet.Dimension?.End.Row ?? 0;

                // Append the data from the departmentList to the worksheet
                int rowIndex = lastUsedRow + 1;

                foreach (DepartmentInfo department in departmentList)
                {
                    worksheet.Cells[rowIndex, 1].Value = department.التاريخ;
                    worksheet.Cells[rowIndex, 1].Style.Numberformat.Format = "dd/MM/yyyy"; // or any other format you prefer
                    worksheet.Cells[rowIndex, 2].Value = department.اليوم;
                    worksheet.Cells[rowIndex, 3].Value = department.م;
                    worksheet.Cells[rowIndex, 4].Value = department.التخصصات;
                    worksheet.Cells[rowIndex, 5].Value = department.المادة;
                    worksheet.Cells[rowIndex, 6].Value = department.عدد_الطلاب;
                    worksheet.Cells[rowIndex, 7].Value = department.عدد_الحضور;
                    worksheet.Cells[rowIndex, 8].Value = department.عدد_الغياب;
                    worksheet.Cells[rowIndex, 9].Value = department.غش;
                    worksheet.Cells[rowIndex, 10].Value = department.ملاحظات;

                    // Set values for other columns as needed
                    rowIndex++;
                }

                // Save the changes to the Excel file
                package.Save();
            }

        }
        private void Form1_Load(object sender, EventArgs e)
        {
         
                Properties.Settings.Default.classNumbers = "1,2,3,4,5,6,7,8,9,10";
                Properties.Settings.Default.classCapacities = "48,46,47,49,42,40,34,38,45,43";
                Properties.Settings.Default.Save();
            
            createButtonsAndLables();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            GetStudentsByDay();
            addnessesaryColumns();
            int rowcount = dataGridView1.RowCount;
            MessageBox.Show($"Number of rows: {rowcount}");

            PopulateDepartmentsComboBox();

            //populate the ComboBox with the numbers from the settings:
            string numbers = Properties.Settings.Default.classNumbers;
            string[] numberArray = numbers.Split(',');

            foreach (string number in numberArray)
            {
                comboBoxClasses.Items.Add(number);
                comboBoxclasses2.Items.Add(number);
            }
            createButtonsAndLables();
        }
        private void addnessesaryColumns()
        {
            List<string> targetColumnNames = new List<string> { "اللجنة", "توقيع الحضور", "توقيع التسليم" };

            foreach (string columnName in targetColumnNames)
            {
                if (!string.IsNullOrWhiteSpace(columnName))
                {
                    // Create a new column with the provided name
                    DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                    newColumn.HeaderText = columnName;
                    newColumn.Name = columnName;

                    // Add the column to the DataGridView control
                    dataGridView1.Columns.Add(newColumn);
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
             
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button2.BackColor = Color.White;
            button2.ForeColor = Color.Black;

            ExportToExcel(dataGridView1);

            button2.BackColor = Color.Navy;
            button2.ForeColor = Color.White;
        }
 
        private void button3_Click(object sender, EventArgs e)
        {



        }

        private void button4_Click(object sender, EventArgs e)
        {

        }
   
        private void button5_Click(object sender, EventArgs e)
        {
            btnEdit.BackColor = Color.White;
            btnEdit.ForeColor = Color.Black;

            int selectedIndex = comboBoxclasses2.SelectedIndex;
            string capacities = Properties.Settings.Default.classCapacities;
            string[] capacityArray = capacities.Split(',');
            int capacity = 0;
            if (selectedIndex >= 0 && selectedIndex < capacityArray.Length)
            {
                if (int.TryParse(classCapacity.Text, out capacity))
                {
                    capacityArray[selectedIndex] = capacity.ToString();
                    Properties.Settings.Default.classCapacities = string.Join(",", capacityArray);
                    Properties.Settings.Default.Save();
                    numberofclass.Text = "";
                    classCapacity.Text = "";
                }
                else
                {
                    MessageBox.Show("سعة اللجنة يجب ان تكون ارقام ");
                }
            }

            btnEdit.BackColor = Color.Navy;
            btnEdit.ForeColor = Color.White;
            label10.Visible = false;
            label11.Visible = false;
            btnAdd.Visible = false;
            btnEdit.Visible = false;
            classCapacity.Visible = false;
            numberofclass.Visible = false;
            createButtonsAndLables();
            
        }
         
        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
             
            string sheetName2 = "sheet1";
            //Display the OpenFileDialog to the user
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx";
            openFileDialog.Title = "Select Excel File";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {

                string filePath = openFileDialog.FileName;

                // string filePath = "D:/Nasser/students.xlsx";

                //DateTime selectedDate = DateTime.Now; // Replace this with your selected date

                connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    try
                    {
                        connection.Open();

                        string query = $"SELECT * FROM [{sheetName2}$] ";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        //command.Parameters.AddWithValue("@Day", selectedDate.ToString("yyyy/MM/dd")); // Format the selected date as "yyyy/MM/dd"

                        OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                        DataTable dataTable = new DataTable();
                        adapter.Fill(dataTable);
                         
                        dataGridView1.DataSource = dataTable;
                        calculateCapacityOfClassesAndShowInButtonText(); // calculate the capacity of classes and show in buttonsText
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

        private void button4_Click_1(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                dataGridView1.AllowUserToAddRows = true;
            }
            else
            {
                dataGridView1.AllowUserToAddRows = false;
            }
        }
         
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void button4_Click_2(object sender, EventArgs e)
        {
            button4.BackColor = Color.White;
            button4.ForeColor = Color.Black;

            string selectedDepartment = comboboxDepartements.SelectedItem?.ToString();
            var selectedItem = comboBoxClasses.SelectedItem;
            int numberOfStudents = (int)numericUpDown1.Value;

            if (string.IsNullOrEmpty(selectedDepartment))
            {
                MessageBox.Show("Please select a department.");
                return;
            }
            if (selectedItem == null)
            {
                MessageBox.Show("Please select a valid class.");
                return;
            }

            int selectedClassNumber;
            if (!int.TryParse(selectedItem.ToString(), out selectedClassNumber))
            {
                MessageBox.Show("Invalid class selected.");
                return;
            }

            string materialColumnName = "اسم المادة"; // Replace with the actual column name for "material"
            string departmentColumnName = "التخصص"; // Replace with the actual column name for "dept"
            string selectedMaterial = ""; // Define the selectedMaterial variable

                    // Find the rows in the DataGridView that match the selected department
                    var matchingRows = dataGridView1.Rows.Cast<DataGridViewRow>()
            .Where(row => row.Cells[departmentColumnName].Value?.ToString() == selectedDepartment &&
                           string.IsNullOrEmpty(row.Cells["اللجنة"].Value?.ToString()))
            .Take(numberOfStudents);

            foreach (DataGridViewRow row in matchingRows)
            {
                if (row.Cells[materialColumnName].Value != null)
                {
                    selectedMaterial = row.Cells[materialColumnName].Value.ToString();
                    break; // Assuming there is only one selected material for the selected department
                }
            }

            Button button = buttons[selectedClassNumber - 1];

            string capacities = Properties.Settings.Default.classCapacities;
            string[] capacityArray = capacities.Split(',');
            int classCapacity = 0;

            if (selectedClassNumber >= 1 && selectedClassNumber <= capacityArray.Length)
            {
                if (!int.TryParse(capacityArray[selectedClassNumber - 1], out classCapacity))
                {
                    MessageBox.Show("Invalid class selected.");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Invalid class selected.");
                return;
            }

            int targetNumber = int.Parse(comboBoxClasses.SelectedItem.ToString()); // Get the selected number from the ComboBox
            int duplicateCount = 0; // Variable to store the count of duplicates

            // Iterate through each row in the DataGridView
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                //if (row.Cells[materialColumnName].Value != null && row.Cells[departmentColumnName].Value != null)
                //{
                //    string material = row.Cells[materialColumnName].Value.ToString();
                //    string department = row.Cells[departmentColumnName].Value.ToString();

                //    // Check if the material and department match the target values
                //    if (material == selectedMaterial && department != selectedDepartment)
                //    {
                //        MessageBox.Show("The same material cannot be added for different departments in the same class.");
                //        return;
                //    }
                //}

                // Get the value in the specified column ("اللجنة")
                if (row.Cells["اللجنة"].Value != null)
                {
                    int value;
                    if (int.TryParse(row.Cells["اللجنة"].Value.ToString(), out value))
                    {
                        // Check if the value matches the target number
                        if (value == targetNumber)
                        {
                            duplicateCount++;
                        }
                    }

                    if (duplicateCount == classCapacity)
                    {
                        MessageBox.Show("اللجنة ممتلئة.");
                        return; // Stop distributing students once the required number is reached
                    }
                }
            }

            int availableCapacity = classCapacity - duplicateCount;

            // Check if the number of students to assign exceeds the available capacity
            if (numberOfStudents > availableCapacity)
            {
                MessageBox.Show("عدد الطلاب يتجاوز السعة المتبقية في اللجنة " + selectedItem + "، السعة المتبقية في اللجنة هي " + availableCapacity + " طالب فقط، سيتم توزيع " + availableCapacity + " طالب فقط من العدد الذي أدخلته.");

                // Update numberOfStudents to the available capacity
                numberOfStudents = availableCapacity;
            }

            // Distribute the students among the matching rows
            int rowIndex = duplicateCount;

            foreach (DataGridViewRow row in matchingRows)
            {
                // Check if the "اللجنة" column is already assigned a class number
                if (row.Cells["اللجنة"].Value == null)
                {
                    // Check if the required number of students has been assigned
                    if (rowIndex >= classCapacity)
                    {
                        MessageBox.Show("تجاوزت سعة اللجنة " + rowIndex);
                        return; // Stop distributing students once the required number is reached
                    }
                   
                    // Set the value of the "اللجنة" column to the selected class number
                    row.Cells["اللجنة"].Value = selectedClassNumber;

                    //if (row.Cells["اللجنة"].Value != null)
                    //{
                    //    lastModifiedCell = row.Cells["اللجنة"];
                    //    lastModifiedCell.Value = assignedValue; // Modify the cell value
                    //}
                
                    if (selectedClassNumber - 1 < buttons.Count)
                    {
                        button.Text = $"({++rowIndex}/{classCapacity})";
                        
                        // Check the capacity before assigning students
                        if (rowIndex >= classCapacity)
                        {
                            button.BackColor = Color.Red; // Change the color as needed
                        }
                        else if (rowIndex < classCapacity && rowIndex != 0)
                        {
                            button.BackColor = Color.Green; // Change the color as needed
                        }

                        button.Refresh(); // Force an immediate update of the button's appearance
                    }
                }
                else
                {
                    // The "اللجنة" column already has a class number assigned, skip this row
                    rowIndex++;
                }
            }

            //string selectedDepartment = comboboxDepartements.SelectedItem?.ToString();
            //var selectedItem = comboBoxClasses.SelectedItem;
            //int numberOfStudents = (int)numericUpDown1.Value;

            //if (string.IsNullOrEmpty(selectedDepartment))
            //{
            //    MessageBox.Show("Please select a department.");
            //    return;
            //}
            //if (selectedItem == null)
            //{
            //    MessageBox.Show("Please select a valid class.");
            //    return;
            //}
            //int selectedClassNumber;
            //if (int.TryParse(selectedItem.ToString(), out selectedClassNumber))
            //{
            //    // Find the rows in the DataGridView that match the selected department
            //    var matchingRows = dataGridView1.Rows.Cast<DataGridViewRow>()
            //        .Where(row => row.Cells["التخصص"].Value?.ToString() == selectedDepartment).Take(numberOfStudents);

            //    Button button = buttons[selectedClassNumber - 1];

            //    string capacities = Properties.Settings.Default.classCapacities;
            //    string[] capacityArray = capacities.Split(',');
            //    int classCapacity = 0;
            //    if (selectedClassNumber >= 1 && selectedClassNumber <= capacityArray.Length)
            //    {
            //        if (int.TryParse(capacityArray[selectedClassNumber - 1], out classCapacity))
            //        {
            //            // Use the retrieved class capacity
            //        }
            //        else
            //        {
            //            MessageBox.Show("Invalid class selected.");
            //            return;
            //        }
            //    }
            //    else
            //    {
            //        MessageBox.Show("Invalid class selected.");
            //        return;
            //    }

            //    int targetNumber = int.Parse(comboBoxClasses.SelectedItem.ToString()); // Get the selected number from the ComboBox
            //    int duplicateCount = 0; // Variable to store the count of duplicates

            //    // Iterate through each row in the DataGridView
            //    foreach (DataGridViewRow row in dataGridView1.Rows)
            //    {


            //        // Get the value in the specified column ("اللجنة")
            //        if (row.Cells["اللجنة"].Value != null)
            //        {
            //            int value;
            //            if (int.TryParse(row.Cells["اللجنة"].Value.ToString(), out value))
            //            {
            //                // Check if the value matches the target number
            //                if (value == targetNumber)
            //                {
            //                    duplicateCount++;
            //                }
            //            }

            //            if (duplicateCount == classCapacity)
            //            {
            //                MessageBox.Show("اللجنة ممتلئة " );
            //                break; // Stop distributing students once the required number is reached
            //            }
            //        }
            //    }

            //    int availableCapacity = classCapacity - duplicateCount;

            //    // Check if the number of students to assign exceeds the available capacity
            //    if (numberOfStudents > availableCapacity)
            //    {
            //        MessageBox.Show("عدد الطلاب اكبر من السعة المتبقيه في اللجنة "+ selectedItem + ",السعة المتبقية في اللجنة هو  " + availableCapacity + "طالب فقط ," + "سيتم توزيع " + availableCapacity + "طلاب فقط من العدد الذي أدخلته");

            //    }
            //    // Distribute the students among the matching rows
            //    int rowIndex = duplicateCount;

            //    foreach (DataGridViewRow row in matchingRows)
            //    {    
            //        // Check if the "اللجنة" column is already assigned a class number
            //        if (row.Cells["اللجنة"].Value == null)
            //        {
            //            // Check if the required number of students has been assigned
            //            if (rowIndex >= classCapacity)
            //            {
            //                //MessageBox.Show("تجاوزت سعة اللجنة" + rowIndex);
            //                break; // Stop distributing students once the required number is reached
            //            }

            //            // Set the value of the "اللجنة" column to the selected class number
            //            row.Cells["اللجنة"].Value = selectedClassNumber;

            //            if (selectedClassNumber - 1 < buttons.Count)
            //            {   button.Text = $"({++rowIndex}/{classCapacity})";

            //                // Check the capacity before assigning students
            //                if (rowIndex >= classCapacity)
            //                {
            //                    button.BackColor = Color.Red; // Change the color as needed
            //                }
            //                else if (rowIndex < classCapacity && rowIndex != 0)
            //                {
            //                    button.BackColor = Color.Green; // Change the color as needed
            //                }

            //                button.Refresh(); // Force an immediate update of the button's appearance
            //            } 
            //        }

            //    }

            //    if (rowIndex < numberOfStudents)
            //    {
            //        MessageBox.Show("Not enough unassigned students. Only " + rowIndex + " students were assigned.");
            //    }
            //}

            //else
            //{
            //    MessageBox.Show("Invalid class selected.");
            //    return;
            //}



            button4.BackColor = Color.Navy;
            button4.ForeColor = Color.White;

            
            comboboxDepartements.Text = "";
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void comboBoxClasses_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void comboboxDepartements_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void classCapacity_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void numberofclass_TextChanged(object sender, EventArgs e)
        {

        }
        public class DepartmentInfo
        {
            public DateTime التاريخ { get; set; }
            public string اليوم { get; set; }
            public string م { get; set; } = "0";

            public string التخصصات { get; set; }
            public string المادة { get; set; }
            public int عدد_الطلاب { get; set; }
            public int عدد_الحضور { get; set; }
            public int عدد_الغياب { get; set; }
            public int غش { get; set; }
            public string ملاحظات { get; set; }


        }
        private void button7_Click(object sender, EventArgs e)
        {
            button7.BackColor = Color.White;
            button7.ForeColor = Color.Black;
            dateTimePicker2.Visible = true;  // Show the DateTimePicker control
            dateTimePicker2.Focus();         // Set focus to the DateTimePicker control

            button7.BackColor = Color.Navy;
            button7.ForeColor = Color.White;


        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            int i = 0;
            List<DepartmentInfo> departmentList = new List<DepartmentInfo>();
            HashSet<string> uniqueDepartments = new HashSet<string>();

            foreach (DataGridViewRow deptRow in dataGridView1.Rows)
            {
                string department = deptRow.Cells["التخصص"].Value?.ToString();
                string material = deptRow.Cells["اسم المادة"].Value?.ToString();
                DateTime selectedDate = dateTimePicker1.Value;
                string day = selectedDate.ToString("dddd");  // Extract the day from the selected date

                if (!string.IsNullOrEmpty(department) && !string.IsNullOrEmpty(material))
                {
                    string departmentMaterialKey = $"{department}-{material}";

                    if (!uniqueDepartments.Contains(departmentMaterialKey))
                    {
                        int numberOfStudents = 0;

                        foreach (DataGridViewRow studentRow in dataGridView1.Rows)
                        {
                            string studentDept = studentRow.Cells["التخصص"].Value?.ToString();
                            string studentMaterial = studentRow.Cells["اسم المادة"].Value?.ToString();

                            if (department == studentDept && material == studentMaterial)
                            {
                                numberOfStudents++;
                            }
                        }
                        i++;
                        DepartmentInfo departmentInfo = new DepartmentInfo
                        {
                            التخصصات = department,
                            المادة = material,
                            عدد_الطلاب = numberOfStudents,
                            التاريخ = selectedDate,
                            اليوم = day,
                            م = i.ToString(),
                        };


                        departmentList.Add(departmentInfo);
                        uniqueDepartments.Add(departmentMaterialKey);
                    }
                }


            }



            // Serialize the list to JSON
            string json = JsonConvert.SerializeObject(departmentList);

            // Save the JSON string to settings
            Properties.Settings.Default.finalReport = json;
            Properties.Settings.Default.Save();
            appendDataToFinalFile();
            MessageBox.Show("تم ترحيل البيانات اللازمة إلى جدول الكشف النهائي ");
            // Hide the DateTimePicker control after the date is selected
            dateTimePicker2.Visible = false;

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void comboBoxclasses2_SelectedIndexChanged(object sender, EventArgs e)
        {
            label10.Visible = true;
            label11.Visible = true;
            btnAdd.Visible = true;
            btnEdit.Visible = true;
            classCapacity.Visible = true;
            numberofclass.Visible = true;

            int selectedIndex = comboBoxclasses2.SelectedIndex;
            string numbers = Properties.Settings.Default.classNumbers;

            string capacities = Properties.Settings.Default.classCapacities;
            string[] numberArray = numbers.Split(',');
            string[] capacityArray = capacities.Split(',');

            if (selectedIndex >= 0 && selectedIndex < capacityArray.Length)
            {
                numberofclass.Text = numberArray[selectedIndex];
                classCapacity.Text = capacityArray[selectedIndex];
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            // Read the new number from the textbox
            string newNumberText = numberofclass.Text;

            // Validate and parse the new number
            if (!int.TryParse(newNumberText, out int newNumber))
            {
                MessageBox.Show("القيمة المدخله غير صحيحة");
                return;
            }

            // Retrieve the existing numbers from settings
            string numbers = Properties.Settings.Default.classNumbers;

            // Check if the new number already exists
            if (numbers.Split(',').Contains(newNumber.ToString()))
            {
                MessageBox.Show("الرقم مضاف مسبقا");
                return;
            }

            // Read the new capacity from the textbox
            string newCapacityText = classCapacity.Text;

            // Validate and parse the new capacity
            if (!int.TryParse(newCapacityText, out int newCapacity))
            {
                MessageBox.Show("Invalid capacity entered.");
                return;
            }

            // Append the new number and capacity to the existing values
            numbers += $",{newNumber}";
            string capacities = Properties.Settings.Default.classCapacities + $",{newCapacity}";

            // Update the settings with the modified values
            Properties.Settings.Default.classNumbers = numbers;
            Properties.Settings.Default.classCapacities = capacities;
            Properties.Settings.Default.Save();
            numberofclass.Text = "";
            classCapacity.Text = "";
            btnEdit.BackColor = Color.Navy;
            btnEdit.ForeColor = Color.White;
            label10.Visible = false;
            label11.Visible = false;
            btnAdd.Visible = false;
            btnEdit.Visible = false;
            classCapacity.Visible = false;
            numberofclass.Visible = false;
            createButtonsAndLables();
        }

        private void button3_Click_2(object sender, EventArgs e)
        {

        }

        private void button3_Click_3(object sender, EventArgs e)
        {
            
        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void label2_Click_1(object sender, EventArgs e)
        {

        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
        private Stack<DataGridViewCell> modifiedCells = null; // Declare the stack at the class level
        DataGridViewCell lastModifiedCell = null; // Declare a variable to store the last modified cell

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (modifiedCells.Count == 0)
            {
                MessageBox.Show("No assignments to undo.");
                return;
            }

            // Prompt the user if they want to undo the last assignment
            DialogResult result = MessageBox.Show("Do you want to undo the last assignment?", "Undo", MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {
                DataGridViewCell lastModifiedCell = modifiedCells.Pop();
                lastModifiedCell.Value = null; // Reset the value of the last modified cell

                MessageBox.Show("Assignment undone.");
            }

        }
    
    }
}
    
