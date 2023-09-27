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

namespace DistributedStudents
{
    public partial class InitialSystem : Form
    {
        public InitialSystem()
        {
            InitializeComponent();
            //Properties.Settings.Default.finalFilePath =  null;
            //Properties.Settings.Default.Save();
            if (Properties.Settings.Default.ImageHeaderPath != "")
            {
                pictureBox1.Image = Image.FromFile(Properties.Settings.Default.ImageHeaderPath);
            }
        }

        private void btnUpdateClasses_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Image Files (*.jpg;*.png;*.gif;*.bmp)|*.jpg;*.png;*.gif;*.bmp";
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string imagePath = openFileDialog.FileName;

                    // Save the image path to the properties or settings
                    Properties.Settings.Default.ImageHeaderPath = imagePath;
                    Properties.Settings.Default.Save();

                    // Display the selected image or perform any other operations
                    pictureBox1.Image = Image.FromFile(imagePath);
                }
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Check if the finalFilePath setting is empty
            // Prompt the user to choose the file path
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial; // or LicenseContext.NonCommercial

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xlsx";
            saveFileDialog.Title = "Save Excel File";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;

                // Check if the file already exists
                if (!File.Exists(filePath))
                {
                    // Create the file
                    using (ExcelPackage package = new ExcelPackage())
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1"); // Create a worksheet

                        

                        // Save the workbook to the file path
                        package.SaveAs(new FileInfo(filePath));
                    }
                }

                // Assign the selected file path to the setting
                Properties.Settings.Default.finalFilePath = filePath;
                Properties.Settings.Default.Save();

                MessageBox.Show("تم حفظ مسار الملف النهائي، لن تحتاج إلى تحديده مرة أخرى.");
            }

        }
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx";
            openFileDialog.Title = "Select Excel File";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
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

                // Save the file path to the application's properties or settings
                Properties.Settings.Default.mainFilePath = filePath;
                Properties.Settings.Default.Save();

                MessageBox.Show("The main file path has been saved, and the sheet has been renamed to 'Sheet1'. You don't need to specify it again.");
            }
        }

        private void InitialSystem_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.mainFilePath = null;
            Properties.Settings.Default.finalFilePath = null;
            Properties.Settings.Default.ImageHeaderPath = null;

        }
    }
}
