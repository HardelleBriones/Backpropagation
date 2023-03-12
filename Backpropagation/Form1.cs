using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Backprop;

namespace Backpropagation
{
    public partial class Form1 : Form
    {
        NeuralNet nn;
        DataTable dataTable = new DataTable();
        public Form1()
        {
            InitializeComponent();
            openFileDialog1 = new OpenFileDialog();
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.InitialDirectory = "C:\\";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx";

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Naa sa forms duha kabuok ang input neuron tapos usa ang output (katong textbox)
            nn = new NeuralNet(13, 10, 2);
        }

        private void button2_Click(object sender, EventArgs e)
        {
           for(int x = 0; x < Convert.ToInt32(textBox4.Text); x++) 
            {
                //row
                for (int i = 1; i < dataTable.Rows.Count; i++)
                {
                    //column
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {

                        //input all columns except the last column
                        if (j  < dataTable.Columns.Count-1)
                        {
                            nn.setInputs(j, Convert.ToDouble(dataTable.Rows[i][j].ToString()));
                        }
                        else 
                        {
                            nn.setDesiredOutput(0, Convert.ToDouble(dataTable.Rows[i][j].ToString()));
                            nn.learn();
                        }
                    }
                }
           
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            nn.setInputs(0, Convert.ToDouble(textBox1.Text));
            nn.setInputs(1, Convert.ToDouble(textBox2.Text));
            nn.setInputs(2, Convert.ToDouble(textBox5.Text));
            nn.setInputs(3, Convert.ToDouble(textBox6.Text));
            nn.setInputs(4, Convert.ToDouble(textBox7.Text));
            nn.setInputs(5, Convert.ToDouble(textBox8.Text));
            nn.setInputs(6, Convert.ToDouble(textBox9.Text));
            nn.setInputs(7, Convert.ToDouble(textBox10.Text));
            nn.setInputs(8, Convert.ToDouble(textBox11.Text));
            nn.setInputs(9, Convert.ToDouble(textBox12.Text));
            nn.setInputs(10, Convert.ToDouble(textBox13.Text));
            nn.setInputs(11, Convert.ToDouble(textBox14.Text));
            nn.setInputs(12, Convert.ToDouble(textBox15.Text));
            nn.run();
            textBox3.Text = "" + nn.getOuputData(0);

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (DialogResult.OK == openFileDialog1.ShowDialog())
            {
                string filePath = openFileDialog1.FileName;
                // Specify the sheet name
                string sheetName = "Sheet1";

                // Specify the connection string
                string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties=\"Excel 12.0;HDR=YES;\"";





            try
            {
                // Open a connection to the Excel file using the connection string
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    // Open the connection
                    connection.Open();

                    // Create a command to select data from the specified sheet
                    OleDbCommand command = new OleDbCommand($"SELECT * FROM [{sheetName}$]", connection);

                    // Execute the command and fill the DataTable with the data
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }

                    // Close the connection
                    connection.Close();
                }
            }
            catch (System.InvalidOperationException ex) 
            {
                MessageBox.Show(ex.ToString());
            }
           
            // Ask the user which columns to include in a popup dialog
            DialogResult dialogResult = DialogResult.Cancel;
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                DataColumn column = dataTable.Columns[i];

                dialogResult = MessageBox.Show($"Would you like to include the column {column.ColumnName}?", "Include Column?", MessageBoxButtons.YesNoCancel);

                if (dialogResult == DialogResult.No)
                {
                    // Remove the column from the DataTable
                    dataTable.Columns.Remove(column);
                    i--;
                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    // Exit the loop if the user clicks cancel
                    break;
                }
            }

            }
            dataGridView1.DataSource = dataTable;
            dataGridView1.Dock = DockStyle.None;
            dataGridView1.Width = 440;
            dataGridView1.Height = 300;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }
    }
}
