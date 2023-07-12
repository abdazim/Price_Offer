using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; //References-using Excel=Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.IO;

namespace Products3._0
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            timer1.Start(); //timer to run the datetime

            ////////////// automatic excel file(xls.xls) read
            string filePath = string.Empty;
            string fileExt = string.Empty;
            filePath = @"xls.xlsx"; //directory
            //filePath = file.FileName; //get the path of the file  
            newfilepath = filePath;   // get the path from the dialog user path file 
            fileExt = Path.GetExtension(filePath); //get the file extension  
            DataTable dtExcel = new DataTable();
            dtExcel = ReadExcel(filePath, fileExt); //read excel file  
            dataGridView1.Visible = true;
            dataGridView1.DataSource = dtExcel;
        }
        string newfilepath;  // get the path from the dialog user path file 



        /////////////////////////Excel sheet area/////////////////////////////////////////////////////

        /// <summary>
        /// show new file 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (xlApp != null) // check if excel app is installed in platform 
                {
                    if (File.Exists(newfilepath))
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; ++i) //loop for datagridview rows 
                        { //if row [1] empty show message
                            if (String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[0].Value as String)|| String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[1].Value as String)|| String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[2].Value as String)|| String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[3].Value as String)|| String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[4].Value as String) || String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[5].Value as String))
                            {MessageBox.Show("No Data insrted.."); }
                            else // if row 1 is full go to function
                            {
                                ExportToExcelshow();//GO TO  function (show excel file)
                            }
                        }                                              
                    }
                     else { MessageBox.Show("xlsx file not found in the Directory"); }
                }
                else
                {
                    MessageBox.Show("Excel is not properly installed!!"); // else of excel app installed
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Button4-show file button: " + ex.Message);
            }
            finally
            {
                //kill();
            }
        }



        /// <summary>
        /// save file  button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        //initialize the Excel application Object.check excel installed,shoud install excel in platform
        Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (xlApp != null) // check if excel app is installed in platform 
                {
                    if (File.Exists(newfilepath))   //check if the xls.xlsx file exist in the directory
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; ++i) //loop for datagridview rows 
                        { //if row [1] empty show message
                            if (string.IsNullOrEmpty(dataGridView1.Rows[i].Cells[0].Value as string) || string.IsNullOrEmpty(dataGridView1.Rows[i].Cells[1].Value as string) || string.IsNullOrEmpty(dataGridView1.Rows[i].Cells[2].Value as string) || string.IsNullOrEmpty(dataGridView1.Rows[i].Cells[3].Value as string) || string.IsNullOrEmpty(dataGridView1.Rows[i].Cells[4].Value as string) || string.IsNullOrEmpty(dataGridView1.Rows[i].Cells[5].Value as string))
                            {
                                MessageBox.Show("No Data insrted..");
                                
                            }
                            else // if row 1 data filled go to function
                            {
                                saveas(); //save the excel file by the user  
                                return;
                            }
                        }
                    }
                    else
                    {
                         MessageBox.Show("xlsx file not found in the Directory"); 
                    }
                }
                else
                {
                    MessageBox.Show("Excel is not properly installed!!"); // else of excel app installed
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Button3-save file button: " + ex.Message);
            }
            finally { kill();  }
        }


        /// <summary>
        /// clear all Rows
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; ++i) //loop for datagridview rows 
            { //if row [0-5] empty do nothing 
              if (String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[0].Value as String) && String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[1].Value as String) && String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[2].Value as String) && String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[3].Value as String) && String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[4].Value as String) && String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[5].Value as String))
                {
                    dataGridView1.Rows[i].Cells[6].Value = null;//id all row empty-> datagridview(6)=comment = " "                   
                }
                else
                {
                    if (dataGridView1.Rows.Count > 0) //if row >0 (not empty) remove all
                    {
                        dataGridView1.Rows.RemoveAt(i);
                        i--; //loop1:remove last row , loop2:remove before 1 ...
                    }

                    else { dataGridView1.Rows.RemoveAt(i);  }   //if row ==0 remove just this specific row 
                }
                dataGridView1.AllowUserToAddRows = true;  //show one empty row after remove(i)    
            }
            textBox3.Text = "";
            textBox2.Text = "";
            textBox1.Text = "";
        }

        

        ///////////////functions///////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// check func (count)
        /// </summary>
        private void validfunccount(string check)
        {
            int intvalue;
            bool myValue = int.TryParse(check, out intvalue); //intvalue get the original string(num)
            if (myValue == true) //check if the string just number ( no digits no / no , )
            {
                // check if number between 0-1000
                if (intvalue > 0 && intvalue <= 1000)
                {
                    countresult = intvalue;
                    //MessageBox.Show("dvalue " + dValue +",price:"+ price);
                }
                else //if number not between 1-1000 show this message
                {
                    MessageBox.Show("Enter a valid Count");
                }
            }
            else //if price not number show this message 
            {
                MessageBox.Show(" Not Valid, Enter A Valid Count Number.");
            }
        }



        /// <summary>
        /// check func (price)
        /// </summary>
        private void validfuncprice(string check)
        {
            double dValue;   //can use double or int
            bool myValue = double.TryParse(check, out dValue); //dvalue get the original string(num)
            if (myValue == true)
            {
                //TO check if price between 1 and 800.000
                if (dValue > 0 && dValue <= 800000)
                {
                    priceresult = dValue;
                    //MessageBox.Show("dvalue " + dValue +",price:"+ price);
                }
                else //if number not between 1-800.000 show this message
                {
                    MessageBox.Show("Enter a valid Price");
                }
            }
            else //if price not number show this message 
            {
                MessageBox.Show(" Not Valid, Enter A Valid Price .");
            }
        }



        /// <summary>
        ///total calculate
        /// </summary>
        /// 
        int count_row;
        private void total()
        {
            //VAT calculate 
            dec = Convert.ToDouble(textBox2.Text);  //17(double) =17(text)
            dec = dec / 100;    //17 / 100 = 0.17
            //pricecount = price * count 
            pricecount = priceresult * countresult;
            //sum to pricecount
            sum += pricecount;
            totalprice = sum + txt1parse;  //sum+work rates
            //+vat to original (price*count)
            vatcalc = totalprice * dec; //totalprice* 0.17f(textbox2)(dec)=(340vat)
            totalprice += vatcalc;   //add vat to totalprice
            // rows number (products number)
            count_row = dataGridView1.Rows.Count;
        }



        /// <summary>
        /// total price function get textbox1 and textbox2(dec/100) and select cell and do sum 
        /// </summary>
        /// 
        int countresult;  // int64	8 bytes  , int/int32 4 bytes –2,147,483,648 to 2,147,483,647
        double priceresult; //double   8 bytes
        double pricecount;
        double sum;
        double vatcalc;
        double totalprice;
        double txt1parse; //float.Parse(textBox1.Text)
        private void Totalprices()
        {
            try
            {
                txt1parse = double.Parse(textBox1.Text);
                sum = 0;
                priceresult = 0;
                for (int i = 0; i < dataGridView1.Rows.Count; ++i)
                {
                    if (String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[0].Value as String))
                    { MessageBox.Show("Enter Product Name"); return; }
                    if (String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[1].Value as String))
                    { MessageBox.Show("Enter Product Color"); return; }

                    //price//////////////////////////////////////////////////////                                
                    if (String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[2].Value as String))
                    { MessageBox.Show("Enter Product Price"); return; }
                    else   //function to regular exception (justnumbers double)
                    { validfuncprice(Convert.ToString(dataGridView1.Rows[i].Cells[2].Value)); }
                    ////////////////////////////////////////////////////////////

                    //count ////////////////////////////////////////////////////
                    if (String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[3].Value as String))
                    { MessageBox.Show("Enter Product Count"); return; }
                    else   //function to regular exception (justnumbers int)
                    { validfunccount(Convert.ToString(dataGridView1.Rows[i].Cells[3].Value)); }
                    ////////////////////////////////////////////////////////////


                    if (String.IsNullOrEmpty(dataGridView1.Rows[i].Cells[4].Value as String))
                    { MessageBox.Show("Enter Product Description"); return; }

                    dataGridView1.AllowUserToAddRows = false;  /////////////////////// no add rows automaticlly .

                    //go to total function if count and price valid numbers
                    if (countresult > 0 && countresult < 100 && priceresult > 0 && priceresult < 800000)
                    {
                        total();   //total calculate
                        dataGridView1.Rows[0].Cells[5].Value = "Price: " + totalprice.ToString() + " , Products: " + count_row.ToString();
                        dataGridView1.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells; //autofit - autosize column
                    }
                    else { break; }   // break the loop if count not from 0-100 and price not 0 to 800000  (enter product name)                   
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        ////////excel sheet area functions//////////////////

        /// <summary>
        /// method ReadExcel who returns a datatable using the following logic.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="fileExt"></param>
        /// <returns></returns>
        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007
            //HDR is the header, IMEX = 1 is used to retrieve the mixed data from the columns.  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HRD=Yes;IMEX=1';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch { }
            }
            return dtexcel;
        }


        

        /// <summary> 
        /// dataGridView show function
        /// </summary> 
        private void ExportToExcelshow()
        {
            try
            {
                Excel.Application excel = new Excel.Application();
                excel.Visible = true;   //open the excel file 
                Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
                Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[1];
                int StartCol = 1;
                int StartRow = 1;
                int j = 0, i = 0;

                //Write Headers
                for (j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Excel.Range myRange = (Excel.Range)sheet1.Cells[StartRow, StartCol + j];
                    myRange.Value2 = dataGridView1.Columns[j].HeaderText;
                }
                StartRow++;
                //Write datagridview content
                for (i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        try
                        {
                            Excel.Range myRange = (Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];
                            myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                        }
                        catch
                        {
                            ;
                        }
                    }
                }
                sheet1.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }



        /// <summary>
        /// export datagridview to excel file (.xlsx)
        /// </summary>
        private void saveas()
        {
            try
            {
                Excel.Application excel = new Excel.Application();
                excel.Visible = false;   //True =open the excel file 
                Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
                Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[1];
                int StartCol = 1;
                int StartRow = 1;
                int j = 0, i = 0;

                //Write Headers
                for (j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Excel.Range myRange = (Excel.Range)sheet1.Cells[StartRow, StartCol + j];
                    myRange.Value2 = dataGridView1.Columns[j].HeaderText;
                }
                StartRow++;
                //Write datagridview content
                for (i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        try
                        {
                            Excel.Range myRange = (Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];
                            myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                        }
                        catch
                        {
                            ;
                        }
                    }
                }
                sheet1.Columns.AutoFit(); //autofit columns
                //Getting the location and file name of the excel to save from user.
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel Worksheets(*.xlsx)|*.xlsx";
                //saveDialog.SupportMultiDottedExtensions = true;
                //saveDialog.FilterIndex = 2;
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Export Successful");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        //////////////////////price offer area/////////////////////////////////////////////////////////
        /// <summary>
        /// textbox1 - Work
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_TextChanged(object sender, EventArgs e)
        {  
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox1.Text, "[^0-9]"))
            {
                MessageBox.Show("Enter a valid Number 0-9");
                textBox1.Text = textBox1.Text.Remove(textBox1.Text.Length - 1);
            }        
        }



        /// <summary>
        /// textBox2-  VAT (maa'm)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        double dec=0; 
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox2.Text, "[^0-9]"))
            {
                    MessageBox.Show("Enter a valid Number 0-9");
                    textBox2.Text = textBox2.Text.Remove(textBox2.Text.Length - 1);
            }
        }



        /// <summary>
        /// price offer clear textboxes 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
        }


        /// <summary>
        /// total price button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (xlApp != null) // check if excel app is installed in platform 
                {
                    if (File.Exists(newfilepath)) //check if file xlsx exist in directory 
                    {
                        if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))  //regular-expressions
                        {
                            if (System.Text.RegularExpressions.Regex.IsMatch(textBox1.Text, "[0-9]") && System.Text.RegularExpressions.Regex.IsMatch(textBox2.Text, "[0-9]"))
                            {
                                Totalprices(); // function
                            }
                            else
                            {
                                MessageBox.Show("Enter Valid Number");
                            }
                        }
                    }
                    else { MessageBox.Show("xlsx file not found in the Directory"); }
                }
                else
                {
                    MessageBox.Show("Excel is not properly installed!!"); // else of excel app installed
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Total price: " + ex.Message);
            }
        }


        /// <summary>
        /// choose excel file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        //string newfilepath;
        private void button1_Click(object sender, EventArgs e)
        {
            //string filePath = string.Empty;
            //string fileExt = string.Empty;
            ////OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            ////if (file.ShowDialog() == DialogResult.OK) //if there is a file choosen by the user  
            //{
            //    //string file = @"//xls.xls";
            //    filePath = @"..\xls.xls";
            //    //filePath = file.FileName; //get the path of the file  
            //    newfilepath = filePath;   // get the path from the dialog user path file 
            //    fileExt = Path.GetExtension(filePath); //get the file extension  
            //    if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
            //    {
            //        try
            //        {
            //            DataTable dtExcel = new DataTable();
            //            dtExcel = ReadExcel(filePath, fileExt); //read excel file  
            //            dataGridView1.Visible = true;
            //            dataGridView1.DataSource = dtExcel;
            //        }
            //        catch (Exception ex)
            //        {
            //            MessageBox.Show(ex.Message.ToString());
            //        }
            //    }
            //    else
            //    {
            //        MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
            //    }
            //}
        }

        //////////////////////////////////menu strip1///////////////////////////////////

        /// <summary>
        /// Instructions
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void instructionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var Instructions = new Instructions();
            Instructions.Show();
        }



        /// <summary>
        /// about
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var About = new About();
            About.Show();
        }



        /// <summary>
        /// mail to 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mailToToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var mail2 = new mail2();
            mail2.Show();
        }



        //Description area /////////////////////////////////////////////////////
        /// <summary>
        /// clear
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            textBox3.Text = "";
        }


        /// <summary>
        /// add desk
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click(object sender, EventArgs e)
        {
            try
            { 
            if (xlApp != null) // check if excel app is installed in platform 
            {
                if (File.Exists(newfilepath))
                {
                        if (!string.IsNullOrEmpty(textBox3.Text)) //check if textbox3 empty
                        { 
                    dataGridView1.Rows[0].Cells[6].Value = textBox3.Text;
                    dataGridView1.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells; //autofit-autosize column
                        }
                }
                else { MessageBox.Show("xlsx file not found in the Directory"); }
            }
            else
            {
                MessageBox.Show("Excel is not properly installed!!"); // else of excel app installed
            }


        }
            catch (Exception ex)
            {
                MessageBox.Show("Button8-add comment button: " + ex.Message);
            }
}

        ////////////////////////////////////////////////////////

        /// <summary>
        /// kill excel program from task manager
        /// </summary>
        private void kill()
        {
            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }
        }



        /// <summary>
        /// dataGridView1
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }



        /// <summary>
        /// close the app
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
            Application.ExitThread(); //Exits the message loop on the current thread and closes all windows on the thread 
            Application.Exit(); // app exit

        }




        /// <summary>
        ///timer to run the time  
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime currentTime = DateTime.Now;
            label4.Text = "Date: " + currentTime.ToString("dd/MM/yyyy") + " , Time:" + DateTime.Now.ToLongTimeString();
            timer1.Start();  //and start timer in public Form1 ()
        }



        /// <summary>
        /// add new row button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists(newfilepath))
                {
                    dataGridView1.AllowUserToAddRows = true;
                }
                else { MessageBox.Show("xlsx file not found in the Directory"); }
            }
            catch (Exception ex)
            {
                MessageBox.Show("add new row button: " + ex.Message);
            }
        }



        /// <summary>
        /// form load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {


        }


    }
}

