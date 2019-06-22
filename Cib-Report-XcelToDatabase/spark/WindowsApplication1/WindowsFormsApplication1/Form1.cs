using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        string dir = System.Configuration.ConfigurationManager.AppSettings["dirp"];
        string bac = System.Configuration.ConfigurationManager.AppSettings["bac"];
        string location = @"";
        int strat =40;


         //   u.Segments.Last().Split('.')[0];


        public Form1()
        {
            InitializeComponent();



           
        }
        //for sheet clean up
       

        private void button2_Click(object sender, EventArgs e)
        {

            doing();

        }


        public void doing()
    {



        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range range;
        int track = 0;
        string str;
        int rCnt;
       // int cCnt;
        int rw = 0;
        int cl = 0;
        int redirect = 0;

        int flag = 0;
        String temp = "";


        xlApp = new Excel.Application();
        xlWorkBook = xlApp.Workbooks.Open(location, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);



        range = xlWorkSheet.UsedRange;
        rw = range.Rows.Count;
        cl = range.Columns.Count;

        this.progressBar1.Value = (strat / rw) * 100;   
        bool xx = true;

        progressBar1.Value = 10;
        
       
            for (rCnt = strat; rCnt <= rw; rCnt++)
            {
                //this.label1.Text = rCnt.ToString();

                if ((String)(range.Cells[rCnt, 1] as Excel.Range).Text.ToString() == "NOTES:")
                {


                    xx = false;

                    break;



                }





                if (((String)(range.Cells[rCnt, 1] as Excel.Range).Text.ToString()).Length >= 12)
                {
                    str = ((String)(range.Cells[rCnt, 1] as Excel.Range).Text.ToString()).Substring(0, 12);


                    if (str == "CONFIDENTIAL")
                    {

                        if (flag == 0) {
                            
                           
                            if((String)(range.Cells[rCnt-1, 1] as Excel.Range).Text.ToString()=="")
                            
                            { 
                            temp = (rCnt - 1).ToString(); flag++; 
                     
                            
                            
                            }


                            else
                            {

                                temp = (rCnt).ToString(); flag++; 
                                redirect = 1;


                            }
                        
                        
                        
                        
                        
                        }
                        else
                        {




                            String send = "";
                            
                            
                            flag = 0;
                            if (redirect == 0) { send = "A" + temp + ":" + "A" + rCnt; } else { send = "A" + temp + ":" + "A" + (rCnt+1); }
                           
                          //  MessageBox.Show("" + send);
                            clean(send);
                            strat = rCnt;


                            label2.Text ="Stat IS :"+strat+ "size :"+rw+ " deleted " + send;
                            break;
                   
                        
                        
                        
                        
                        }



                        //MessageBox.Show("" + rCnt);

                    }

                }

                if (rCnt >= rw) { xx = false; }



            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            if (xx == false)
            {
               // MessageBox.Show("Clean Done");

                strat = 40;

            }
            else
            {


                doing();
            }

    }

        public void upload ()
        {
            if (location == "") {

                MessageBox.Show("Choose File Frist");
            }
            else { 

            button1.Enabled = false;
            label2.Text = "Starting Process";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            int track = 0;
            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

        



            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(location, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);



            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            progressBar1.Value = 10;
            for (rCnt = 10; rCnt <= rw; rCnt++)
            {
                for (cCnt = 3; cCnt <= cl; cCnt++)
                {
                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();
                   //  MessageBox.Show("" + str );

                    if (str == "Type of subject:")
                    {
                        track = 1;
                        label2.Text = "Type of subject FOUND";



                        //  MessageBox.Show(""+(cCnt+1)+ "," + rCnt);

                        if ((String)(range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Value2 == "INDIVIDUAL")
                        {
                            label2.ForeColor = Color.Green;
                         label2.Text = "TYPE IS INDIVIDUAL";
                    //   MessageBox.Show("INDIVIDUAL");
                     Individual();
                       


                           this.progressBar1.Value = 100;
                           
                            
                          // MessageBox.Show(""+location);

                            break;
                        }

                        else if ((String)(range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Value2 == "COMPANY")
                        {
                            label2.ForeColor = Color.Blue;
                            label2.Text = "TYPE IS COMPANY";

                  //       MessageBox.Show("COMPANY");



                             company();
                            progressBar1.Value = 100;

                          //  MessageBox.Show("upload Compleate");

                            break;
                        }

                        else
                        {


                            MessageBox.Show("The Statement Type not found ");

                        }


                        break;
                    }

                    if (track == 1) { break; }
                }


                if (track == 1) { break; }


            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            button1.Enabled = true;


            progressBar1.Value = 100;
           // MessageBox.Show(location + "is compleate");


            }
        }





//indi


        private void Individual()
        {

            int master_id;
            label2.Text = "Detecting CIB subject code";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(location, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            //master table values 
            progressBar1.Value = 30;
            string CibCode = "";
            string UserId = "";
            string DateOfInq = "";
            string FiCode = "";
            string BranchCode = "";
            string FiName = "";







            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();
                    // MessageBox.Show("" + str );

                    if (str == "CIB subject code:")
                    {



                        CibCode = (String)(range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Value2;

                        //   MessageBox.Show("" + CibCode);

                        label2.Text = "CIB SUBJECT CODE FOUND ";
                        break;

                    }
                }

                if (CibCode != "") { break; }


            }
            progressBar1.Value = 40;
            //Master Table Start



            for (rCnt = 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "User ID")
                    {



                        UserId = (String)(range.Cells[(rCnt + 1), cCnt] as Excel.Range).Text;
                        DateOfInq = (String)(range.Cells[(rCnt + 1), (cCnt - 1)] as Excel.Range).Text;
                        FiCode = (String)(range.Cells[(rCnt + 1), (cCnt + 1)] as Excel.Range).Text;
                        BranchCode = (String)(range.Cells[(rCnt + 1), (cCnt + 2)] as Excel.Range).Text;
                        FiName = (String)(range.Cells[(rCnt + 1), (cCnt + 3)] as Excel.Range).Text;
                        


                        break;

                    }
                }

                if (UserId != "") { break; }
            }
            progressBar1.Value = 50;


            var context = new CIBEntities();

            var t = new IMaster //Make sure you have a table called test in DB
            {

                CIB_subject_code = CibCode,
                Date_of_Inquiry = Convert.ToDateTime(DateOfInq),
                User_ID = UserId,
                FI_Code = FiCode,
                Branch_Code = BranchCode,
                FI_Name = FiName,
                file_location = location,
                Upload_date = DateTime.Now.ToString("yyyy-MM-ddThh:mm:ssTZD"),
            };

            context.IMasters.Add(t);
            context.SaveChanges();

            label2.Text = "master table hasbeen Uplod Compleate ";
            label2.ForeColor = Color.Red;

            master_id = t.cib_bb_id;

            //Master Table End
            progressBar1.Value = 60;

            //Inquary table starts

            int inque = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();
                //    MessageBox.Show("sub in " + (String)(range.Cells[(rCnt + 1), cCnt] as Excel.Range).Text);
                    if (str == "SUBJECT INFORMATION")
                    {
                        inque = 1;
                        break;
                    
                    }
                    if (str == "Trade name" || str == "Name")
                    {

                        inque = 1;
                        var context23 = new CIBEntities();
                       //    MessageBox.Show("In Inquary " + (String)(range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Text);


                        //   MessageBox.Show("DOB" + (range.Cells[(rCnt + 3), (cCnt + 1)] as Excel.Range).Text);


                           if (str == "Trade name")
                           {

                        var t23 = new I_INQUIRED //Make sure you have a table called test in DB
                        {
                            Trade_name = (String)(range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Text,

                            Proprietorship_District = (String)(range.Cells[(rCnt + 1), (cCnt + 1)] as Excel.Range).Text,
                            Proprietorship_Address = (String)(range.Cells[(rCnt + 2), (cCnt + 1)] as Excel.Range).Text,
                            Owner_Name = (String)(range.Cells[(rCnt + 3), (cCnt + 1)] as Excel.Range).Text,
                            Father_name = (String)(range.Cells[(rCnt + 4), (cCnt + 1)] as Excel.Range).Text,
                            Mother_name = (String)(range.Cells[(rCnt + 5), (cCnt + 1)] as Excel.Range).Text,
                          

                             

                            DOB = (String)(range.Cells[(rCnt + 6), (cCnt + 1)] as Excel.Range).Text,


                            Proprietorship_Postalcode = (String)(range.Cells[(rCnt + 1), (cCnt + 3)] as Excel.Range).Text,
                            NID = (String)(range.Cells[(rCnt + 2), (cCnt + 3)] as Excel.Range).Text,
                            Owner_Address = (String)(range.Cells[(rCnt + 3), (cCnt + 3)] as Excel.Range).Text,

                            Postcode = (String)(range.Cells[(rCnt + 4), (cCnt + 3)] as Excel.Range).Text,
                            District = (String)(range.Cells[(rCnt + 5), (cCnt + 3)] as Excel.Range).Text,
                            TIN = (String)(range.Cells[(rCnt + 6), (cCnt + 3)] as Excel.Range).Text,



                            cib_bb_id = master_id,

                        };
                        context23.I_INQUIRED.Add(t23);
                        context23.SaveChanges();

                        label2.Text = "Inquary table hasbeen Uplod Compleate ";
                        label2.ForeColor = Color.Gold;



                        context23.Dispose();
                    }

                    else{

                        var t23 = new I_INQUIRED //Make sure you have a table called test in DB
                        {
                            Trade_name = (String)(range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Text,

                            Father_name = (String)(range.Cells[(rCnt + 1), (cCnt + 1)] as Excel.Range).Text,
                            Mother_name = (String)(range.Cells[(rCnt + 2), (cCnt + 1)] as Excel.Range).Text,
                            

                              DOB = (String)(range.Cells[(rCnt + 3), (cCnt + 1)] as Excel.Range).Text,

                              NID  = (String)(range.Cells[(rCnt + 4), (cCnt + 1)] as Excel.Range).Text,
                           
                           
                             Owner_Address  = (String)(range.Cells[rCnt, (cCnt + 3)] as Excel.Range).Text,
                             Postcode = (String)(range.Cells[(rCnt + 1), (cCnt + 3)] as Excel.Range).Text,
                             District = (String)(range.Cells[(rCnt + 2), (cCnt + 3)] as Excel.Range).Text,
                             TIN = (String)(range.Cells[(rCnt + 3), (cCnt + 3)] as Excel.Range).Text,

                           
                          
                            



                            cib_bb_id = master_id,

                        };



                        context23.I_INQUIRED.Add(t23);
                        context23.SaveChanges();

                        label2.Text = "Inquary table hasbeen Uplod Compleate ";
                        label2.ForeColor = Color.Gold;



                        context23.Dispose();
                    }








                        break;

                    }
                }

                if (inque != 0) { break; }
            }



            label2.Text = "Inquary table hasbeen Uplod Compleate ";
            label2.ForeColor = Color.Green;


            progressBar1.Value = 70;

            //Inquary table finbish


            // SUBJECT INFORMATION starts

            int xx = 0;

            for (rCnt = rCnt+0 ; rCnt <= rw; rCnt++)
            {
               
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();
                    

                    if (str == "Title, Name:")
                    {
                        xx = 1;

                       

                        DateTime loadedDate = DateTime.ParseExact((String)(range.Cells[(rCnt + 4), (cCnt + 1)] as Excel.Range).Text, "dd/MM/yyyy", null);

                      //  MessageBox.Show("sub in " + (String)(range.Cells[(rCnt + 2), (cCnt + 3)] as Excel.Range).Text);
                        var context2 = new CIBEntities();

                        var t2 = new Sub__INFO //Make sure you have a table called test in DB
                        {

                            CIB_subject_code = CibCode,
                            Title_Name = (String)(range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Text,
                            Fathername = (String)(range.Cells[(rCnt + 1), (cCnt + 1)] as Excel.Range).Text,
                            SpouseName = (String)(range.Cells[(rCnt + 3), (cCnt + 1)] as Excel.Range).Text,
                            Mothername = (String)(range.Cells[(rCnt + 2), (cCnt + 1)] as Excel.Range).Text,
                            Dob = loadedDate,
                            Gender = (String)(range.Cells[(rCnt + 5), (cCnt + 1)] as Excel.Range).Text,
                            District_Country = (String)(range.Cells[(rCnt + 6), (cCnt + 1)] as Excel.Range).Text,
                            NID = (String)(range.Cells[(rCnt + 7), (cCnt + 1)] as Excel.Range).Text,
                            TIN = (String)(range.Cells[(rCnt + 8), (cCnt + 1)] as Excel.Range).Text,


                            Type_of_sub = (String)(range.Cells[(rCnt - 1), (cCnt + 3)] as Excel.Range).Text,
                            Ref_number = (String)(range.Cells[(rCnt), (cCnt + 3)] as Excel.Range).Text,
                            Sector_type = (String)(range.Cells[(rCnt + 1), (cCnt + 3)] as Excel.Range).Text,

                            ID_type = (String)(range.Cells[(rCnt + 3), (cCnt + 3)] as Excel.Range).Text,
                            ID_number = (String)(range.Cells[(rCnt + 4), (cCnt + 3)] as Excel.Range).Text,
                            ID_issue_date = (String)(range.Cells[(rCnt + 5), (cCnt + 3)] as Excel.Range).Text,

                            sector_code = (String)(range.Cells[(rCnt + 2), (cCnt + 3)] as Excel.Range).Text,


                            ID_issue_country = (String)(range.Cells[(rCnt + 6), (cCnt + 3)] as Excel.Range).Text,
                            Telephone = (String)(range.Cells[(rCnt + 7), (cCnt + 3)] as Excel.Range).Text,
                            Remarks = (String)(range.Cells[(rCnt + 8), (cCnt + 3)] as Excel.Range).Text,
                            cib_bb_id = master_id,
                        };

                        context2.Sub__INFO.Add(t2);
                        context2.SaveChanges();

                        label2.Text = "Subject INFO table hasbeen Uplod Compleate ";
                        label2.ForeColor = Color.Black;

                      //  MessageBox.Show("Done" );



                        break;

                    }
                }

                if (xx != 0) { break; }
            }


            progressBar1.Value = 80;



            // SUBJECT INFORMATION Ends
            //Adress table starts










            int xxad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "Address Type")
                    {
                        xxad = 1;

                       //  MessageBox.Show("Address in : " + (String)(range.Cells[(rCnt + 1), cCnt] as Excel.Range).Text);

                        int z = 1;

                        while (true)
                        {


                            if ((String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text != "")
                            {
                                var context23 = new CIBEntities();
                                //      MessageBox.Show("" + (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text);
                                var t23 = new I_ADDRESS //Make sure you have a table called test in DB
                                {
                                    Address_Type = (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text,
                                    Address = (String)(range.Cells[(rCnt + z), (cCnt + 1)] as Excel.Range).Text,
                                    Postal_code = (String)(range.Cells[(rCnt + z), (cCnt + 2)] as Excel.Range).Text,
                                    District = (String)(range.Cells[(rCnt + z), (cCnt + 3)] as Excel.Range).Text,
                                    Country = (String)(range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text,
                                    cib_bb_id = master_id,
                                    flag=0,
                                };

                                context23.I_ADDRESS.Add(t23);
                                context23.SaveChanges();

                                label2.Text = "Adress INFO table hasbeen Uplod Compleate ";
                                label2.ForeColor = Color.Gold;


                                z++;
                                context23.Dispose();
                            }
                            else
                            {

                                break;

                            }

                        }




                        break;

                    }
                }

                if (xxad != 0) { break; }
            }

            int tt = rCnt;
            //Adress table finish

            //company start









            xxad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= 1; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();

                    //    MessageBox.Show("" + (String)(range.Cells[(rCnt), cCnt] as Excel.Range).Text);
                    if (str == "1.(B) SUMMARY OF THE NON-FUNDED FACILITIES AS BORROWER & CO-BORROWER")
                    {

                        xxad = 1;
                        break;
                    } if (str == "COMPANY(S) LIST")
                    {
                        xxad = 1;

                     //   MessageBox.Show("In com[any list  " + (String)(range.Cells[(rCnt + 1), cCnt] as Excel.Range).Text);

                        int z = 3;

                        while (true)
                        {


                            if ((String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text != "")
                            {
                                var context23 = new CIBEntities();
                                //      MessageBox.Show("" + (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text);
                                var t23 = new Com_i //Make sure you have a table called test in DB
                                {
                                    cib_s_c = (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text,
                                    name_of_owner = (String)(range.Cells[(rCnt + z), (cCnt + 1)] as Excel.Range).Text,
                                    role = (String)(range.Cells[(rCnt + z), (cCnt + 2)] as Excel.Range).Text,
                                    fi = (String)(range.Cells[(rCnt + z), (cCnt + 3)] as Excel.Range).Text,
                                    legal = (String)(range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text,
                                    cib_bb_id = master_id,
                                    // flag = 0,
                                };

                                context23.Com_i.Add(t23);
                                context23.SaveChanges();

                                label2.Text = "Adress INFO table hasbeen Uplod Compleate ";
                                label2.ForeColor = Color.Gold;


                                z++;
                                context23.Dispose();
                            }
                            else
                            {

                                break;

                            }

                        }




                        break;

                    }
                }

                if (xxad != 0) { break; }
            }




            //compny ends


            rCnt = tt;


            //linked profit start

            int link = 0;

            for (rCnt = tt; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();

                    if (str == "1.(B) SUMMARY OF THE NON-FUNDED FACILITIES AS BORROWER & CO-BORROWER")
                    {

                        link = 1;
                        break;
                    }
                    if (str == "Sector type:")
                    {

                        link = 1;
                        var context23 = new CIBEntities();
                        //MessageBox.Show("" + (String)(range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Text);
                        var t23 = new PROP_CONCERN //Make sure you have a table called test in DB
                        {
                            CIb_sub_Code = (String)(range.Cells[(rCnt - 1), (cCnt + 1)] as Excel.Range).Text,
                            Sector_type = (String)(range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Text,

                            Sector_code = (String)(range.Cells[(rCnt), (cCnt + 3)] as Excel.Range).Text,
                            Ref_number = (String)(range.Cells[(rCnt - 1), (cCnt + 3)] as Excel.Range).Text,
                            Trade_Name = (String)(range.Cells[rCnt - 1, (cCnt + 5)] as Excel.Range).Text,
                            Tele_number = (String)(range.Cells[rCnt, (cCnt + 5)] as Excel.Range).Text,



                            cib_bb_id = master_id,

                        };

                        context23.PROP_CONCERN.Add(t23);
                        context23.SaveChanges();



                        context23.Dispose();




                        break;

                    }
                }

                if (link != 0) { break; }
            }



            label2.Text = "PROPRIETORSHIP CONCERN table hasbeen Uplod Compleate ";
            label2.ForeColor = Color.Red;

            //linked profit end
            progressBar1.Value = 90;
            //linked address start


            rCnt = tt;

            int xxlad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();

                    if (str == "1.(B) SUMMARY OF THE NON-FUNDED FACILITIES AS BORROWER & CO-BORROWER")
                    {

                        link = 1;
                        break;
                    }
                    if (str == "Address Type")
                    {
                        xxlad = 1;

                        //  MessageBox.Show("" + (String)(range.Cells[(rCnt + 1), cCnt] as Excel.Range).Text);

                        int z = 1;

                        while (true)
                        {


                            if ((String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text != "")
                            {
                                var context23 = new CIBEntities();
                                // MessageBox.Show("" + (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text);
                                var t23 = new I_ADDRESS //Make sure you have a table called test in DB
                                {
                                    Address_Type = (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text,
                                    Address = (String)(range.Cells[(rCnt + z), (cCnt + 1)] as Excel.Range).Text,
                                    Postal_code = (String)(range.Cells[(rCnt + z), (cCnt + 2)] as Excel.Range).Text,
                                    District = (String)(range.Cells[(rCnt + z), (cCnt + 3)] as Excel.Range).Text,
                                    Country = (String)(range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text,
                                    cib_bb_id = master_id,
                                    flag = 1,
                                };

                                context23.I_ADDRESS.Add(t23);
                                context23.SaveChanges();

                                label2.Text = " Linked Adress INFO table hasbeen Uplod Compleate ";
                                label2.ForeColor = Color.Gold;


                                z++;
                                context23.Dispose();
                            }
                            else
                            {

                                break;

                            }

                        }




                        break;

                    }
                }

                if (xxlad != 0) { break; }
            }



            //linked address end
            //I_SUM_OF_FACILITY_S_AS_BOR Starts


            rCnt = tt;

            int xxb = 0;

            for (rCnt = tt; rCnt <= rw; rCnt++)
            {


                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "No of reporting Institutes:")
                    {
                        xxb = 1;

                        //    MessageBox.Show("" + (String)(range.Cells[(rCnt + 1), cCnt] as Excel.Range).Text);





                        var context23 = new CIBEntities();
                        // MessageBox.Show("" + (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text);
                        var t23 = new SUM_OF_FACILITY_S_AS_BOR //Make sure you have a table called test in DB
                        {
                            No_of_reporting_Institutes = Convert.ToInt32((range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Text),
                            No_of_Living_Contracts = Convert.ToInt32((range.Cells[(rCnt + 1), (cCnt + 1)] as Excel.Range).Text),

                            Total_Outstanding_Amount =Convert.ToDecimal((range.Cells[(rCnt + 2), (cCnt + 1)] as Excel.Range).Text),
                            Total_Overdue_Amount = Convert.ToDecimal((range.Cells[rCnt, (cCnt + 3)] as Excel.Range).Text),
                            No_of_Stay_order_contracts = Convert.ToInt32((range.Cells[(rCnt + 1), (cCnt + 3)] as Excel.Range).Text),
                            Total_Outstanding_amount_for_Stay = Convert.ToDecimal((range.Cells[(rCnt + 2), (cCnt + 3)] as Excel.Range).Text),

                            flag = 0,

                            cib_bb_id = master_id,

                        };

                        context23.SUM_OF_FACILITY_S_AS_BOR.Add(t23);
                        context23.SaveChanges();

                        label2.Text = " I_SUM_OF_FACILITY_S_AS_BOR INFO table hasbeen Uplod Compleate ";
                        label2.ForeColor = Color.Gold;

                        context23.Dispose();





                        break;

                    }
                }

                if (xxb != 0) { break; }
            }



            //I_SUM_OF_FACILITY_S_AS_BOR ends
            //1a







            xxlad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "1.(A) SUMMARY OF THE FUNDED FACILITIES AS BORROWER & CO-BORROWER")
                    {
                        int z = 6;
                        xxlad = 1;
                        //  MessageBox.Show("in");
                        //  



                        while (true)
                        {


                            if ((String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text != "")
                            {
                                var context23 = new CIBEntities();
                                //  MessageBox.Show("" + (String)(range.Cells[(rCnt + z), (cCnt + 16)] as Excel.Range).Text);
                                var t23 = new SUM_OF_FUNDED_FACILI_AS_BOR //Make sure you have a table called test in DB
                                {
                                    Contract_Category = (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text,
                                    UC_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 1)] as Excel.Range).Text),
                                    UC_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 2)] as Excel.Range).Text),

                                    SMA_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 3)] as Excel.Range).Text),
                                    SMA_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text),

                                    SS_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 5)] as Excel.Range).Text),
                                    SS_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 6)] as Excel.Range).Text),

                                    DF_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 7)] as Excel.Range).Text),
                                    DF_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 8)] as Excel.Range).Text),


                                    B_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 9)] as Excel.Range).Text),
                                    BL_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 10)] as Excel.Range).Text),


                                    BLW_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 11)] as Excel.Range).Text),
                                    BLW_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 12)] as Excel.Range).Text),

                                    Terminated_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 13)] as Excel.Range).Text),
                                    Terminated_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 14)] as Excel.Range).Text),

                                    Requested_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 15)] as Excel.Range).Text),
                                    Requested_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 16)] as Excel.Range).Text),

                                    Stay_Order_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 17)] as Excel.Range).Text),
                                    Stay_Order_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 18)] as Excel.Range).Text),
                                    

                                    cib_bb_id = master_id,
                                    flag = 0,
                                };

                                context23.SUM_OF_FUNDED_FACILI_AS_BOR.Add(t23);
                                context23.SaveChanges();

                                label2.Text = " I_1a_SUM_OF_FUNDED_FACILI_AS_BOR table hasbeen Uplod Compleate ";
                                label2.ForeColor = Color.Blue;


                                z++;
                                context23.Dispose();
                            }
                            else
                            {

                                break;

                            }

                        }




                        break;

                    }
                }

                if (xxlad != 0) { break; }
            }



            //1a end
            //1b






            xxlad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "1.(B) SUMMARY OF THE NON-FUNDED FACILITIES AS BORROWER & CO-BORROWER")
                    {
                        int z = 6;
                        xxlad = 1;
                        //  MessageBox.Show("in");
                        //  


                        while (true)
                        {


                            if ((String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text != "")
                            {
                                var context23 = new CIBEntities();
                                // MessageBox.Show("" + (String)(range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text);
                                var t23 = new SUM_OF_NON_FUNDED_FACILI_AS_BOR //Make sure you have a table called test in DB
                                {
                                    Type_of_Financing = (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text,

                                    Living_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 1)] as Excel.Range).Text),
                                    Living_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 2)] as Excel.Range).Text),

                                    Terminated_NO =  Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 3)] as Excel.Range).Text),
                                    Terminated_Amount =Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text),

                                    Requested_NO =  Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 5)] as Excel.Range).Text),
                                    Requested_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 6)] as Excel.Range).Text),

                                    Stay_Order_NO =  Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 7)] as Excel.Range).Text),
                                    Stay_Order_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 8)] as Excel.Range).Text),

                                    flag = 0,

                                    cib_bb_id = master_id,



                                };

                                context23.SUM_OF_NON_FUNDED_FACILI_AS_BOR.Add(t23);
                                context23.SaveChanges();

                                label2.Text = " I__1b_SUM_OF_NON_FUNDED_FACILI_AS_BOR table hasbeen Uplod Compleate ";
                                label2.ForeColor = Color.Green;


                                z++;
                                context23.Dispose();
                            }
                            else
                            {

                                break;

                            }

                        }




                        break;

                    }
                }

                if (xxlad != 0) { break; }
            }






            //1b end



            //2




            xxb = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {


                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "No of reporting Institutes:")
                    {
                        xxb = 1;

                        //    MessageBox.Show("" + (String)(range.Cells[(rCnt + 1), cCnt] as Excel.Range).Text);





                        var context23 = new CIBEntities();
                        // MessageBox.Show("" + (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text);
                        var t23 = new SUM_OF_FACILITY_S_AS_BOR //Make sure you have a table called test in DB
                        {
                            No_of_reporting_Institutes =  Convert.ToInt32((range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Text),
                            No_of_Living_Contracts = Convert.ToInt32((range.Cells[(rCnt + 1), (cCnt + 1)] as Excel.Range).Text),

                            Total_Outstanding_Amount = Convert.ToDecimal((range.Cells[(rCnt + 2), (cCnt + 1)] as Excel.Range).Text),
                            Total_Overdue_Amount = Convert.ToDecimal((range.Cells[rCnt, (cCnt + 3)] as Excel.Range).Text),
                            No_of_Stay_order_contracts =  Convert.ToInt32((range.Cells[(rCnt + 1), (cCnt + 3)] as Excel.Range).Text),
                            Total_Outstanding_amount_for_Stay = Convert.ToDecimal((range.Cells[(rCnt + 2), (cCnt + 3)] as Excel.Range).Text),
                            flag = 1,

                            cib_bb_id = master_id,

                        };

                        context23.SUM_OF_FACILITY_S_AS_BOR.Add(t23);
                        context23.SaveChanges();

                        label2.Text = " I_2_SUM_OF_FACILITY_S_AS_GUARANTOR1 INFO table hasbeen Uplod Compleate ";
                        label2.ForeColor = Color.Red;

                        context23.Dispose();





                        break;

                    }
                }

                if (xxb != 0) { break; }
            }


            //2 end

            //2a







            xxlad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "2.(A) SUMMARY OF THE FUNDED FACILITIES AS GUARANTOR")
                    {
                        int z = 6;
                        xxlad = 1;
                        //  MessageBox.Show("in");
                        //  



                        while (true)
                        {


                            if ((String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text != "")
                            {
                                var context23 = new CIBEntities();
                                //  MessageBox.Show("" + (String)(range.Cells[(rCnt + z), (cCnt + 16)] as Excel.Range).Text);
                                var t23 = new SUM_OF_FUNDED_FACILI_AS_BOR //Make sure you have a table called test in DB
                                {
                                    Contract_Category = (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text,


                                    BLW_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 12)] as Excel.Range).Text),
                                    BL_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 10)] as Excel.Range).Text),
                                    Terminated_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 14)] as Excel.Range).Text),
                                    DF_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 8)] as Excel.Range).Text),
                                    SS_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 6)] as Excel.Range).Text),
                                    Requested_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 16)] as Excel.Range).Text),
                                    SMA_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text),
                                    Stay_Order_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 18)] as Excel.Range).Text),
                                    UC_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 2)] as Excel.Range).Text),


                                    UC_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 1)] as Excel.Range).Text),
                                    Stay_Order_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 17)] as Excel.Range).Text),
                                    Requested_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 15)] as Excel.Range).Text),
                                    SMA_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 3)] as Excel.Range).Text),
                                    Terminated_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 13)] as Excel.Range).Text),
                                    BLW_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 11)] as Excel.Range).Text),
                                    SS_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 5)] as Excel.Range).Text),
                                    DF_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 7)] as Excel.Range).Text),
                                    B_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 9)] as Excel.Range).Text),
                               

                                   
                                    



                                    cib_bb_id = master_id,
                                    flag = 1,
                                };

                                context23.SUM_OF_FUNDED_FACILI_AS_BOR.Add(t23);
                                context23.SaveChanges();

                                label2.Text = " I_2a_SUM_OF_FUNDED_FACILI_AS_GUARANTOR1 table hasbeen Uplod Compleate ";
                                label2.ForeColor = Color.Yellow;


                                z++;
                                context23.Dispose();
                            }
                            else
                            {

                                break;

                            }

                        }




                        break;

                    }
                }

                if (xxlad != 0) { break; }
            }



            //2a end




            //2b






            xxlad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "2.(B) SUMMARY OF THE NON-FUNDED FACILITIES AS GUARANTOR")
                    {
                        int z = 6;
                        xxlad = 1;

                        //   MessageBox.Show("in");



                        while (true)
                        {


                            if ((String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text != "")
                            {
                                var context23 = new CIBEntities();
                                // MessageBox.Show("" + (String)(range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text);
                                var t23 = new SUM_OF_NON_FUNDED_FACILI_AS_BOR //Make sure you have a table called test in DB
                                {
                                    Type_of_Financing = (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text,

                                    Living_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 1)] as Excel.Range).Text),
                                    Living_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 2)] as Excel.Range).Text),

                                    Terminated_NO =  Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 3)] as Excel.Range).Text),
                                    Terminated_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text),

                                    Requested_NO =  Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 5)] as Excel.Range).Text),
                                    Requested_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 6)] as Excel.Range).Text),

                                    Stay_Order_NO =  Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 7)] as Excel.Range).Text),
                                    Stay_Order_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 8)] as Excel.Range).Text),

                                    flag = 1,

                                    cib_bb_id = master_id,



                                };

                                context23.SUM_OF_NON_FUNDED_FACILI_AS_BOR.Add(t23);
                                context23.SaveChanges();

                                label2.Text = " I_2b_SUM_OF_FACILITY_S_AS_GUARANTOR1 table hasbeen Uplod Compleate ";
                                label2.ForeColor = Color.Black;


                                z++;
                                context23.Dispose();
                            }
                            else
                            {

                                break;

                            }

                        }




                        break;

                    }
                }

                if (xxlad != 0) { break; }
            }



            //2b end





            //req starts






            xxlad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "REQUESTED CONTRACT DETAILS")
                    {
                        int z = 3;
                        xxlad = 1;
                        // MessageBox.Show("in");
                        //  



                        while (true)
                        {


                            if ((String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text != "")
                            {

                                var context23 = new CIBEntities();
                                // MessageBox.Show("" + (String)(range.Cells[(rCnt + 1), (cCnt +1)] as Excel.Range).Text);
                                var t23 = new REQUESTED_CONTRACT //Make sure you have a table called test in DB
                                {
                                    SL =Convert.ToInt32((range.Cells[(rCnt + z), cCnt] as Excel.Range).Text),

                                    Type_of_Contract = (String)(range.Cells[(rCnt + z), (cCnt + 1)] as Excel.Range).Text,
                                    Facility = (String)(range.Cells[(rCnt + z), (cCnt + 2)] as Excel.Range).Text,

                                    Phase = (String)(range.Cells[(rCnt + z), (cCnt + 3)] as Excel.Range).Text,
                                    Role = (String)(range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text,
                                    FI_Code = (String)(range.Cells[(rCnt + z), (cCnt + 5)] as Excel.Range).Text,
                                    Branch_Code = (String)(range.Cells[(rCnt + z), (cCnt + 6)] as Excel.Range).Text,
                                    Request_date = (String)(range.Cells[(rCnt + z), (cCnt + 7)] as Excel.Range).Text,

                                    Total_Requested_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 8)] as Excel.Range).Text),
                                    CIB_subject_code = (String)(range.Cells[(rCnt + z), (cCnt + 9)] as Excel.Range).Text,
                                    CIB_contract_code = (String)(range.Cells[(rCnt + z), (cCnt + 10)] as Excel.Range).Text,
                                    FI_0contract_codede = (String)(range.Cells[(rCnt + z), (cCnt + 11)] as Excel.Range).Text,




                                    cib_bb_id = master_id,



                                };

                                context23.REQUESTED_CONTRACT.Add(t23);
                                context23.SaveChanges();

                                label2.Text = " I_2b_SUM_OF_FACILITY_S_AS_GUARANTOR1 table hasbeen Uplod Compleate ";
                                label2.ForeColor = Color.Black;

                                z++;
                                context23.Dispose();

                            }
                            else
                            {

                                break;

                            }

                        }




                        break;

                    }
                }

                if (xxlad != 0) { break; }
            }


            //req ends



    //detail table starts


            int row_check = 2;


            xxlad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "DETAILS OF INSTALLMENT FACILITY(S)")
                    {
                        int z = 2;
                        xxlad = 1;

                        row_check = rCnt;

                        // MessageBox.Show("in");

                        string str_new = "";






                        while (str_new != "NOTES:")
                        {

                            //  MessageBox.Show("new loop");
                            string v_Ref = "";
                            string v_FI_code = "";
                            string v_Branch_code = "";
                            string v_CIB_contract_code = "";
                            string v_FI_contract_code = "";

                            string v_Role = "";
                            string v_Phase = "";
                            string v_Facility = "";
                            string v_Starting_date = "";
                            string v_End_date_of_contract = "";
                            string v_Sanction_Limit = "";
                            string v_Total_Disbursement_Amount = "";
                            string v_Total_number_of_installments = "";
                            string v_Installment_Amount = "";
                            string v_Remaining_installments_Number = "";
                            string v_Security_Amount = "";
                            string v_Third_Party_guarantee_Amount = "";
                            string v_Security_Type = "";



                            string v_Date_of_Last_Update = "";
                            string v_Date_of_Law_suit = "";
                            string v_Date_of_Last_payment = "";
                            string v_Date_of_classification = "";
                            string v_Date_of_last_rescheduling = "";
                            string v_Method_of_payment = "";
                            string v_Payments_periodicity = "";
                            string v_Number_of_time_rescheduled = "";
                            string v_Remaining_installments_Amount = "";
                            string v_Reorganized_credit = "";
                            string v_Basis_for_classification_qualitative_judgment = "";
                            string v_Remarks = "";





                            for (int xxx = row_check; rCnt <= rw; xxx++)
                            {


                                str_new = (String)(range.Cells[xxx, 1] as Excel.Range).Text.ToString();
                             //   MessageBox.Show("IN detail : ",str_new);
                                if (str_new == "NOTES:")
                                {
                                     


                                    break;



                                }
                                if (str_new == "Ref")
                                {
                                    v_Ref = (String)(range.Cells[xxx + 1, 1] as Excel.Range).Text.ToString();
                                    v_FI_code = (String)(range.Cells[xxx + 1, 2] as Excel.Range).Text.ToString();
                                    v_Branch_code = (String)(range.Cells[xxx + 1, 3] as Excel.Range).Text.ToString();
                                    v_CIB_contract_code = (String)(range.Cells[xxx + 1, 4] as Excel.Range).Text.ToString();
                                    v_FI_contract_code = (String)(range.Cells[xxx + 1, 5] as Excel.Range).Text.ToString();

                                    xxx = xxx + 1;
                                    //   MessageBox.Show(v_Ref);
                                    //   MessageBox.Show(v_FI_code);
                                    //   MessageBox.Show(v_Branch_code);
                                    //   MessageBox.Show(v_FI_contract_code);
                                    ////  row_check = xxx;
                                    //   break;



                                }

                                else if (str_new == "Role:")
                                {

                                    // MessageBox.Show("in role");
                                    v_Role = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Date of Last Update:")
                                    {
                                        v_Date_of_Last_Update = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Date_of_Last_Update = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //   MessageBox.Show(v_Role);
                                    //  MessageBox.Show(v_Date_of_Last_Update);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }
                                else if (str_new == "Phase:")
                                {
                                    //  MessageBox.Show("in phase");
                                    v_Phase = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Date of Law suit:")
                                    {
                                        v_Date_of_Law_suit = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Date_of_Law_suit = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //   MessageBox.Show(v_Phase);
                                    //   MessageBox.Show(v_Date_of_Last_Update);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }


                                else if (str_new == "Facility:")
                                {

                                    //  / MessageBox.Show("in facility");
                                    v_Facility = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Date of last payment:")
                                    {
                                        v_Date_of_Last_payment = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Date_of_Last_payment = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //   MessageBox.Show(v_Facility);
                                    //   MessageBox.Show(v_Date_of_Last_payment);

                                    //  xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }

                                else if (str_new == "Starting date:")
                                {


                                    //   MessageBox.Show("in starting date");
                                    v_Starting_date = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Date of classification:")
                                    {
                                        v_Date_of_classification = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Date_of_classification = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //   MessageBox.Show(v_Starting_date);
                                    //    MessageBox.Show(v_Date_of_classification);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }


                                else if (str_new == "End date of contract:")
                                {

                                    //  MessageBox.Show("in end date");
                                    v_End_date_of_contract = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Date of last rescheduling:")
                                    {
                                        v_Date_of_last_rescheduling = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Date_of_last_rescheduling = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //     MessageBox.Show(v_End_date_of_contract);
                                    //   MessageBox.Show(v_Date_of_last_rescheduling);

                                    //  xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }


                                else if (str_new == "Sanction Limit:")
                                {


                                    //    MessageBox.Show("in section");
                                    v_Sanction_Limit = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Method of payment:")
                                    {
                                        v_Method_of_payment = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Method_of_payment = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //     MessageBox.Show(v_Sanction_Limit);
                                    //  MessageBox.Show(v_Method_of_payment);
                                    //   MessageBox.Show((String)(range.Cells[xxx+2, 1] as Excel.Range).Text.ToString());
                                    //  xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }

                                else if (str_new == "Total Disbursement\nAmount:")
                                {

                                    //    MessageBox.Show("in total disburment");
                                    v_Total_Disbursement_Amount = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Payments periodicity:")
                                    {
                                        v_Payments_periodicity = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Payments_periodicity = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //      MessageBox.Show(v_Total_Disbursement_Amount);
                                    //      MessageBox.Show(v_Payments_periodicity);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }

                                else if (str_new == "Total number of\ninstallments:")
                                {

                                    //     MessageBox.Show("in total ninstallments");
                                    v_Total_number_of_installments = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Number of time(s)\nrescheduled:")
                                    {
                                        v_Payments_periodicity = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Payments_periodicity = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //      MessageBox.Show(v_Total_number_of_installments);
                                    //     MessageBox.Show(v_Payments_periodicity);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }

                                else if (str_new == "Installment Amount:")
                                {

                                    //    MessageBox.Show("in Installment Amount:");
                                    v_Installment_Amount = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Remaining installments\nAmount:")
                                    {
                                        v_Remaining_installments_Amount = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Remaining_installments_Amount = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //   MessageBox.Show(v_Installment_Amount);
                                    //    MessageBox.Show(v_Remaining_installments_Amount);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }




                                else if (str_new == "Remaining installments\nNumber:")
                                {

                                    //     MessageBox.Show("in Remaining ");
                                    v_Remaining_installments_Number = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Reorganized credit:")
                                    {
                                        v_Reorganized_credit = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Reorganized_credit = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //MessageBox.Show(v_Remaining_installments_Number);
                                    //MessageBox.Show(v_Reorganized_credit);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }


                                else if (str_new == "Security Amount:")
                                {

                                    //    MessageBox.Show("in Security Amount: ");
                                    v_Security_Amount = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Basis for\nclassification:qualitative\njudgment:")
                                    {
                                        v_Basis_for_classification_qualitative_judgment = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Basis_for_classification_qualitative_judgment = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //MessageBox.Show(v_Security_Amount);
                                    //MessageBox.Show(v_Basis_for_classification_qualitative_judgment);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }


                                else if (str_new == "Third Party guarantee\nAmount:")
                                {

                                    //       MessageBox.Show("in Third Party guarante");
                                    v_Third_Party_guarantee_Amount = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Remarks:")
                                    {
                                        v_Remarks = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Remarks = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //MessageBox.Show(v_Third_Party_guarantee_Amount);
                                    //MessageBox.Show(v_Remarks);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }



                                else if (str_new == "Security Type:")
                                {

                                    int tress = xxx;
                                    v_Security_Type = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();


                                    //detail table

                                    var context23 = new CIBEntities();
                                    //  MessageBox.Show("" + (String)(range.Cells[(rCnt + 2), (cCnt + 1)] as Excel.Range).Text);
                                    var t23 = new DETAILS_OF_INSTALL_Faca //Make sure you have a table called test in DB
                                    {
                                        Ref = v_Ref,
                                        FI_code = v_FI_code,
                                        Branch_code = v_Branch_code,
                                        CIB_contract_code = v_CIB_contract_code,
                                        FI_contract_code = v_FI_contract_code,
                                        Role = v_Role,
                                        Phase = v_Phase,
                                        Facility = v_Facility,
                                        Starting_date = v_Starting_date,
                                        End_date_of_contract = v_End_date_of_contract,
                                        Sanction_Limit = v_Sanction_Limit,
                                        Total_Disbursement_Amount = Convert.ToDecimal( v_Total_Disbursement_Amount),
                                        Total_number_of_installments = v_Total_number_of_installments,
                                        Installment_Amount = Convert.ToDecimal( v_Installment_Amount),
                                        Remaining_installments_Number = v_Remaining_installments_Number,
                                        Security_Amount = Convert.ToDecimal( v_Security_Amount),
                                        Third_Party_guarantee_Amount = v_Third_Party_guarantee_Amount,
                                        Security_Type = v_Security_Type,
                                        Date_of_Last_Update = v_Date_of_Last_Update,
                                        Date_of_Law_suit = v_Date_of_Law_suit,
                                        Date_of_Last_payment = v_Date_of_Last_payment,
                                        Date_of_classification = v_Date_of_classification,
                                        Date_of_last_rescheduling = v_Date_of_last_rescheduling,
                                        Method_of_payment = v_Method_of_payment,
                                        Payments_periodicity = v_Payments_periodicity,
                                        Number_of_time_rescheduled = v_Number_of_time_rescheduled,
                                        Remaining_installments_Amount = v_Remaining_installments_Amount,
                                        Reorganized_credit = v_Reorganized_credit,
                                        Basis_for_classification_qualitative_judgment = v_Basis_for_classification_qualitative_judgment,
                                        Remarks = v_Remarks,
                                        cib_bb_id = master_id,

                                    };




                                    context23.DETAILS_OF_INSTALL_Faca.Add(t23);
                                    context23.SaveChanges();
                                    int d_id = t23.D_id;
                                    label2.Text = " Detail table hasbeen Uplod Compleate ";
                                    label2.ForeColor = Color.Red;

                                    z++;
                                    context23.Dispose();

                                    //detail contact



                                    xxlad = 0;

                                    for (rCnt = row_check; rCnt <= rw; rCnt++)
                                    {

                                        for (cCnt = 1; cCnt <= cl; cCnt++)
                                        {

                                            str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                                            if (str == "Contract History")
                                            {


                                                int xz = 1;
                                                xxlad = 1;
                                                if ((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text == "Date")
                                                {
                                                    xz = 2;

                                                }
                                                else
                                                {
                                                    xz = 3;


                                                }



                                                //      MessageBox.Show(""+xz);
                                                //MessageBox.Show("found detail");
                                                //MessageBox.Show((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text);
                                                //MessageBox.Show((String)(range.Cells[(rCnt + xz+1), cCnt] as Excel.Range).Text);
                                                //MessageBox.Show(""+rCnt);
                                                //MessageBox.Show("" + tress);

                                                while (rCnt < tress)
                                                {
                                                    //MessageBox.Show("in loop");

                                                    if ((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text == "")
                                                    {
                                                        //MessageBox.Show("in if 1");

                                                        if ((String)(range.Cells[(rCnt + xz + 1), cCnt] as Excel.Range).Text == "")
                                                        {
                                                            //  MessageBox.Show("in if 2");

                                                            break;




                                                        }
                                                        else
                                                        {
                                                            xz++;

                                                        }



                                                    }

                                                    else
                                                    {
                                                        //  MessageBox.Show("" + (String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text);
                                                        xz++;
                                                        var context232 = new CIBEntities();
                                                        // MessageBox.Show("" + (String)(range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text);
                                                        var t232 = new D_Contract_History //Make sure you have a table called test in DB
                                                        {
                                                            Date = (String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text,

                                                            Outstanding = (String)(range.Cells[(rCnt + xz), (cCnt + 1)] as Excel.Range).Text,
                                                            Overdue = (String)(range.Cells[(rCnt + xz), (cCnt + 2)] as Excel.Range).Text,

                                                            NPI = (String)(range.Cells[(rCnt + xz), (cCnt + 3)] as Excel.Range).Text,
                                                            Status = (String)(range.Cells[(rCnt + xz), (cCnt + 4)] as Excel.Range).Text,

                                                            Defa = (String)(range.Cells[(rCnt + xz), (cCnt + 5)] as Excel.Range).Text,



                                                            D_id = d_id,



                                                        };

                                                        context232.D_Contract_History.Add(t232);
                                                        if ((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text == "")
                                                        {

                                                        }
                                                        else
                                                        {
                                                            context232.SaveChanges();
                                                        }


                                                        label2.Text = " contrac detail table hasbeen Uplod Compleate ";
                                                        label2.ForeColor = Color.Black;



                                                        context23.Dispose();


                                                    }
                                                }


                                            }

                                        }
                                    }

                                    //other 


                                    xxlad = 0;

                                    for (rCnt = tress; rCnt <= rw; rCnt++)
                                    {

                                        for (cCnt = 1; cCnt <= cl; cCnt++)
                                        {

                                            str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                                            if (str == "Other subjects linked to the same contract")
                                            {
                                                int flag = 0;

                                                int xz = 1;
                                                xxlad = 1;
                                                if ((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text == "CIB subject code                           Role                               Name")
                                                {
                                                    xz++;

                                                    flag = 1;

                                                }


                                                if ((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text == "CIB subject code")
                                                {
                                                    xz = 1;

                                                    flag = 2;
                                                }
                                                else if ((String)(range.Cells[(rCnt + 2), cCnt] as Excel.Range).Text == "CIB subject code")
                                                {
                                                    xz = 3;


                                                    flag = 2;
                                                }



                                                //   MessageBox.Show(""+xz);

                                                // MessageBox.Show("fl" + flag);

                                                //MessageBox.Show("found detail");
                                                //MessageBox.Show((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text);
                                                //MessageBox.Show((String)(range.Cells[(rCnt + xz+1), cCnt] as Excel.Range).Text);
                                                //MessageBox.Show(""+rCnt);
                                                //MessageBox.Show("" + tress);

                                                while (true)
                                                {

                                                    //    MessageBox.Show("in loop");

                                                    if ((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text == "")
                                                    {
                                                        break;


                                                    }
                                                    String v_cib = "", v_role = "", v_name1 = "", v_name2 = "", v_name = "";

                                                    //     MessageBox.Show("" + (String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text);


                                                    if (flag == 1)
                                                    {
                                                        string all = (String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text;
                                                        string[] words = all.Split(' ');
                                                        int zoorow = 0;
                                                        int stat = 1;
                                                        while (zoorow < words.Length)
                                                        {
                                                            if (words[zoorow] != "")
                                                            {
                                                                // MessageBox.Show("" + words[zoorow]);



                                                                if (stat == 1) { v_cib = words[zoorow]; stat++; }
                                                                else if (stat == 2) { v_role = words[zoorow]; stat++; }
                                                                else if (stat == 3) { v_name1 = words[zoorow]; stat++; }
                                                                else if (stat == 4) { v_name2 = words[zoorow]; stat++; }


                                                                zoorow++;
                                                            }
                                                            else
                                                            {

                                                                zoorow++;
                                                            }




                                                        }
                                                        //MessageBox.Show("" + v_cib);
                                                        //MessageBox.Show("" + v_role);
                                                        //MessageBox.Show("" + v_name1);
                                                        //MessageBox.Show("" + v_name2);
                                                        v_name = v_name1 + " " + v_name2;
                                                        // MessageBox.Show("" + v_name);




                                                    }


                                                    else if (flag == 2)
                                                    {




                                                        if ((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text != "")
                                                        {
                                                            // MessageBox.Show("" + words[zoorow]);



                                                            v_cib = (String)(range.Cells[(rCnt + xz), 1] as Excel.Range).Text;
                                                            v_role = (String)(range.Cells[(rCnt + xz), 3] as Excel.Range).Text;
                                                            v_name = (String)(range.Cells[(rCnt + xz), 5] as Excel.Range).Text;




                                                        }

                                                        //MessageBox.Show("" + v_cib);
                                                        //MessageBox.Show("" + v_role);


                                                        //MessageBox.Show("" + v_name);




                                                    }






                                                    var context2321 = new CIBEntities();
                                                    // MessageBox.Show("" + (String)(range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text);
                                                    var t2321 = new d_Other_sub_linked //Make sure you have a table called test in DB
                                                    {
                                                        CIB_s_c = v_cib,

                                                        Role = v_role,
                                                        Name = v_name,




                                                        D_id = d_id,



                                                    };

                                                    context2321.d_Other_sub_linked.Add(t2321);
                                                    if ((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text == "")
                                                    {

                                                    }
                                                    else
                                                    {
                                                        context2321.SaveChanges();
                                                    }


                                                    label2.Text = " contrac detail table hasbeen Uplod Compleate ";
                                                    label2.ForeColor = Color.Black;
                                                    context23.Dispose();
                                                    xz++;


                                                }




                                                break;

                                            }
                                        }

                                        if (xxlad != 0) { break; }
                                    }








                                    row_check = xxx + 2;
                                    break;
                                }


                            }



                        }




                        break;

                    }
                }

                if (xxlad != 0) { break; }
            }








            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

        }

        public void start()
        {
            int fcou = 0;
            int done = 0;

            //   MessageBox.Show("" + dir);
            //while (true)
            //{
            DirectoryInfo d = new DirectoryInfo(dir);//Assuming Test is your Folder
            FileInfo[] Files = d.GetFiles("*.xlsx"); //Getting Text files
            foreach (FileInfo file in Files)
            {
                fcou++;
            }

            this.label7.Text = fcou.ToString();


            foreach (FileInfo file in Files)
            {
                this.label11.Text = file.Name;
                location = dir + "\\" + file.Name;

                //  MessageBox.Show("" + location);
                string sourceFile = location;
                string destinationFile = bac + file.Name;

                // To move a file or folder to a new location:
                this.label1.Text = "Cleaning - ";
                doing();


               this.label1.Text = "Uploading - ";
            //  MessageBox.Show("Clean OK For" + file.Name);

                upload();
            System.IO.File.Move(sourceFile, destinationFile);

            done++;


            this.label8.Text = done.ToString();

            }




            //}



            MessageBox.Show("All files are uploaded and moved for back up");
            //Application.Exit();

        }
        private void Form1_Load(object sender, EventArgs e)
        {
         

         ////   MessageBox.Show("" + dir);
         //   //while (true)
         //   //{
         //       DirectoryInfo d = new DirectoryInfo(dir);//Assuming Test is your Folder
         //       FileInfo[] Files = d.GetFiles("*.xlsx"); //Getting Text files

         //       foreach (FileInfo file in Files)
         //       {

         //           location = dir + "\\" + file.Name;

         //           //  MessageBox.Show("" + location);
         //           string sourceFile = location;
         //           string destinationFile = bac +file.Name;

         //           // To move a file or folder to a new location:
         //           doing();
         //         //  upload();
         //          // System.IO.File.Move(sourceFile, destinationFile);


         //       }


         //       //MessageBox.Show("All files are uploaded and moved for back up");
         //       //Application.Exit();


         //   //}


        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }








//company starts



        private void company()
        {

            int master_id;
            label2.Text = "Detecting CIB subject code";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(location, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            //master table values 
            progressBar1.Value = 30;
            string CibCode = "";
            string UserId = "";
            string DateOfInq = "";
            string FiCode = "";
            string BranchCode = "";
            string FiName = "";







            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();
                    // MessageBox.Show("" + str );

                    if (str == "CIB subject code:")
                    {



                        CibCode = (String)(range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Value2;

                        //   MessageBox.Show("" + CibCode);

                        label2.Text = "CIB SUBJECT CODE FOUND ";
                        break;

                    }
                }

                if (CibCode != "") { break; }


            }
            progressBar1.Value = 40;
            //Master Table Start



            for (rCnt = 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "User ID")
                    {



                        UserId = (String)(range.Cells[(rCnt + 1), cCnt] as Excel.Range).Text;
                        DateOfInq = (String)(range.Cells[(rCnt + 1), (cCnt - 1)] as Excel.Range).Text;
                        FiCode = (String)(range.Cells[(rCnt + 1), (cCnt + 1)] as Excel.Range).Text;
                        BranchCode = (String)(range.Cells[(rCnt + 1), (cCnt + 2)] as Excel.Range).Text;
                        FiName = (String)(range.Cells[(rCnt + 1), (cCnt + 3)] as Excel.Range).Text;



                        break;

                    }
                }

                if (UserId != "") { break; }
            }
            progressBar1.Value = 50;


            var context = new CIBEntities();

            var t = new IMaster //Make sure you have a table called test in DB
            {

                CIB_subject_code = CibCode,
                Date_of_Inquiry = Convert.ToDateTime(DateOfInq),
                User_ID = UserId,
                FI_Code = FiCode,
                Branch_Code = BranchCode,
                FI_Name = FiName,
                file_location = location,
                Upload_date = DateTime.Now.ToString("yyyy-MM-ddThh:mm:ssTZD"),
            };

            context.IMasters.Add(t);
            context.SaveChanges();

            label2.Text = "master table hasbeen Uplod Compleate ";
            label2.ForeColor = Color.Red;

            master_id = t.cib_bb_id;

            //Master Table End
            progressBar1.Value = 60;

            //Inquary table starts

            int inque = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "Trade name")
                    {

                        inque = 1;
                        var context23 = new CIBEntities();
                        //    MessageBox.Show("" + (String)(range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Text);
                        var t23 = new I_INQUIRED //Make sure you have a table called test in DB
                        {
                            

                          
                           

                            Postcode = (String)(range.Cells[(rCnt + 1), (cCnt + 3)] as Excel.Range).Text,
                            Owner_Address = (String)(range.Cells[(rCnt + 2), (cCnt + 1)] as Excel.Range).Text,
                            Trade_name = (String)(range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Text,
                            District = (String)(range.Cells[(rCnt +1), (cCnt + 1)] as Excel.Range).Text,
                            TIN = (String)(range.Cells[rCnt, (cCnt + 3)] as Excel.Range).Text,



                            cib_bb_id = master_id,

                        };

                        context23.I_INQUIRED.Add(t23);
                        context23.SaveChanges();

                        label2.Text = "Inquary table hasbeen Uplod Compleate ";
                        label2.ForeColor = Color.Gold;



                        context23.Dispose();




                        break;

                    }
                }

                if (inque != 0) { break; }
            }



            label2.Text = "Inquary table hasbeen Uplod Compleate ";
            label2.ForeColor = Color.Green;


            progressBar1.Value = 70;

            //Inquary table finbish


            // SUBJECT INFORMATION starts

            int xx = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "Title:")
                    {
                        xx = 1;

                        //    MessageBox.Show("" + (String)(range.Cells[(rCnt + 1), cCnt] as Excel.Range).Text);




                        var context2 = new CIBEntities();

                     
                        

                        var t2 = new Sub__INFO //Make sure you have a table called test in DB
                        {

                            CIB_subject_code = CibCode,



                            Title_Name = (String)(range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Text,
                            Type_of_sub = (String)(range.Cells[(rCnt - 1), (cCnt + 3)] as Excel.Range).Text,
                            Ref_number = (String)(range.Cells[(rCnt+1), (cCnt +1)] as Excel.Range).Text,
                            TIN = (String)(range.Cells[(rCnt + 1), (cCnt + 3)] as Excel.Range).Text,
                            Telephone = (String)(range.Cells[(rCnt + 2), (cCnt + 1)] as Excel.Range).Text,
                            Sector_type = (String)(range.Cells[(rCnt + 2), (cCnt + 3)] as Excel.Range).Text,
                            Remarks = (String)(range.Cells[(rCnt + 5), (cCnt + 1)] as Excel.Range).Text,
                          
                            trade_name = (String)(range.Cells[(rCnt + 1), (cCnt + 3)] as Excel.Range).Text,
                            sector_code = (String)(range.Cells[(rCnt + 3), (cCnt + 1)] as Excel.Range).Text,
                            legal_form = (String)(range.Cells[(rCnt + 3), (cCnt + 3)] as Excel.Range).Text,

                            reg_num = (String)(range.Cells[(rCnt + 4), (cCnt + 1)] as Excel.Range).Text,
                            reg_date = (String)(range.Cells[(rCnt + 4), (cCnt + 3)] as Excel.Range).Text,
                         
                           


                            
                            cib_bb_id = master_id,
                        };

                        context2.Sub__INFO.Add(t2);

                        try {
                        context2.SaveChanges(); }

                        catch (Exception  we) {

                         //   MessageBox.Show("" + we);
                        
                        
                        }
                        label2.Text = "Subject INFO table hasbeen Uplod Compleate ";
                        label2.ForeColor = Color.Black;





                        break;

                    }
                }

                if (xx != 0) { break; }
            }


            progressBar1.Value = 80;



            // SUBJECT INFORMATION Ends
            //Adress table starts










            int xxad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "Address Type")
                    {
                        xxad = 1;

                        //  MessageBox.Show("" + (String)(range.Cells[(rCnt + 1), cCnt] as Excel.Range).Text);

                        int z = 1;

                        while (true)
                        {


                            if ((String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text != "")
                            {
                                var context23 = new CIBEntities();
                                //      MessageBox.Show("" + (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text);
                                var t23 = new I_ADDRESS //Make sure you have a table called test in DB
                                {
                                    Address_Type = (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text,
                                    Address = (String)(range.Cells[(rCnt + z), (cCnt + 1)] as Excel.Range).Text,
                                    Postal_code = (String)(range.Cells[(rCnt + z), (cCnt + 2)] as Excel.Range).Text,
                                    District = (String)(range.Cells[(rCnt + z), (cCnt + 3)] as Excel.Range).Text,
                                    Country = (String)(range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text,
                                    cib_bb_id = master_id,
                                    flag = 0,
                                };

                                context23.I_ADDRESS.Add(t23);
                                context23.SaveChanges();

                                label2.Text = "Adress INFO table hasbeen Uplod Compleate ";
                                label2.ForeColor = Color.Gold;


                                z++;
                                context23.Dispose();
                            }
                            else
                            {

                                break;

                            }

                        }




                        break;

                    }
                }

                if (xxad != 0) { break; }
            }


            //Adress table finish

            //owner starts
            int back = 0;
            xxad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();

                    if (str == "No of reporting Institutes:")
                    { xxad = 1;
                    break;
                    }

                    if (str == "OWNERS LIST")
                    {
                        xxad = 1;

                        //  MessageBox.Show("" + (String)(range.Cells[(rCnt + 1), cCnt] as Excel.Range).Text);

                        int z = 3;

                        while (true)
                        {


                            if ((String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text != "")
                            {
                                var context23 = new CIBEntities();
                                //      MessageBox.Show("" + (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text);
                                var t23 = new owner_list //Make sure you have a table called test in DB
                                {
                                    Name_owner = (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text,
                                    Role = (String)(range.Cells[(rCnt + z), (cCnt + 1)] as Excel.Range).Text,
                                    stay_order = (String)(range.Cells[(rCnt + z), (cCnt + 2)] as Excel.Range).Text,
                                    cib_sub = (String)(range.Cells[(rCnt + z), (cCnt + 3)] as Excel.Range).Text,

                                    cib_bb_id = master_id,


                                };

                                context23.owner_list.Add(t23);
                                context23.SaveChanges();

                                label2.Text = "Owner INFO table hasbeen Uplod Compleate ";
                                label2.ForeColor = Color.Gold;


                                z++;
                                context23.Dispose();
                            }
                            else
                            {

                                break;

                            }

                        }




                        break;

                    }
                }

                if (xxad != 0) { break; }
            }

            //owner ends
           
          //  progressBar1.Value = 70;

            //I_SUM_OF_FACILITY_S_AS_BOR Starts




            int xxb = 0;

            for (rCnt = back + 1; rCnt <= rw; rCnt++)
            {


                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "No of reporting Institutes:")
                    {
                        xxb = 1;

                        //    MessageBox.Show("" + (String)(range.Cells[(rCnt + 1), cCnt] as Excel.Range).Text);





                        var context23 = new CIBEntities();
                        // MessageBox.Show("" + (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text);
                        var t23 = new SUM_OF_FACILITY_S_AS_BOR //Make sure you have a table called test in DB
                        {
                            No_of_reporting_Institutes = Convert.ToInt32((range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Text),
                            No_of_Living_Contracts = Convert.ToInt32((range.Cells[(rCnt + 1), (cCnt + 1)] as Excel.Range).Text),

                            Total_Outstanding_Amount = Convert.ToDecimal((range.Cells[(rCnt + 2), (cCnt + 1)] as Excel.Range).Text),
                            Total_Overdue_Amount = Convert.ToDecimal((range.Cells[rCnt, (cCnt + 3)] as Excel.Range).Text),
                            No_of_Stay_order_contracts = Convert.ToInt32((range.Cells[(rCnt + 1), (cCnt + 3)] as Excel.Range).Text),
                            Total_Outstanding_amount_for_Stay = Convert.ToDecimal((range.Cells[(rCnt + 2), (cCnt + 3)] as Excel.Range).Text),

                            flag = 0,

                            cib_bb_id = master_id,

                        };

                        context23.SUM_OF_FACILITY_S_AS_BOR.Add(t23);
                        context23.SaveChanges();

                        label2.Text = " I_SUM_OF_FACILITY_S_AS_BOR INFO table hasbeen Uplod Compleate ";
                        label2.ForeColor = Color.Gold;

                        context23.Dispose();





                        break;

                    }
                }

                if (xxb != 0) { break; }
            }



            //I_SUM_OF_FACILITY_S_AS_BOR ends
            //1a







      int      xxlad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "1.(A) SUMMARY OF THE FUNDED FACILITIES AS BORROWER & CO-BORROWER")
                    {
                        int z = 6;
                        xxlad = 1;
                        //  MessageBox.Show("in");
                        //  



                        while (true)
                        {


                            if ((String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text != "")
                            {
                                var context23 = new CIBEntities();
                                //  MessageBox.Show("" + (String)(range.Cells[(rCnt + z), (cCnt + 16)] as Excel.Range).Text);
                                var t23 = new SUM_OF_FUNDED_FACILI_AS_BOR //Make sure you have a table called test in DB
                                {
                                    Contract_Category = (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text,
                                    UC_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 1)] as Excel.Range).Text),
                                    UC_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 2)] as Excel.Range).Text),

                                    SMA_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 3)] as Excel.Range).Text),
                                    SMA_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text),

                                    SS_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 5)] as Excel.Range).Text),
                                    SS_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 6)] as Excel.Range).Text),

                                    DF_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 7)] as Excel.Range).Text),
                                    DF_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 8)] as Excel.Range).Text),


                                    B_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 9)] as Excel.Range).Text),
                                    BL_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 10)] as Excel.Range).Text),


                                    BLW_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 11)] as Excel.Range).Text),
                                    BLW_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 12)] as Excel.Range).Text),

                                    Terminated_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 13)] as Excel.Range).Text),
                                    Terminated_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 14)] as Excel.Range).Text),

                                    Requested_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 15)] as Excel.Range).Text),
                                    Requested_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 16)] as Excel.Range).Text),

                                    Stay_Order_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 17)] as Excel.Range).Text),
                                    Stay_Order_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 18)] as Excel.Range).Text),


                                    cib_bb_id = master_id,
                                    flag = 0,
                                };

                                context23.SUM_OF_FUNDED_FACILI_AS_BOR.Add(t23);
                                context23.SaveChanges();

                                label2.Text = " I_1a_SUM_OF_FUNDED_FACILI_AS_BOR table hasbeen Uplod Compleate ";
                                label2.ForeColor = Color.Blue;


                                z++;
                                context23.Dispose();
                            }
                            else
                            {

                                break;

                            }

                        }




                        break;

                    }
                }

                if (xxlad != 0) { break; }
            }



            //1a end
            //1b






            xxlad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "1.(B) SUMMARY OF THE NON-FUNDED FACILITIES AS BORROWER & CO-BORROWER")
                    {
                        int z = 6;
                        xxlad = 1;
                        //  MessageBox.Show("in");
                        //  


                        while (true)
                        {


                            if ((String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text != "")
                            {
                                var context23 = new CIBEntities();
                                // MessageBox.Show("" + (String)(range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text);
                                var t23 = new SUM_OF_NON_FUNDED_FACILI_AS_BOR //Make sure you have a table called test in DB
                                {
                                    Type_of_Financing = (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text,

                                    Living_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 1)] as Excel.Range).Text),
                                    Living_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 2)] as Excel.Range).Text),

                                    Terminated_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 3)] as Excel.Range).Text),
                                    Terminated_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text),

                                    Requested_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 5)] as Excel.Range).Text),
                                    Requested_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 6)] as Excel.Range).Text),

                                    Stay_Order_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 7)] as Excel.Range).Text),
                                    Stay_Order_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 8)] as Excel.Range).Text),

                                    flag = 0,

                                    cib_bb_id = master_id,



                                };

                                context23.SUM_OF_NON_FUNDED_FACILI_AS_BOR.Add(t23);
                                context23.SaveChanges();

                                label2.Text = " I__1b_SUM_OF_NON_FUNDED_FACILI_AS_BOR table hasbeen Uplod Compleate ";
                                label2.ForeColor = Color.Green;


                                z++;
                                context23.Dispose();
                            }
                            else
                            {

                                break;

                            }

                        }




                        break;

                    }
                }

                if (xxlad != 0) { break; }
            }






            //1b end



            //2




            xxb = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {


                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "No of reporting Institutes:")
                    {
                        xxb = 1;

                        //    MessageBox.Show("" + (String)(range.Cells[(rCnt + 1), cCnt] as Excel.Range).Text);





                        var context23 = new CIBEntities();
                        // MessageBox.Show("" + (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text);
                        var t23 = new SUM_OF_FACILITY_S_AS_BOR //Make sure you have a table called test in DB
                        {
                            No_of_reporting_Institutes = Convert.ToInt32((range.Cells[rCnt, (cCnt + 1)] as Excel.Range).Text),
                            No_of_Living_Contracts = Convert.ToInt32((range.Cells[(rCnt + 1), (cCnt + 1)] as Excel.Range).Text),

                            Total_Outstanding_Amount = Convert.ToDecimal((range.Cells[(rCnt + 2), (cCnt + 1)] as Excel.Range).Text),
                            Total_Overdue_Amount = Convert.ToDecimal((range.Cells[rCnt, (cCnt + 3)] as Excel.Range).Text),
                            No_of_Stay_order_contracts = Convert.ToInt32((range.Cells[(rCnt + 1), (cCnt + 3)] as Excel.Range).Text),
                            Total_Outstanding_amount_for_Stay = Convert.ToDecimal((range.Cells[(rCnt + 2), (cCnt + 3)] as Excel.Range).Text),
                            flag = 1,

                            cib_bb_id = master_id,

                        };

                        context23.SUM_OF_FACILITY_S_AS_BOR.Add(t23);
                        context23.SaveChanges();

                        label2.Text = " I_2_SUM_OF_FACILITY_S_AS_GUARANTOR1 INFO table hasbeen Uplod Compleate ";
                        label2.ForeColor = Color.Red;

                        context23.Dispose();





                        break;

                    }
                }

                if (xxb != 0) { break; }
            }


            //2 end

            //2a







            xxlad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "2.(A) SUMMARY OF THE FUNDED FACILITIES AS GUARANTOR")
                    {
                        int z = 6;
                        xxlad = 1;
                        //  MessageBox.Show("in");
                        //  



                        while (true)
                        {


                            if ((String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text != "")
                            {
                                var context23 = new CIBEntities();
                                //  MessageBox.Show("" + (String)(range.Cells[(rCnt + z), (cCnt + 16)] as Excel.Range).Text);
                                var t23 = new SUM_OF_FUNDED_FACILI_AS_BOR //Make sure you have a table called test in DB
                                {
                                    Contract_Category = (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text,


                                    BLW_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 12)] as Excel.Range).Text),
                                    BL_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 10)] as Excel.Range).Text),
                                    Terminated_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 14)] as Excel.Range).Text),
                                    DF_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 8)] as Excel.Range).Text),
                                    SS_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 6)] as Excel.Range).Text),
                                    Requested_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 16)] as Excel.Range).Text),
                                    SMA_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text),
                                    Stay_Order_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 18)] as Excel.Range).Text),
                                    UC_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 2)] as Excel.Range).Text),


                                    UC_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 1)] as Excel.Range).Text),
                                    Stay_Order_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 17)] as Excel.Range).Text),
                                    Requested_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 15)] as Excel.Range).Text),
                                    SMA_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 3)] as Excel.Range).Text),
                                    Terminated_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 13)] as Excel.Range).Text),
                                    BLW_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 11)] as Excel.Range).Text),
                                    SS_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 5)] as Excel.Range).Text),
                                    DF_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 7)] as Excel.Range).Text),
                                    B_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 9)] as Excel.Range).Text),







                                    cib_bb_id = master_id,
                                    flag = 1,
                                };

                                context23.SUM_OF_FUNDED_FACILI_AS_BOR.Add(t23);
                                context23.SaveChanges();

                                label2.Text = " I_2a_SUM_OF_FUNDED_FACILI_AS_GUARANTOR1 table hasbeen Uplod Compleate ";
                                label2.ForeColor = Color.Yellow;


                                z++;
                                context23.Dispose();
                            }
                            else
                            {

                                break;

                            }

                        }




                        break;

                    }
                }

                if (xxlad != 0) { break; }
            }



            //2a end




            //2b






            xxlad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "2.(B) SUMMARY OF THE NON-FUNDED FACILITIES AS GUARANTOR")
                    {
                        int z = 6;
                        xxlad = 1;

                        //   MessageBox.Show("in");



                        while (true)
                        {


                            if ((String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text != "")
                            {
                                var context23 = new CIBEntities();
                                // MessageBox.Show("" + (String)(range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text);
                                var t23 = new SUM_OF_NON_FUNDED_FACILI_AS_BOR //Make sure you have a table called test in DB
                                {
                                    Type_of_Financing = (String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text,

                                    Living_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 1)] as Excel.Range).Text),
                                    Living_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 2)] as Excel.Range).Text),

                                    Terminated_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 3)] as Excel.Range).Text),
                                    Terminated_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text),

                                    Requested_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 5)] as Excel.Range).Text),
                                    Requested_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 6)] as Excel.Range).Text),

                                    Stay_Order_NO = Convert.ToInt32((range.Cells[(rCnt + z), (cCnt + 7)] as Excel.Range).Text),
                                    Stay_Order_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 8)] as Excel.Range).Text),

                                    flag = 1,

                                    cib_bb_id = master_id,



                                };

                                context23.SUM_OF_NON_FUNDED_FACILI_AS_BOR.Add(t23);
                                context23.SaveChanges();

                                label2.Text = " I_2b_SUM_OF_FACILITY_S_AS_GUARANTOR1 table hasbeen Uplod Compleate ";
                                label2.ForeColor = Color.Black;


                                z++;
                                context23.Dispose();
                            }
                            else
                            {

                                break;

                            }

                        }




                        break;

                    }
                }

                if (xxlad != 0) { break; }
            }



            //2b end





            //req starts






            xxlad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "REQUESTED CONTRACT DETAILS")
                    {
                        int z = 3;
                        xxlad = 1;
                        // MessageBox.Show("in");
                        //  



                        while (true)
                        {


                            if ((String)(range.Cells[(rCnt + z), cCnt] as Excel.Range).Text != "")
                            {

                                var context23 = new CIBEntities();
                                // MessageBox.Show("" + (String)(range.Cells[(rCnt + 1), (cCnt +1)] as Excel.Range).Text);
                                var t23 = new REQUESTED_CONTRACT //Make sure you have a table called test in DB
                                {
                                    SL = Convert.ToInt32((range.Cells[(rCnt + z), cCnt] as Excel.Range).Text),

                                    Type_of_Contract = (String)(range.Cells[(rCnt + z), (cCnt + 1)] as Excel.Range).Text,
                                    Facility = (String)(range.Cells[(rCnt + z), (cCnt + 2)] as Excel.Range).Text,

                                    Phase = (String)(range.Cells[(rCnt + z), (cCnt + 3)] as Excel.Range).Text,
                                    Role = (String)(range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text,
                                    FI_Code = (String)(range.Cells[(rCnt + z), (cCnt + 5)] as Excel.Range).Text,
                                    Branch_Code = (String)(range.Cells[(rCnt + z), (cCnt + 6)] as Excel.Range).Text,
                                    Request_date = (String)(range.Cells[(rCnt + z), (cCnt + 7)] as Excel.Range).Text,

                                    Total_Requested_Amount = Convert.ToDecimal((range.Cells[(rCnt + z), (cCnt + 8)] as Excel.Range).Text),
                                    CIB_subject_code = (String)(range.Cells[(rCnt + z), (cCnt + 9)] as Excel.Range).Text,
                                    CIB_contract_code = (String)(range.Cells[(rCnt + z), (cCnt + 10)] as Excel.Range).Text,
                                    FI_0contract_codede = (String)(range.Cells[(rCnt + z), (cCnt + 11)] as Excel.Range).Text,




                                    cib_bb_id = master_id,



                                };

                                context23.REQUESTED_CONTRACT.Add(t23);
                                context23.SaveChanges();

                                label2.Text = " I_2b_SUM_OF_FACILITY_S_AS_GUARANTOR1 table hasbeen Uplod Compleate ";
                                label2.ForeColor = Color.Black;

                                z++;
                                context23.Dispose();

                            }
                            else
                            {

                                break;

                            }

                        }




                        break;

                    }
                }

                if (xxlad != 0) { break; }
            }






            int row_check = 2;


            xxlad = 0;

            for (rCnt = rCnt + 1; rCnt <= rw; rCnt++)
            {

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {

                    str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                    if (str == "DETAILS OF INSTALLMENT FACILITY(S)")
                    {
                        int z = 2;
                        xxlad = 1;

                        row_check = rCnt;

                        // MessageBox.Show("in");

                        string str_new = "";






                        while (str_new != "NOTES:")
                        {

                            //  MessageBox.Show("new loop");
                            string v_Ref = "";
                            string v_FI_code = "";
                            string v_Branch_code = "";
                            string v_CIB_contract_code = "";
                            string v_FI_contract_code = "";

                            string v_Role = "";
                            string v_Phase = "";
                            string v_Facility = "";
                            string v_Starting_date = "";
                            string v_End_date_of_contract = "";
                            string v_Sanction_Limit = "";
                            string v_Total_Disbursement_Amount = "";
                            string v_Total_number_of_installments = "";
                            string v_Installment_Amount = "";
                            string v_Remaining_installments_Number = "";
                            string v_Security_Amount = "";
                            string v_Third_Party_guarantee_Amount = "";
                            string v_Security_Type = "";



                            string v_Date_of_Last_Update = "";
                            string v_Date_of_Law_suit = "";
                            string v_Date_of_Last_payment = "";
                            string v_Date_of_classification = "";
                            string v_Date_of_last_rescheduling = "";
                            string v_Method_of_payment = "";
                            string v_Payments_periodicity = "";
                            string v_Number_of_time_rescheduled = "";
                            string v_Remaining_installments_Amount = "";
                            string v_Reorganized_credit = "";
                            string v_Basis_for_classification_qualitative_judgment = "";
                            string v_Remarks = "";





                            for (int xxx = row_check; rCnt <= rw; xxx++)
                            {


                                str_new = (String)(range.Cells[xxx, 1] as Excel.Range).Text.ToString();

                                if (str_new == "NOTES:")
                                {
                                    //  MessageBox.Show("NOTE");


                                    break;



                                }
                                if (str_new == "Ref")
                                {
                                    v_Ref = (String)(range.Cells[xxx + 1, 1] as Excel.Range).Text.ToString();
                                    v_FI_code = (String)(range.Cells[xxx + 1, 2] as Excel.Range).Text.ToString();
                                    v_Branch_code = (String)(range.Cells[xxx + 1, 3] as Excel.Range).Text.ToString();
                                    v_CIB_contract_code = (String)(range.Cells[xxx + 1, 4] as Excel.Range).Text.ToString();
                                    v_FI_contract_code = (String)(range.Cells[xxx + 1, 5] as Excel.Range).Text.ToString();

                                    xxx = xxx + 1;
                                    //   MessageBox.Show(v_Ref);
                                    //   MessageBox.Show(v_FI_code);
                                    //   MessageBox.Show(v_Branch_code);
                                    //   MessageBox.Show(v_FI_contract_code);
                                    ////  row_check = xxx;
                                    //   break;



                                }

                                else if (str_new == "Role:")
                                {

                                    // MessageBox.Show("in role");
                                    v_Role = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Date of Last Update:")
                                    {
                                        v_Date_of_Last_Update = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Date_of_Last_Update = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //   MessageBox.Show(v_Role);
                                    //  MessageBox.Show(v_Date_of_Last_Update);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }
                                else if (str_new == "Phase:")
                                {
                                    //  MessageBox.Show("in phase");
                                    v_Phase = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Date of Law suit:")
                                    {
                                        v_Date_of_Law_suit = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Date_of_Law_suit = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //   MessageBox.Show(v_Phase);
                                    //   MessageBox.Show(v_Date_of_Last_Update);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }


                                else if (str_new == "Facility:")
                                {

                                    //  / MessageBox.Show("in facility");
                                    v_Facility = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Date of last payment:")
                                    {
                                        v_Date_of_Last_payment = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Date_of_Last_payment = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //   MessageBox.Show(v_Facility);
                                    //   MessageBox.Show(v_Date_of_Last_payment);

                                    //  xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }

                                else if (str_new == "Starting date:")
                                {


                                    //   MessageBox.Show("in starting date");
                                    v_Starting_date = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Date of classification:")
                                    {
                                        v_Date_of_classification = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Date_of_classification = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //   MessageBox.Show(v_Starting_date);
                                    //    MessageBox.Show(v_Date_of_classification);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }


                                else if (str_new == "End date of contract:")
                                {

                                    //  MessageBox.Show("in end date");
                                    v_End_date_of_contract = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Date of last rescheduling:")
                                    {
                                        v_Date_of_last_rescheduling = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Date_of_last_rescheduling = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //     MessageBox.Show(v_End_date_of_contract);
                                    //   MessageBox.Show(v_Date_of_last_rescheduling);

                                    //  xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }


                                else if (str_new == "Sanction Limit:")
                                {


                                    //    MessageBox.Show("in section");
                                    v_Sanction_Limit = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Method of payment:")
                                    {
                                        v_Method_of_payment = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Method_of_payment = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //     MessageBox.Show(v_Sanction_Limit);
                                    //  MessageBox.Show(v_Method_of_payment);
                                    //   MessageBox.Show((String)(range.Cells[xxx+2, 1] as Excel.Range).Text.ToString());
                                    //  xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }

                                else if (str_new == "Total Disbursement\nAmount:")
                                {

                                    //    MessageBox.Show("in total disburment");
                                    v_Total_Disbursement_Amount = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Payments periodicity:")
                                    {
                                        v_Payments_periodicity = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Payments_periodicity = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //      MessageBox.Show(v_Total_Disbursement_Amount);
                                    //      MessageBox.Show(v_Payments_periodicity);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }

                                else if (str_new == "Total number of\ninstallments:")
                                {

                                    //     MessageBox.Show("in total ninstallments");
                                    v_Total_number_of_installments = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Number of time(s)\nrescheduled:")
                                    {
                                        v_Payments_periodicity = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Payments_periodicity = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //      MessageBox.Show(v_Total_number_of_installments);
                                    //     MessageBox.Show(v_Payments_periodicity);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }

                                else if (str_new == "Installment Amount:")
                                {

                                    //    MessageBox.Show("in Installment Amount:");
                                    v_Installment_Amount = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Remaining installments\nAmount:")
                                    {
                                        v_Remaining_installments_Amount = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Remaining_installments_Amount = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //   MessageBox.Show(v_Installment_Amount);
                                    //    MessageBox.Show(v_Remaining_installments_Amount);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }




                                else if (str_new == "Remaining installments\nNumber:")
                                {

                                    //     MessageBox.Show("in Remaining ");
                                    v_Remaining_installments_Number = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Reorganized credit:")
                                    {
                                        v_Reorganized_credit = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Reorganized_credit = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //MessageBox.Show(v_Remaining_installments_Number);
                                    //MessageBox.Show(v_Reorganized_credit);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }


                                else if (str_new == "Security Amount:")
                                {

                                    //    MessageBox.Show("in Security Amount: ");
                                    v_Security_Amount = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Basis for\nclassification:qualitative\njudgment:")
                                    {
                                        v_Basis_for_classification_qualitative_judgment = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Basis_for_classification_qualitative_judgment = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //MessageBox.Show(v_Security_Amount);
                                    //MessageBox.Show(v_Basis_for_classification_qualitative_judgment);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }


                                else if (str_new == "Third Party guarantee\nAmount:")
                                {

                                    //       MessageBox.Show("in Third Party guarante");
                                    v_Third_Party_guarantee_Amount = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();
                                    if ((String)(range.Cells[xxx, 3] as Excel.Range).Text.ToString() == "Remarks:")
                                    {
                                        v_Remarks = (String)(range.Cells[xxx, 4] as Excel.Range).Text.ToString();

                                    }
                                    else
                                    {
                                        v_Remarks = (String)(range.Cells[xxx, 6] as Excel.Range).Text.ToString();

                                    }

                                    //MessageBox.Show(v_Third_Party_guarantee_Amount);
                                    //MessageBox.Show(v_Remarks);

                                    //   xxx++;
                                    // row_check = xxx;
                                    //  break;


                                }



                                else if (str_new == "Security Type:")
                                {

                                    int tress = xxx;
                                    v_Security_Type = (String)(range.Cells[xxx, 2] as Excel.Range).Text.ToString();


                                    //detail table

                                    var context23 = new CIBEntities();
                                    //  MessageBox.Show("" + (String)(range.Cells[(rCnt + 2), (cCnt + 1)] as Excel.Range).Text);
                                    var t23 = new DETAILS_OF_INSTALL_Faca //Make sure you have a table called test in DB
                                    {
                                        Ref = v_Ref,
                                        FI_code = v_FI_code,
                                        Branch_code = v_Branch_code,
                                        CIB_contract_code = v_CIB_contract_code,
                                        FI_contract_code = v_FI_contract_code,
                                        Role = v_Role,
                                        Phase = v_Phase,
                                        Facility = v_Facility,
                                        Starting_date = v_Starting_date,
                                        End_date_of_contract = v_End_date_of_contract,
                                        Sanction_Limit = v_Sanction_Limit,
                                        Total_Disbursement_Amount = Convert.ToDecimal(v_Total_Disbursement_Amount),
                                        Total_number_of_installments = v_Total_number_of_installments,
                                        Installment_Amount = Convert.ToDecimal(v_Installment_Amount),
                                        Remaining_installments_Number = v_Remaining_installments_Number,
                                        Security_Amount = Convert.ToDecimal(v_Security_Amount),
                                        Third_Party_guarantee_Amount = v_Third_Party_guarantee_Amount,
                                        Security_Type = v_Security_Type,
                                        Date_of_Last_Update = v_Date_of_Last_Update,
                                        Date_of_Law_suit = v_Date_of_Law_suit,
                                        Date_of_Last_payment = v_Date_of_Last_payment,
                                        Date_of_classification = v_Date_of_classification,
                                        Date_of_last_rescheduling = v_Date_of_last_rescheduling,
                                        Method_of_payment = v_Method_of_payment,
                                        Payments_periodicity = v_Payments_periodicity,
                                        Number_of_time_rescheduled = v_Number_of_time_rescheduled,
                                        Remaining_installments_Amount = v_Remaining_installments_Amount,
                                        Reorganized_credit = v_Reorganized_credit,
                                        Basis_for_classification_qualitative_judgment = v_Basis_for_classification_qualitative_judgment,
                                        Remarks = v_Remarks,
                                        cib_bb_id = master_id,

                                    };




                                    context23.DETAILS_OF_INSTALL_Faca.Add(t23);
                                    context23.SaveChanges();
                                    int d_id = t23.D_id;
                                    label2.Text = " Detail table hasbeen Uplod Compleate ";
                                    label2.ForeColor = Color.Red;

                                    z++;
                                    context23.Dispose();

                                    //detail contact



                                    xxlad = 0;

                                    for (rCnt = row_check; rCnt <= rw; rCnt++)
                                    {

                                        for (cCnt = 1; cCnt <= cl; cCnt++)
                                        {

                                            str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                                            if (str == "Contract History")
                                            {


                                                int xz = 1;
                                                xxlad = 1;
                                                if ((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text == "Date")
                                                {
                                                    xz = 2;

                                                }
                                                else
                                                {
                                                    xz = 3;


                                                }



                                                //      MessageBox.Show(""+xz);
                                                //MessageBox.Show("found detail");
                                                //MessageBox.Show((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text);
                                                //MessageBox.Show((String)(range.Cells[(rCnt + xz+1), cCnt] as Excel.Range).Text);
                                                //MessageBox.Show(""+rCnt);
                                                //MessageBox.Show("" + tress);

                                                while (rCnt < tress)
                                                {
                                                    //MessageBox.Show("in loop");

                                                    if ((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text == "")
                                                    {
                                                        //MessageBox.Show("in if 1");

                                                        if ((String)(range.Cells[(rCnt + xz + 1), cCnt] as Excel.Range).Text == "")
                                                        {
                                                            //  MessageBox.Show("in if 2");

                                                            break;




                                                        }
                                                        else
                                                        {
                                                            xz++;

                                                        }



                                                    }

                                                    else
                                                    {
                                                        //  MessageBox.Show("" + (String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text);
                                                        xz++;
                                                        var context232 = new CIBEntities();
                                                        // MessageBox.Show("" + (String)(range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text);
                                                        var t232 = new D_Contract_History //Make sure you have a table called test in DB
                                                        {
                                                            Date = (String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text,

                                                            Outstanding = (String)(range.Cells[(rCnt + xz), (cCnt + 1)] as Excel.Range).Text,
                                                            Overdue = (String)(range.Cells[(rCnt + xz), (cCnt + 2)] as Excel.Range).Text,

                                                            NPI = (String)(range.Cells[(rCnt + xz), (cCnt + 3)] as Excel.Range).Text,
                                                            Status = (String)(range.Cells[(rCnt + xz), (cCnt + 4)] as Excel.Range).Text,

                                                            Defa = (String)(range.Cells[(rCnt + xz), (cCnt + 5)] as Excel.Range).Text,



                                                            D_id = d_id,



                                                        };

                                                        context232.D_Contract_History.Add(t232);
                                                        if ((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text == "")
                                                        {

                                                        }
                                                        else
                                                        {
                                                            context232.SaveChanges();
                                                        }


                                                        label2.Text = " contrac detail table hasbeen Uplod Compleate ";
                                                        label2.ForeColor = Color.Black;



                                                        context23.Dispose();


                                                    }
                                                }


                                            }

                                        }
                                    }

                                    //other 


                                    xxlad = 0;

                                    for (rCnt = tress; rCnt <= rw; rCnt++)
                                    {

                                        for (cCnt = 1; cCnt <= cl; cCnt++)
                                        {

                                            str = (String)(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToString();


                                            if (str == "Other subjects linked to the same contract")
                                            {
                                                int flag = 0;

                                                int xz = 1;
                                                xxlad = 1;
                                                if ((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text == "CIB subject code                           Role                               Name")
                                                {
                                                    xz++;

                                                    flag = 1;

                                                }


                                                if ((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text == "CIB subject code")
                                                {
                                                    xz = 1;

                                                    flag = 2;
                                                }
                                                else if ((String)(range.Cells[(rCnt + 2), cCnt] as Excel.Range).Text == "CIB subject code")
                                                {
                                                    xz = 3;


                                                    flag = 2;
                                                }



                                                //   MessageBox.Show(""+xz);

                                                // MessageBox.Show("fl" + flag);

                                                //MessageBox.Show("found detail");
                                                //MessageBox.Show((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text);
                                                //MessageBox.Show((String)(range.Cells[(rCnt + xz+1), cCnt] as Excel.Range).Text);
                                                //MessageBox.Show(""+rCnt);
                                                //MessageBox.Show("" + tress);

                                                while (true)
                                                {

                                                    //    MessageBox.Show("in loop");

                                                    if ((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text == "")
                                                    {
                                                        break;


                                                    }
                                                    String v_cib = "", v_role = "", v_name1 = "", v_name2 = "", v_name = "";

                                                    //     MessageBox.Show("" + (String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text);


                                                    if (flag == 1)
                                                    {
                                                        string all = (String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text;
                                                        string[] words = all.Split(' ');
                                                        int zoorow = 0;
                                                        int stat = 1;
                                                        while (zoorow < words.Length)
                                                        {
                                                            if (words[zoorow] != "")
                                                            {
                                                                // MessageBox.Show("" + words[zoorow]);



                                                                if (stat == 1) { v_cib = words[zoorow]; stat++; }
                                                                else if (stat == 2) { v_role = words[zoorow]; stat++; }
                                                                else if (stat == 3) { v_name1 = words[zoorow]; stat++; }
                                                                else if (stat == 4) { v_name2 = words[zoorow]; stat++; }


                                                                zoorow++;
                                                            }
                                                            else
                                                            {

                                                                zoorow++;
                                                            }




                                                        }
                                                        //MessageBox.Show("" + v_cib);
                                                        //MessageBox.Show("" + v_role);
                                                        //MessageBox.Show("" + v_name1);
                                                        //MessageBox.Show("" + v_name2);
                                                        v_name = v_name1 + " " + v_name2;
                                                        // MessageBox.Show("" + v_name);




                                                    }


                                                    else if (flag == 2)
                                                    {




                                                        if ((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text != "")
                                                        {
                                                            // MessageBox.Show("" + words[zoorow]);



                                                            v_cib = (String)(range.Cells[(rCnt + xz), 1] as Excel.Range).Text;
                                                            v_role = (String)(range.Cells[(rCnt + xz), 3] as Excel.Range).Text;
                                                            v_name = (String)(range.Cells[(rCnt + xz), 5] as Excel.Range).Text;




                                                        }

                                                        //MessageBox.Show("" + v_cib);
                                                        //MessageBox.Show("" + v_role);


                                                        //MessageBox.Show("" + v_name);




                                                    }






                                                    var context2321 = new CIBEntities();
                                                    // MessageBox.Show("" + (String)(range.Cells[(rCnt + z), (cCnt + 4)] as Excel.Range).Text);
                                                    var t2321 = new d_Other_sub_linked //Make sure you have a table called test in DB
                                                    {
                                                        CIB_s_c = v_cib,

                                                        Role = v_role,
                                                        Name = v_name,




                                                        D_id = d_id,



                                                    };

                                                    context2321.d_Other_sub_linked.Add(t2321);
                                                    if ((String)(range.Cells[(rCnt + xz), cCnt] as Excel.Range).Text == "")
                                                    {

                                                    }
                                                    else
                                                    {
                                                        context2321.SaveChanges();
                                                    }


                                                    label2.Text = " contrac detail table hasbeen Uplod Compleate ";
                                                    label2.ForeColor = Color.Black;



                                                    context23.Dispose();































                                                    xz++;


                                                }




                                                break;

                                            }
                                        }

                                        if (xxlad != 0) { break; }
                                    }








                                    row_check = xxx + 2;
                                    break;
                                }






                            }





























































                        }




                        break;

                    }
                }

                if (xxlad != 0) { break; }
            }







            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

        }
        OpenFileDialog ofd = new OpenFileDialog();
        private void button3_Click(object sender, EventArgs e)
        {


            ofd.Filter = "Excel File Only |*.xlsx";
           if( ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
           {

               this.label4.Text = ofd.FileName;


               location = @ofd.FileName;


             //  MessageBox.Show(location);



           }
        }

















        private Excel.Application _app;
        private Excel.Workbooks _books;
        private Excel.Workbook _book;
        protected Excel.Sheets _sheets;
        protected Excel.Worksheet _sheet;
   
        //celan funtion
        public void clean(String sss)
        {





            OpenExcelWorkbook(location);
            _sheet = (Excel.Worksheet)_sheets[1];
            _sheet.Select(Type.Missing);
            //Excel.Range range = _sheet.get_Range("A1:A1", Type.Missing);
            Excel.Range range = _sheet.get_Range(sss, Type.Missing).EntireRow;
            range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

         //range = _sheet.get_Range("A5", Type.Missing).EntireRow;
         //   range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            //range = _sheet.get_Range("A7", Type.Missing).EntireRow;
            //range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            NAR(range);
            NAR(_sheet);
            CloseExcelWorkbook();
            NAR(_book);
            _app.Quit();
            NAR(_app);
          //  MessageBox.Show("oj");








            _app=null;
        _books=null;
       _book=null;
        _sheets=null;
        _sheet = null;
   




        }






        protected void OpenExcelWorkbook(string fileName)
        {
            _app = new Excel.Application();
            if (_book == null)
            {
                _books = _app.Workbooks;
                _book = _books.Open(fileName, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _sheets = _book.Worksheets;

            }
        }
        protected void CloseExcelWorkbook()
        {
            _book.Save();
            _book.Close(false, Type.Missing, Type.Missing);

            
        }
        protected void NAR(object o)
        {
            try
            {
                if (o != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
            finally
            {
                o = null;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            start();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }



























    }
}
