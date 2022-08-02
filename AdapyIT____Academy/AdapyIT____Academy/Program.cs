using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AdapyIT____Academy
{
    class Program
    {
        static void Main(string[] args)

            
        {
            DataTableCollection tableCollection;
            AdaptItAcaDataContext dbContext = new AdaptItAcaDataContext();
            string action;


            Console.WriteLine("Enter 1 if you want to update database: ");
            action = Console.ReadLine();

            if (action == "1")
            {

                FileInfo existingFile = new FileInfo(@"C:\Users\zama.phiri\Desktop\ZIPHIRI\Training_Zana.xlsx");
            //use EPPlus

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(existingFile.FullName, FileMode.Open, FileAccess.Read))

            {

                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    });

                    tableCollection = result.Tables;

                    foreach (System.Data.DataColumn table in tableCollection[0].Columns)
                    {

                        Console.WriteLine(table);
                    }

                }

            }


            DataTable dt = tableCollection[0];
           

            if (dt != null)
            {

                for (int i = 0; i < dt.Rows.Count; i++)
                {

                    }
                }
            }

            Console.WriteLine("************************************************************************************************************************ ");
            Console.WriteLine("Enter 2 if you want to list database: ");
            action = Console.ReadLine();
            if (action == "2")
            {

                Console.WriteLine("A List Of All Available Courses:");
                var dlName = from dlt in dbContext.Courses
                             select dlt.CourseName;

                foreach (string n in dlName)
                {
                    Console.WriteLine("{0}", n);
                }

                Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
                Console.WriteLine("Training Registration Closing Date:");
                var dlClosingDate = from tlt in dbContext.CourseTrainings
                                select tlt.RegistraitionClosingDate;
                foreach (DateTime st in dlClosingDate)
                {
                    Console.WriteLine("{0}", st);
                }

                Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
                Console.WriteLine("Training Start Date:");
                var dlStartDate = from tlt in dbContext.Trainings
                                 select tlt.TrainingStartDate;
                foreach(DateTime st in dlStartDate)
                {
                    Console.WriteLine("{0}", st);
                }

                Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
                Console.WriteLine("Training End Date:");
                var dlEndDate = from tlt in dbContext.Trainings
                                 select tlt.TrainingEndDate;
                foreach (DateTime st in dlEndDate)
                {
                    Console.WriteLine("{0}", st);
                }
            }




            Console.WriteLine("\n************************************************************************************************************************ ");
                Console.WriteLine("Enter 3 If You Want To Register A New Delegate For The Training: ");
                action = Console.ReadLine();
                if (action == "3")
                {
                    string connString = (@"Data Source=JHBHO-MICSUP023\SQLEXPRESS;Initial Catalog=AdaptIT Academy;Integrated Security=True");
                    using (SqlConnection con = new SqlConnection(connString))
                    {
                        con.Open();
                        try
                        {
                            Console.Write("\n Connection Successfully Connected");

                            Console.Write("\n Enter Your FirstName: ");
                            string FirstName = Console.ReadLine();

                            Console.Write("\n Enter Your LastName: ");
                            string LastName = Console.ReadLine();

                            Console.Write("\n Enter Your PhoneNumber: ");
                            string PhoneNumber = Console.ReadLine();

                            Console.Write("\n Enter Your Email: ");
                            string Email = Console.ReadLine();

                            Console.Write("\n Enter Your CompanyName: ");
                            string CompanyName = Console.ReadLine();

                            Console.Write("\n Enter Your DietaryRequirement: ");
                            string DietaryRequirement = Console.ReadLine();

                            String insertQuery = "INSERT INTO Delegate (FirstName, LastName, PhoneNumber, Email, CompanyName, DietaryRequirement) " +
                              "VALUES('" + FirstName + "','" + LastName + "','" + PhoneNumber + "','" + Email + "','" + CompanyName + "','" + DietaryRequirement + "')";

                            SqlCommand insertCommand = new SqlCommand(insertQuery, con);
                            insertCommand.ExecuteNonQuery();
                            Console.Write("\n Data stored successfully");
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }

                    }
                }


            Console.WriteLine("\n**********************************************************************************************************************");
            Console.WriteLine("Enter 4 To Enter Course Details: ");
            action = Console.ReadLine();
            if (action == "4")
            {
                string connString = (@"Data Source=JHBHO-MICSUP023\SQLEXPRESS;Initial Catalog=AdaptIT Academy;Integrated Security=True");
                using (SqlConnection con = new SqlConnection(connString))
                {
                    con.Open();
                    try
                    {
                        Console.Write("\n Connection Successfully Connected");

                        Console.Write("\n Enter The CourseCode: ");
                        string CourseCode = Console.ReadLine();

                        Console.Write("\n Enter The CourseName: ");
                        string CourseName = Console.ReadLine();

                        Console.Write("\n Enter The CourseDescription: ");
                        string CourseDescription = Console.ReadLine();

                        String insertQuery = "INSERT INTO Course (CourseCode, CourseName, CourseDescription) " +
                          "VALUES('" + CourseCode + "','" + CourseName + "','" + CourseDescription + "')";

                        SqlCommand insertCommand = new SqlCommand(insertQuery, con);
                        insertCommand.ExecuteNonQuery();
                        Console.Write("\n Data stored successfully");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }

                }
            }

            Console.WriteLine("\n**********************************************************************************************************************");
            Console.WriteLine("Enter 5 To Enter Training Details: ");
            action = Console.ReadLine();
            if (action == "5")
            {
                string connString = (@"Data Source=JHBHO-MICSUP023\SQLEXPRESS;Initial Catalog=AdaptIT Academy;Integrated Security=True");
                using (SqlConnection con = new SqlConnection(connString))
                {
                    con.Open();
                    try
                    {
                        Console.Write("\n Connection Successfully Connected");

                        Console.Write("\n Enter The TrainingStartDate: ");
                        DateTime TrainingStartDate = DateTime.Parse(Console.ReadLine().ToString());

                        Console.Write("\n Enter The TrainingEndDate: ");
                        DateTime TrainingEndDate = DateTime.Parse(Console.ReadLine().ToString());

                        Console.Write("\n Enter The TrainingVenue: ");
                        string TrainingVenue = Console.ReadLine();

                        Console.Write("\n Enter The TrainingVenueTotalSeats: ");
                        int TrainingVenueTotalSeats = int.Parse(Console.ReadLine().ToString());

                        String insertQuery = "INSERT INTO Training (TrainingStartDate, TrainingEndDate, TrainingVenue, TrainingVenueTotalSeats) " +
                          "VALUES('" + TrainingStartDate + "','" + TrainingEndDate + "','" + TrainingVenue + "','" + TrainingVenueTotalSeats + "')";

                        SqlCommand insertCommand = new SqlCommand(insertQuery, con);
                        insertCommand.ExecuteNonQuery();
                        Console.Write("\n Data stored successfully");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }

                }
            }


            Console.WriteLine("\n**********************************************************************************************************************");
            Console.WriteLine("Enter 6 To Enter Course Training Details: ");
            action = Console.ReadLine();
            if (action == "6")
            {
                string connString = (@"Data Source=JHBHO-MICSUP023\SQLEXPRESS;Initial Catalog=AdaptIT Academy;Integrated Security=True");
                using (SqlConnection con = new SqlConnection(connString))
                {
                    con.Open();
                    try
                    {
                        Console.Write("\n Connection Successfully Connected");

                        Console.Write("\n Enter The CourseCode: ");
                        string CourseCode = Console.ReadLine();

                        Console.Write("\n Enter The CourseTrainingCost: ");
                        decimal CourseTrainingCost = decimal.Parse(Console.ReadLine().ToString());

                        Console.Write("\n Enter The RegistraitionClosingDate: ");
                        DateTime RegistraitionClosingDate = DateTime.Parse(Console.ReadLine().ToString());

                     
                        String insertQuery = "INSERT INTO CourseTraining (CourseCode, CourseTrainingCost, RegistraitionClosingDate) " +
                          "VALUES('" + CourseCode + "','" + CourseTrainingCost + "','" + RegistraitionClosingDate + "')";

                        SqlCommand insertCommand = new SqlCommand(insertQuery, con);
                        insertCommand.ExecuteNonQuery();
                        Console.Write("\n Data stored successfully");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }

                }
            }

            Console.WriteLine("\n**********************************************************************************************************************");
            Console.WriteLine("Enter 7 To Enter Address Details: ");
            action = Console.ReadLine();
            if (action == "7")
            {
                string connString = (@"Data Source=JHBHO-MICSUP023\SQLEXPRESS;Initial Catalog=AdaptIT Academy;Integrated Security=True");
                using (SqlConnection con = new SqlConnection(connString))
                {
                    con.Open();
                    try
                    {
                        Console.Write("\n Connection Successfully Connected");

                        Console.Write("\n Enter The PhysicalAddressLine1: ");
                        string PhysicalAddressLine1 = Console.ReadLine();

                        Console.Write("\n Enter The PhysicalAddressLine2: ");
                        string PhysicalAddressLine2 = Console.ReadLine();

                        Console.Write("\n Enter The PhysicalAddressCode: ");
                        int PhysicalAddressCode = int.Parse(Console.ReadLine().ToString());

                        Console.Write("\n Enter The PostalAddressLine1: ");
                        string PostalAddressLine1 = Console.ReadLine();

                        Console.Write("\n Enter The PostalAddressLine2: ");
                        string PostalAddressLine2 = Console.ReadLine();

                        Console.Write("\n Enter The PostalAddressCode: ");
                        int PostalAddressCode = int.Parse(Console.ReadLine().ToString());


                        String insertQuery = "INSERT INTO Address (PhysicalAddressLine1, PhysicalAddressLine2, PhysicalAddressCode, PostalAddressLine1, PostalAddressLine2, PostalAddressCode) " +
                               "VALUES('" + PhysicalAddressLine1 + "','" + PhysicalAddressLine2 + "','" + PhysicalAddressCode + "','" + PostalAddressLine1 + "','" + PostalAddressLine2 + "','" + PostalAddressCode + "')";

                        SqlCommand insertCommand = new SqlCommand(insertQuery, con);
                        insertCommand.ExecuteNonQuery();
                        Console.Write("\n Data stored successfully");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }

                }
            }

            Console.WriteLine("\n************************************************************************************************************************ ");
            Console.WriteLine("Enter 8 To Display A List Of All Delegates Registered: ");
            action = Console.ReadLine();
            if (action == "8")
            {

                Console.WriteLine("Registered Deligate Firstnames:");
                var dlDeliFName = from dlt in dbContext.Delegates
                                 select dlt.FirstName;
                foreach (string n in dlDeliFName)
                {
                    Console.WriteLine("{0}", n);
                }

                Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
                Console.WriteLine("Registered Deligate Lastnames:");
                var dlDeliLName = from dlt in dbContext.Delegates
                                 select dlt.LastName;
                foreach (string n in dlDeliLName)
                {
                    Console.WriteLine("{0}", n);
                }

                Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
                Console.WriteLine("Registered Deligate PhoneNumber Details:");
                var dlDeliPhoneNum = from dlt in dbContext.Delegates
                                 select dlt.PhoneNumber;
                foreach (string n in dlDeliPhoneNum)
                {
                    Console.WriteLine("{0}", n);
                }

                Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
                Console.WriteLine("Registered Deligate Email Address Details:");
                var dlDeliEmail= from dlt in dbContext.Delegates
                                     select dlt.Email;
                foreach (string n in dlDeliEmail)
                {
                    Console.WriteLine("{0}", n);
                }

                Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
                Console.WriteLine("Registered Deligate Company Name Details:");
                var dlDeliCompNam = from dlt in dbContext.Delegates
                                  select dlt.CompanyName;
                foreach (string n in dlDeliCompNam)
                {
                    Console.WriteLine("{0}", n);
                }

                Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
                Console.WriteLine("Registered Deligate Dietary Requirement Details:");
                var dlDeliDietaryReq = from dlt in dbContext.Delegates
                                    select dlt.DietaryRequirement;
                foreach (string n in dlDeliDietaryReq)
                {
                    Console.WriteLine("{0}", n);
                }

                Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
                Console.WriteLine("Registered Deligate Physical Address Details:");
                var dlDeliAdd1 = from dlt in dbContext.Addresses
                                       select dlt.PhysicalAddressLine1;
                foreach (string n in dlDeliAdd1
                {
                    Console.WriteLine("{0}", n);
                }

                Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
                Console.WriteLine("Registered Deligate Postal Address Details:");
                var dlDeliAdd2 = from dlt in dbContext.Addresses
                                select dlt.PostalAddressLine1;
                foreach (string n in dlDeliAdd2)
                {
                    Console.WriteLine("{0}", n);
                }


            }



            Console.ReadKey();

        }

       
    }
}
