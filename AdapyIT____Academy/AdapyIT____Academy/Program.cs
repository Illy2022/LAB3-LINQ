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

                var dlName = from dlt in dbContext.Courses
                             select dlt.CourseName;

                foreach (string n in dlName)
                {
                    Console.WriteLine("{0}", n);
                }

                var dlTraining = from tlt in dbContext.Trainings
                                 select tlt.TrainingStartDate;
                foreach(DateTime st in dlTraining)
                {
                    Console.WriteLine("{0}", st);
                }

            }




            Console.WriteLine("************************************************************************************************************************ ");
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


            Console.WriteLine("\n **********************************************************************************************************************");
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

            Console.WriteLine("\n **********************************************************************************************************************");
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


            Console.ReadKey();

        }

       
    }
}
