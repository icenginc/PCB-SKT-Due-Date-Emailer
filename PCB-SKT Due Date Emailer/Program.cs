using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PCB_SKT_Due_Date_Emailer
{
	class Program
	{
		static void Main(string[] args)
		{
			DateTime today = DateTime.UtcNow;
			TimeZoneInfo timeInfo = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
			DateTime userTime = TimeZoneInfo.ConvertTimeFromUtc(today, timeInfo);
			string date = userTime.Month.ToString().PadLeft(2, '0') + "/" + userTime.Day.ToString().PadLeft(2, '0') + "/" + userTime.Year + " " + userTime.ToString("HH:mm:ss tt");

			List<List<string>> dataListReturn = getAllPCBList();
		}

		static private List<List<string>> getAllPCBList()
		{

			List<List<string>> dataListReturn = new List<List<string>>();

			string databaseLocation = "\\\\QA_SERVER\\QA Database\\QUALITY_Database_BE.mdb";
			OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + databaseLocation + ";");

			//Console.Write("Board Inventory Database Location: {0}\n", databaseLocation);

			con.Open();

			try
			{

				if (con.State == ConnectionState.Open)
				{
					Console.Write("Quality Database Opened\n");

					OleDbCommand cmd = new OleDbCommand();
					cmd.Connection = con;

					string statement = "SELECT * FROM [tbl_Equipment PM List] WHERE [PM Completed]=FALSE";


					cmd.CommandText = statement;

					using (OleDbDataReader rdr = cmd.ExecuteReader())
					{

						while (rdr.Read())
						{

							if (rdr.FieldCount > 0)
							{
								//Console.Write("Executing...\n");
								List<string> tempStringList = new List<string>();

								tempStringList.Add(rdr["System Number"].ToString());
								tempStringList.Add(rdr["Equipment List Category"].ToString());
								tempStringList.Add(rdr["Serial No"].ToString());
								tempStringList.Add(rdr["System Description"].ToString());
								tempStringList.Add(rdr["PM Cycle (Months)"].ToString());
								tempStringList.Add(rdr["Last PM Date"].ToString());
								tempStringList.Add(rdr["Next Due Date"].ToString());
								tempStringList.Add(rdr["PM Actual Date"].ToString());
								tempStringList.Add(rdr["Comments"].ToString());

								dataListReturn.Add(tempStringList);

							}//End if

						}//Endof while

					}//End of using


				}//End if connection open

				con.Close();

			}
			catch (Exception err)
			{
				Console.Write("ERROR Accessing Quality DB: {0}\n", err.Message);
			}

			return dataListReturn;

		}//End of getAllPMList
	}
}
