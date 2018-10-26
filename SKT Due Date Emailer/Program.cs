using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SKT_Due_Date_Emailer
{
	class Program
	{
		static void Main(string[] args)
		{
			DateTime today = DateTime.UtcNow;
			TimeZoneInfo timeInfo = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
			DateTime userTime = TimeZoneInfo.ConvertTimeFromUtc(today, timeInfo);
			string date = userTime.Month.ToString().PadLeft(2, '0') + "/" + userTime.Day.ToString().PadLeft(2, '0') + "/" + userTime.Year + " " + userTime.ToString("HH:mm:ss tt");

			var dataListReturn = getAllSKTList();
			var filtered_list = filter_list(dataListReturn);
		}

		static private List<socket_entry> getAllSKTList()
		{

			List<socket_entry> dataListReturn = new List<socket_entry>();

			string databaseLocation = "\\\\ICEDATA_SERVER\\Log-Book\\BoardDesign_Assy_Database_BE.mdb";
			OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + databaseLocation + ";");

			//Console.Write("Board Inventory Database Location: {0}\n", databaseLocation);

			con.Open();

			try
			{

				if (con.State == ConnectionState.Open)
				{
					Console.Write("SKT Database Opened\n");

					OleDbCommand cmd = new OleDbCommand();
					cmd.Connection = con;

					string statement = "SELECT * FROM [tblSocketTracking] WHERE [SKT - Order Confirmation Date] > #6/14/18#";


					cmd.CommandText = statement;

					using (OleDbDataReader rdr = cmd.ExecuteReader())
					{
						Console.WriteLine("Collecting Entries...");
						while (rdr.Read())
						{

							if (rdr.FieldCount > 0)
							{
								//Console.Write("Executing...\n");
								List<string> tempStringList = new List<string>();
								socket_entry socket = new socket_entry();
								
								socket.PO_num = (rdr["SKT - PO #"].ToString());
								socket.part_num = (rdr["SKT - Part #"].ToString());
								socket.qty_ordered = (rdr["SKT - Qty Ordered"].ToString()); //quantity
								socket.order_date_string = (rdr["SKT - PO Date"].ToString());
								socket.due_date_string = (rdr["SKT - Date Due"].ToString());
								socket.date_in_string = (rdr["SKT - Date In"].ToString());
								socket.customer = (rdr["SKT - Customer"].ToString());
								socket.vendor = (rdr["SKT - Vendor"].ToString());
								socket.work_order_id = (rdr["SKT - Work Order ID"].ToString());
								/*
								tempStringList.Add(rdr["Comments"].ToString());
								*/

								dataListReturn.Add(socket);

							}//End if

						}//Endof while

					}//End of using


				}//End if connection open

				con.Close();
				Console.WriteLine("Connection closed");
			}
			catch (Exception err)
			{
				Console.Write("ERROR Accessing BoardDesign_Assy_Database_BE: {0}\n", err.Message);
			}

			return dataListReturn;

		}//End of getAllPMList

		static private List<socket_entry> filter_list(List<socket_entry> input)
		{
			var now = DateTime.Now;

			for (int i = 0; i < input.Count; i++)
			{
				var entry = input[i];
				//populate the datetime field
				entry.convert_dates(); //converts dates and saves conversion success status

				if (entry.conversion[1])//conversion 2 is order_date, make sure its valid
				{
					if (entry.conversion[2]) //if date in exists (its already done)
					{
						input.RemoveAt(i);
						i--;
					}
				}
			}

			return input;
		}
	}

	class socket_entry
	{
		public string due_date_string;
		public string order_date_string;
		public string date_in_string;
		public string qty_ordered;
		public string PO_num;
		public string part_num;
		public string customer;
		public string work_order_id;
		public string vendor;

		public DateTime due_date;
		public DateTime order_date;
		public DateTime date_in;

		public bool[] conversion;

		public void convert_dates()
		{
			var due_date_bool = convert_due_date();
			var order_date_bool = convert_order_date();
			var date_in_bool = convert_date_in();

			conversion = new bool[]{due_date_bool, order_date_bool, date_in_bool};
		}

		private bool convert_date_in()
		{
			return DateTime.TryParse(date_in_string, out date_in);
		}

		private bool convert_order_date()
		{
			return DateTime.TryParse(order_date_string, out order_date);
		}

		private bool convert_due_date()
		{
			return DateTime.TryParse(due_date_string, out due_date);
		}
	}
}
