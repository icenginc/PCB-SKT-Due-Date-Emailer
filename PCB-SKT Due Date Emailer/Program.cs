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

			var dataListReturn = getAllPCBList();
			var filtered_list = filter_list(dataListReturn);
		}

		static private List<PCB_entry> getAllPCBList()
		{

			List<PCB_entry> dataListReturn = new List<PCB_entry>();

			string databaseLocation = "\\\\ICEDATA_SERVER\\Log-Book\\BoardDesign_Assy_Database_BE.mdb";
			OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + databaseLocation + ";");

			//Console.Write("Board Inventory Database Location: {0}\n", databaseLocation);

			con.Open();

			try
			{

				if (con.State == ConnectionState.Open)
				{
					Console.Write("PCB Database Opened\n");

					OleDbCommand cmd = new OleDbCommand();
					cmd.Connection = con;

					string statement = "SELECT * FROM [tblPCB_Jobs] WHERE [PCB PO Date] > #6/14/2018#";

					cmd.CommandText = statement;

					using (OleDbDataReader rdr = cmd.ExecuteReader())
					{
						Console.WriteLine("Collecting Information...");
						while (rdr.Read())
						{

							if (rdr.FieldCount > 0)
							{
								//Console.Write("Executing...\n");
								PCB_entry entry = new PCB_entry();

								entry.job_num = (rdr[0].ToString()); //job number
								entry.qty_ordered = (rdr["Qty Ordered"].ToString());
								entry.customer = (rdr["Customer - PCB"].ToString());
								entry.vendor = (rdr["PCB Vendor"].ToString());
								entry.due_date_string = (rdr["PCB Due Date"].ToString());
								entry.date_in_string = (rdr["PCB Recv'd Date"].ToString());
								entry.order_date_string = (rdr["PCB PO Date"].ToString());
								entry.pcb_work_ext = (rdr["PCB_Work Ext"].ToString());
								
								dataListReturn.Add(entry);

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

		static private List<PCB_entry> filter_list(List<PCB_entry> input)
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

	class PCB_entry
	{
		public string due_date_string;
		public string order_date_string;
		public string date_in_string;
		public string qty_ordered;
		public string customer;
		public string job_num;
		public string vendor;
		public string pcb_work_ext;


		public DateTime due_date;
		public DateTime order_date;
		public DateTime date_in;

		public bool[] conversion;

		public void convert_dates()
		{
			var due_date_bool = convert_due_date();
			var order_date_bool = convert_order_date();
			var date_in_bool = convert_date_in();

			conversion = new bool[] { due_date_bool, order_date_bool, date_in_bool };
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
