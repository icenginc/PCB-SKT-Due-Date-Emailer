using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PCB_Due_Date_Emailer
{
	class Program
	{
		static Outlook.Application outlookApplication = new Outlook.Application();

		static void Main(string[] args)
		{
			DateTime today = DateTime.UtcNow;
			TimeZoneInfo timeInfo = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
			DateTime userTime = TimeZoneInfo.ConvertTimeFromUtc(today, timeInfo);
			string date = userTime.Month.ToString().PadLeft(2, '0') + "/" + userTime.Day.ToString().PadLeft(2, '0') + "/" + userTime.Year + " " + userTime.ToString("HH:mm:ss tt");

			var dataListReturn = getAllPCBList();
			if (dataListReturn.Count < 1)
				Environment.Exit(0);

			var filtered_list = filter_list(dataListReturn);
			if (filtered_list.Count < 1)
				Environment.Exit(0);

			var colored_list = assign_colors(filtered_list);

			var html_string = generate_html(colored_list);

			string subject = "PCB Due Date List - " + userTime.Month + "/" + userTime.Day + "/" + userTime.Year;
			string email_list = "manju@icenginc.com; pamela@icenginc.com; jing@icenginc.com; ariane@icenginc.com";
			string temp_list = "nabeelz@icenginc.com";
			string cc_list = "mike@icenginc.com; nabeelz@icenginc.com";
			sendEmail(subject, html_string, "", cc_list);
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
			var new_list = new List<PCB_entry>();
			for (int i = 0; i < input.Count; i++)
			{
				var entry = input[i];
				//populate the datetime field
				entry.convert_dates(); //converts dates and saves conversion success status

				if (!entry.conversion[2])//if no date in, not received yet
					new_list.Add(entry);

				else if (entry.due_date > now && entry.date_in < now) //gold star - early
					new_list.Add(entry);
					
			}

			return new_list;
		}

		static private List<PCB_entry> assign_colors(List<PCB_entry> input)
		{
			var now = DateTime.Now;

			foreach (PCB_entry entry in input)
			{
				if (entry.due_date < now && !entry.conversion[2]) //due date overdue, no received date
					entry.color = "red";
				else if (entry.due_date.AddDays(-3) < now) //within n days
					entry.color = "yellow";
				if (entry.due_date.AddDays(-3) > now) //outside of n days
					entry.color = "green";
				if (entry.due_date > now && entry.date_in < now && entry.conversion[2]) //also check the bool, otherwise we will read 01/01/01
					entry.color = "blue";
				if (entry.due_date_string == "" || entry.job_num == "0" || entry.qty_ordered == "" || entry.customer == "" || entry.vendor == "" || entry.pcb_work_ext == "")
					entry.color = "orange";
			}

			return input;
		}

		static private string generate_html(List<PCB_entry> input)
		{
			var now = DateTime.Now;

			string html = "Total PCBs: " + input.Count.ToString() + "<br /><br />";
			html += "<table style='border: 1px solid;padding:px;border-collapse:collapse;font-family:Calibri Light;' cellpadding='10'>";

			html += "<tr>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>Job Number</td>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>Vendor</td>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>WO</td>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>Qty.</td>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>PO Date</td>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>Due Date</td>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>Recv'd Date</td>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>Customer</td>";
			//html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;width:300px;'>Comments</td>";

			html += "</tr>";

			foreach (PCB_entry entry in input)
			{
				if (entry.color == "red")//red
					html += "<tr style='background-Color:#F65E5E'>"; //f24004
				else if (entry.color == "yellow") //yellow
					html += "<tr style='background-Color:#FFFAA3'>";
				else if (entry.color == "green")// green
					html += "<tr style='background-Color:#FFFFFF'>"; //default
				else if (entry.color == "blue")//blue
					html += "<tr style='background-Color:#70A9E9'>";
				else if (entry.color == "orange")//orange
					html += "<tr style='background-Color:#F68C5E'>";
				else
					html += "<tr style='background-Color:#FFFFFF'>"; //default

				string date_in = entry.date_in.ToString("MM/dd/yy");
				if (date_in == "01/01/01")
					date_in = "";

				html += "<td style='border: 1px solid black;text-align:center'>" + entry.job_num.ToString() + "</td>";
				html += "<td style='border: 1px solid black;text-align:center'>" + entry.vendor.ToString() + "</td>";
				html += "<td style='border: 1px solid black;text-align:center'>" + entry.pcb_work_ext.ToString() + "</td>";
				html += "<td style='border: 1px solid black;text-align:center'>" + entry.qty_ordered.ToString() + "</td>";
				html += "<td style='border: 1px solid black;text-align:center'>" + entry.order_date.ToString("MM/dd/yy") + "</td>";
				html += "<td style='border: 1px solid black;text-align:center'>" + entry.due_date.ToString("MM/dd/yy") + "</td>";
				html += "<td style='border: 1px solid black;text-align:center'>" + date_in + "</td>";
				html += "<td style='border: 1px solid black;text-align:center'>" + entry.customer.ToString() + "</td>";

				html += "</tr>";
			}

			html += "</table>";

			html += "<br /><br />";

			return html;
		}

		static private void sendEmail(string subject, string emailBody, string toEmailList, string toCC)
		{
			string subjectEmail = subject;
			string bodyEmail = emailBody;
			string toEmail = toEmailList;

			CreateEmailItem(subjectEmail, toEmail, toCC, bodyEmail);

		}//End SendEmailtoContacts

		static private void CreateEmailItem(string subjectEmail, string toEmail, string toCC, string bodyEmail)
		{
			Outlook.MailItem eMail = (Outlook.MailItem)
				outlookApplication.CreateItem(Outlook.OlItemType.olMailItem);

			eMail.Subject = subjectEmail;
			eMail.To = toEmail;
			eMail.CC = toCC;
			eMail.Body = bodyEmail;
			eMail.HTMLBody = bodyEmail;
			eMail.Importance = Outlook.OlImportance.olImportanceNormal;
			((Outlook._MailItem)eMail).Send();

		}//End CreateEmailItem
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

		public string color = "";

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
