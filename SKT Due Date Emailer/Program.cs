﻿using System;
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
		static Outlook.Application outlookApplication = new Outlook.Application();

		static void Main(string[] args)
		{
			DateTime today = DateTime.UtcNow;
			TimeZoneInfo timeInfo = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
			DateTime userTime = TimeZoneInfo.ConvertTimeFromUtc(today, timeInfo);
			string date = userTime.Month.ToString().PadLeft(2, '0') + "/" + userTime.Day.ToString().PadLeft(2, '0') + "/" + userTime.Year + " " + userTime.ToString("HH:mm:ss tt");

			var dataListReturn = getAllSKTList();
			if (dataListReturn.Count < 1)
				Environment.Exit(0);

			var filtered_list = filter_list(dataListReturn);
			if (filtered_list.Count < 1)
				Environment.Exit(0);

			var html_string = generate_html(filtered_list);

			string subject = "Socket Due Date List - " + userTime.Month + "/" + userTime.Day + "/" + userTime.Year;
			string email_list = "manju@icenginc.com; mike@icenginc.com; pamela@icenginc.com; narendra@icenginc.com; ariane@icenginc.com";
			string temp_list = "nabeelz@icenginc.com";
			sendEmail(subject, html_string, temp_list);
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
			var new_list = new List<socket_entry>();

			for (int i = 0; i < input.Count; i++)
			{
				var entry = input[i];
				//populate the datetime field
				entry.convert_dates(); //converts dates and saves conversion success status

				if (!entry.conversion[2])//if no date in, not received yet
					new_list.Add(entry);

				if (entry.due_date > now && entry.date_in < now) //gold star - early
					new_list.Add(entry);

			}

			return new_list;
		}

		static private string generate_html(List<socket_entry> input)
		{
			string html = "Total PCBs: " + input.Count.ToString() + "<br /><br />";
			html += "<table style='border: 1px solid;padding:px;border-collapse:collapse;font-family:Calibri Light;' cellpadding='10'>";

			html += "<tr>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>Work Order ID</td>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>Vendor</td>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>Part No.</td>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>PCB Work Ext.</td>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>Quantity</td>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>PO Date</td>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>Due Date</td>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>Recv'd Date</td>";
			html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;'>Customer</td>";
			//html += "<td style='border: 1px solid black;text-align:center;font-weight: bold;width:300px;'>Comments</td>";

			html += "</tr>";

			foreach (socket_entry entry in input)
			{
				html += "<td style='border: 1px solid black;text-align:center'>" + entry.work_order_id.ToString() + "</td>";
				html += "<td style='border: 1px solid black;text-align:center'>" + entry.vendor.ToString() + "</td>";
				html += "<td style='border: 1px solid black;text-align:center'>" + entry.PO_num.ToString() + "</td>";
				html += "<td style='border: 1px solid black;text-align:center'>" + entry.part_num.ToString() + "</td>";
				html += "<td style='border: 1px solid black;text-align:center'>" + entry.qty_ordered.ToString() + "</td>";
				html += "<td style='border: 1px solid black;text-align:center'>" + entry.order_date.ToString("MM/dd/yyyy") + "</td>";
				html += "<td style='border: 1px solid black;text-align:center'>" + entry.due_date.ToString("MM/dd/yyyy") + "</td>";
				html += "<td style='border: 1px solid black;text-align:center'>" + entry.date_in.ToString("MM/dd/yyyy") + "</td>";
				html += "<td style='border: 1px solid black;text-align:center'>" + entry.customer.ToString() + "</td>";

				html += "</tr>";
			}

			html += "</table>";

			html += "<br /><br />";

			return html;
		}

		static private void sendEmail(string subject, string emailBody, string toEmailList)
		{
			string subjectEmail = subject;
			string bodyEmail = emailBody;
			string toEmail = toEmailList;

			CreateEmailItem(subjectEmail, toEmail, bodyEmail);

		}//End SendEmailtoContacts

		static private void CreateEmailItem(string subjectEmail,
			   string toEmail, string bodyEmail)
		{
			Outlook.MailItem eMail = (Outlook.MailItem)
				outlookApplication.CreateItem(Outlook.OlItemType.olMailItem);

			eMail.Subject = subjectEmail;
			eMail.To = toEmail;
			eMail.Body = bodyEmail;
			eMail.HTMLBody = bodyEmail;
			eMail.Importance = Outlook.OlImportance.olImportanceLow;
			((Outlook._MailItem)eMail).Send();

		}//End CreateEmailItem
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
