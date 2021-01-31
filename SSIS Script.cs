#region Help:  Introduction to the script task
/* The Script Task allows you to perform virtually any operation that can be accomplished in
 * a .Net application within the context of an Integration Services control flow. 
 * 
 * Expand the other regions which have "Help" prefixes for examples of specific ways to use
 * Integration Services features within this script task. */
#endregion


#region Namespaces
using Microsoft.SqlServer.Dts.Runtime;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Threading;
using System.Web;



#endregion

namespace ST_c3cdf7b47210452e84a6cad2079a23f2
{
	[Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute]
	public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
	{
		#region Help:  Firing Integration Services events from a script
		/* This script task can fire events for logging purposes.
		 * 
		 * Example of firing an error event:
		 *  Dts.Events.FireError(18, "Process Values", "Bad value", "", 0);
		 * 
		 * Example of firing an information event:
		 *  Dts.Events.FireInformation(3, "Process Values", "Processing has started", "", 0, ref fireAgain)
		 * 
		 * Example of firing a warning event:
		 *  Dts.Events.FireWarning(14, "Process Values", "No values received for input", "", 0);
		 * */
		#endregion

		#region Help:  Using Integration Services connection managers in a script
		/* Some types of connection managers can be used in this script task.  See the topic 
		 * "Working with Connection Managers Programatically" for details.
		 * 
		 * Example of using an ADO.Net connection manager:
		 *  object rawConnection = Dts.Connections["Sales DB"].AcquireConnection(Dts.Transaction);
		 *  SqlConnection myADONETConnection = (SqlConnection)rawConnection;
		 *  //Use the connection in some code here, then release the connection
		 *  Dts.Connections["Sales DB"].ReleaseConnection(rawConnection);
		 *
		 * Example of using a File connection manager
		 *  object rawConnection = Dts.Connections["Prices.zip"].AcquireConnection(Dts.Transaction);
		 *  string filePath = (string)rawConnection;
		 *  //Use the connection in some code here, then release the connection
		 *  Dts.Connections["Prices.zip"].ReleaseConnection(rawConnection);
		 * */
		#endregion


		readonly DateTime curDate = DateTime.Now;

		public void Main()
		{
			try
			{
				//Fill and sort DataTable by ascending
				OleDbDataAdapter objDA = new OleDbDataAdapter();
				DataTable objDT = new DataTable();
				objDA.Fill(objDT, Dts.Variables["User::usersList"].Value);
				objDT = SortTable(objDT);

				//Connection to the SSRS server
				object mySSISConnection = Dts.Connections["HTTP Connection Manager"].AcquireConnection(null);
				HttpClientConnection httpConn = new HttpClientConnection(mySSISConnection);
				httpConn.GetServerPassword();
				string ssrsServerURL = ((string)Dts.Variables["$Package::ssrsReportURL"].Value);
				ConvertURL(ref ssrsServerURL);

				//Connection to the SMTP server
				SmtpClient smtpServer = new SmtpClient(Dts.Connections["SMTP relay"].Properties["SmtpServer"].GetValue(Dts.Connections["SMTP relay"]).ToString());
				string folderPath = (string)Dts.Variables["$Package::reportPath"].Value + $"\\temp_{curDate:dd.MM.yyyy}";
				CreateTempFolder(folderPath);
				if (!CheckParam(ssrsServerURL))
				{
					//In case there are no parameters in url
					List<string> emails = GetEmailList();
					foreach (var email in emails)
					{
						string outputFile = OutputFile(folderPath);
						httpConn.ServerURL = ssrsServerURL;
						httpConn.DownloadFile(outputFile, true);
						SendMail(smtpServer, email, outputFile);
					}
				}
				else
				{
					//There are parameters in url. Create reports and send them to users.
					CheckEmailInSQLQuery(objDT);
					List<int> paramColumns = CheckParamQuantity(ssrsServerURL, objDT);
					foreach (DataRow dr in objDT.Rows)
					{
						if ((bool)Dts.Variables["$Package::emailTestConnection"].Value)
						{
							List<string> emails = GetEmailList();
							foreach (var email in emails)
							{
								string outputFile = OutputFile(folderPath);
								httpConn.ServerURL = PasteParamToURL(ssrsServerURL, dr, paramColumns);
								httpConn.DownloadFile(outputFile, true);
								SendMail(smtpServer, email, outputFile);
							}
						}
						else
						{
							string outputFile = OutputFile(folderPath);
							httpConn.ServerURL = PasteParamToURL(ssrsServerURL, dr, paramColumns);
							httpConn.DownloadFile(outputFile, true);
							SendMail(smtpServer, dr[GetEmailColumn(objDT)].ToString(), outputFile);
						}
					}
				}
				DeleteFolder(folderPath);
			}
			catch (Exception ex)
			{
				Dts.TaskResult = (int)ScriptResults.Failure;
				if (ex.Message.Contains("0xC001600E"))
				{
					Dts.Variables["User::emailError"].Value = $"Сформированная ссылка не может быть правильно обработана сервером.";
				}
				else
				{
					Dts.Variables["User::emailError"].Value = $"{ex.Message}";
				}

				return;
			}
			if (Dts.TaskResult != 1)
			{
				Dts.TaskResult = (int)ScriptResults.Success;
			}
		}

		string AddDisclaimer(bool disclaimerBool)
		{
			if (disclaimerBool)
			{
				string disclaimerText = @"<p style=""text-align: justify;""><span style=""color: #808080;"">
				<em><a href="""" target=""_blank"" rel=""noopener""></a></em></span></p>";
				return disclaimerText;
			}
			return "";
		}

		bool CheckParam(string ssrsServerURL)
		{
			int paramLook = ssrsServerURL.IndexOf("&rs:parameterlanguage");
			if (paramLook == -1)
				return false;
			else
				return true;
		}
		void CheckEmailInSQLQuery(DataTable objDT)
		{
			/*Looking for column with name "Email" and analyzing text in cells.
			Analyzing part is using template: email must contain @ in it*/
			int i = 0;
			foreach (DataColumn column in objDT.Columns)
			{
				if (column.ColumnName.ToLower() == "email")
				{
					i++;
					foreach (DataRow dr in objDT.Rows)
					{
						if (!objDT.Rows[objDT.Rows.IndexOf(dr)][column.Ordinal].ToString().ToLower().Contains("@"))
						{
							throw new ArgumentException("В столбце \"Email\" SQL запроса отсутствуют электронные адреса или адреса указаны не корректно.");
						}
					}
				}
			}
			if (i == 0)
			{
				throw new ArgumentException("В результате SQL запроса отсутствует столбец с именем \"Email\".");
			}
		}

		int GetEmailColumn(DataTable objDT)
		{
			int columnIndex = 0;
			foreach (DataColumn column in objDT.Columns)
			{
				if (column.ColumnName.ToLower() == "email")
				{
					columnIndex = column.Ordinal;
				}
			}
			return columnIndex;
		}

		List<int> CheckParamQuantity(string url, DataTable objDT)
		{
			/*This methode return list of columns which contains parameter template.
			 *Also it check params in sql query result and compare it with quantity of params in url.
			 */

			int iSQL = 0;
			int iURL = 0;
			string urlTemplate = "&rs:parameterlanguage";
			List<int> paramArray = new List<int>();
			foreach (DataColumn column in objDT.Columns)
			{
				if (column.ColumnName.ToLower().Contains(((string)Dts.Variables["$Package::SQLParamTemplate"].Value).ToLower()))
				{
					iSQL++;
					paramArray.Add(column.Ordinal);
				}
			}
			if (url.Contains(urlTemplate))
			{
				int paramStart = url.IndexOf('&') + 1;
				int paramEnd = url.IndexOf(urlTemplate);

				while (paramStart < paramEnd)
				{
					int paramNameEnd = url.IndexOf('=', paramStart) + 1;
					paramStart = url.IndexOf('&', paramNameEnd);
					paramEnd = url.IndexOf(urlTemplate);
					iURL++;
				}
			}
			if ((iSQL - iURL) < 0)
			{
				throw new ArgumentException("Количество параметров возвращаемых SQL запросом меньше, чем требуется для получения отчета. " +
					$"Или в названиях столбцов отсутствует префикс \"{(string)Dts.Variables["$Package::SQLParamTemplate"].Value}\" ");
			}
			return paramArray;
		}

		string ConvertURL(ref string url)
		{
			/*
			 *Get ATOM-like url and convert it to HTTP-like format.
			 *After url conversion replace param names with case sensetive original names.
			 */

			string urlTemplate = "&amp;rs%3AParameterLanguage=";
			string keyPhrase = "&amp;";
			string format = ((string)Dts.Variables["$Package::ssrsReportFormat"].Value).ToLower();
			int paramStart = url.IndexOf(keyPhrase) + keyPhrase.Length;
			int paramEnd = url.IndexOf(urlTemplate);
			int paramNameEnd;
			List<string> paramArray = new List<string>();

			if (url.Contains(urlTemplate))
			{
				while (paramStart < paramEnd)
				{
					paramNameEnd = url.IndexOf('=', paramStart);
					paramArray.Add(url.Substring(paramStart, (paramNameEnd - paramStart)));
					paramStart = url.IndexOf(keyPhrase, paramNameEnd) + keyPhrase.Length;
					paramEnd = url.IndexOf(urlTemplate);
				}
			}
			url = url.ToLower().Replace(keyPhrase, "&").Replace("%3a", ":").Replace("format=atom", $"format={format}");
			url = url.Substring(0, (url.IndexOf($"format={format}") + ($"format={format}").Length));
			urlTemplate = "&rs:parameterlanguage";
			paramStart = url.IndexOf('&') + 1;
			paramEnd = url.IndexOf(urlTemplate);
			for (int i = 0; paramStart < paramEnd; i++)
			{
				paramNameEnd = url.IndexOf('=', paramStart);
				url = url.Remove(paramStart, (paramNameEnd - paramStart)).Insert(paramStart, paramArray[i]);
				paramStart = url.IndexOf('&', paramNameEnd) + 1;
				paramEnd = url.IndexOf(urlTemplate);
			}
			return url;
		}

		void CreateTempFolder(string folderPath)
		{
			//Creating folder where files should be saved, if folder exist it will be deleted to prevent errors
			if (!Directory.Exists(folderPath))
			{
				_ = Directory.CreateDirectory(folderPath);
			}
			else if (Directory.Exists(folderPath) & !IsDirectoryEmpty(folderPath))
			{
				if (!DeleteFolder(folderPath))
				{
					folderPath = (string)Dts.Variables["$Package::reportPath"].Value + $@"\temp_{curDate:dd.MM.yyyy_HH.mm.ss}";
				}
				_ = Directory.CreateDirectory(folderPath);
			}
		}

		bool DeleteFolder(string path)
		{
			try
			{
				int timeOut = 0;
				do
				{
					Directory.Delete(path, true);
					Thread.Sleep(100);
					timeOut += 1;
				} while (Directory.Exists(path) & timeOut != 20);
				return true;
			}
			catch (Exception ex)
			{
				Dts.TaskResult = (int)ScriptResults.Failure;
				Dts.Variables["User::emailError"].Value = $"{ex.Message} Был создан временный каталог. " +
					$"Дирректория в которой произошла ошибка {path}.";
				return false;
			}
		}

		string EmailBody()
		{
			string mailBody = @"<h3>Отчёт находиться во вложении.</h3>" + AddDisclaimer((bool)Dts.Variables["$Package::emailDisclaimer"].Value);
			return mailBody;
		}

		bool IsDirectoryEmpty(string path)
		{
			return !Directory.EnumerateFileSystemEntries(path).Any();
		}

		List<string> GetEmailList()
		{
			List<string> emailListArray = new List<string>();
			string email;
			if ((bool)Dts.Variables["$Package::emailTestConnection"].Value)
			{
				email = (string)Dts.Variables["$Package::emailTestAdress"].Value;
				if (email == "")
				{
					throw new ArgumentException("Электронный адрес для отправки отчета не указан.");
				}
				else if (!email.Contains(';') & !email.Contains(','))
				{
					if (!email.ToLower().Contains("@"))
					{
						throw new ArgumentException("Адреса \"Email\" указаны не корректно.");
					}
					emailListArray.Add(email);
					return emailListArray;
				}
				else
				{
					emailListArray = GetEmailList(email);
					return emailListArray;
				}
			}
			else
			{
				email = (string)Dts.Variables["$Package::emailTo"].Value;
				if (email == "")
				{
					throw new ArgumentException("Электронный адрес для отправки отчета не указан.");
				}
				else if (!email.Contains(';') | !email.Contains(','))
				{
					if (!email.ToLower().Contains("@"))
					{
						throw new ArgumentException("Адреса \"Email\" указаны не корректно.");
					}
					emailListArray.Add(email);
					return emailListArray;
				}
				else
				{
					emailListArray = GetEmailList(email);
					return emailListArray;
				}
			}
		}

		List<string> GetEmailList(string emailList)
		{
			List<string> emailListArray = new List<string>();
			List<int> separatorPosition = new List<int> { 0 };
			int startindex = emailList.IndexOf(',');
			while (startindex != -1)
			{
				startindex += 1;
				separatorPosition.Add(startindex);
				startindex = emailList.IndexOf(',', startindex);
			}
			startindex = emailList.IndexOf(';');
			while (startindex != -1)
			{
				startindex += 1;
				separatorPosition.Add(startindex);
				startindex = emailList.IndexOf(';', startindex);
			}
			separatorPosition.Sort();
			separatorPosition.Add(emailList.Length + 1);

			for (int i = 0; i < separatorPosition.Count - 1; i++)
			{
				if (emailList.Substring(separatorPosition[i], separatorPosition[i + 1] - 1 - separatorPosition[i]).Trim() != "")
				{
					emailListArray.Add((emailList.Substring(separatorPosition[i], separatorPosition[i + 1] - 1 - separatorPosition[i])).Trim());
				}
			}
			return emailListArray;
		}

		string OutputFile(string folderPath)
		{
			string outputFile = $@"{folderPath}\{(string)Dts.Variables["$Package::emailSubject"].Value}_{curDate:dd.MM.yyyy}.{((string)Dts.Variables["$Package::ssrsReportFormat"].Value).ToLower()}";
			return outputFile;
		}

		string PasteParamToURL(string url, DataRow dr, List<int> paramColumns)
		{
			string urlTemplate = "&rs:parameterlanguage";
			if (url.Contains(urlTemplate))
			{
				int paramStart = url.IndexOf('&') + 1;
				int paramEnd = url.IndexOf(urlTemplate);
				for (int i = 0; paramStart < paramEnd; i++)
				{
					int paramNameEnd = url.IndexOf('=', paramStart) + 1;
					url = url.Remove(paramNameEnd, (url.IndexOf('&', paramNameEnd) - paramNameEnd));
					url = url.Insert(paramNameEnd, HttpUtility.UrlEncode(dr[paramColumns[i]].ToString()));
					paramStart = url.IndexOf('&', paramNameEnd) + 1;
					paramEnd = url.IndexOf(urlTemplate);
				}
			}
			return url;
		}

		void SendMail(SmtpClient SMTP, string sendMailTo, string sendMailAttachments)
		{
			MailMessage email = new MailMessage();
			email.To.Add(sendMailTo);
			email.Subject = (string)Dts.Variables["$Package::emailSubject"].Value;
			email.Body = EmailBody();
			email.From = new MailAddress((string)Dts.Variables["$Package::emailFrom"].Value);
			email.IsBodyHtml = true;
			email.Priority = MailPriority.High;
			email.Attachments.Add(new Attachment(sendMailAttachments));
			if ((bool)Dts.Variables["$Package::emailBCC"].Value)
			{
				email.Bcc.Add((string)Dts.Variables["$Package::emailBCCUser"].Value);
			}
			SMTP.Send(email);
			email.Attachments.Dispose();
		}

		DataTable SortTable(DataTable DT)
		{
			List<string> columnNames = new List<string>();
			foreach (DataColumn column in DT.Columns)
			{
				columnNames.Add(column.ColumnName.ToString());
			}
			columnNames.Sort();
			string[] cNamesArr = columnNames.ToArray();
			DataTable newDT = DT.DefaultView.ToTable(false, cNamesArr);
			return newDT;
		}

		#region ScriptResults declaration
		/// <summary>
		/// This enum provides a convenient shorthand within the scope of this class for setting the
		/// result of the script.
		/// 
		/// This code was generated automatically.
		/// </summary>
		enum ScriptResults
		{
			Success = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success,
			Failure = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure
		};
		#endregion

	}
}