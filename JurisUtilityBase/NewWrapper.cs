using System.Globalization;
using System.Runtime.InteropServices;
using System;
using System.Windows.Forms;
using JurisSVR;
using JDataEngine;
using Gizmox.ActiveReports;
using DataDynamics.ActiveReports.Export.Pdf;
using Gizmox.Extensions;
using Gizmox.Controls;
using Gizmox.CSharp;
using Gizmox.RC;
using System.Collections.Generic;
using System.IO;

namespace JurisUtilityBase
{
    public class NewWrapper : Gizmox.RC.IDisposableUnknownRC, Gizmox.Extensions.IClassTerminate
    {
		private JServer _jurisServer;

		public Exception WrapperException { get; set; }

		public NewWrapper()
		{
			var objJuris = default(JurisEntryPoint);
			WrapperException = null;

			//Gizmox: added ErrorHelper // ERROR: Not supported in C#: OnErrorStatement
			var eh = new Gizmox.Extensions.ErrorHelperEx(delegate (Exception err)
			{
				WrapperException = err;
				return false;
			});
			eh.Try(() => objJuris = new JurisEntryPoint());
			eh.Try(() => _jurisServer = objJuris.JurisServer);
			eh.Try(() => objJuris = null);
		}

		/// <summary>
		/// Handles the Dispose event of the IDisposable control.
		/// </summary>
		/// <returns></returns>
		public void IDisposable_Dispose()
		{
			this.Dispose(true);
			GC.SuppressFinalize(this);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		void IDisposable.Dispose()
		{
			IDisposable_Dispose();
		}


		/// <summary>
		/// 
		/// </summary>
		/// <param name="disposing"></param>
		/// <returns></returns>
		protected void Dispose(bool disposing)
		{
			DoTerminate(disposing);
		}


		public void DoTerminate(bool disposing)
		{

			_jurisServer = null;
		}

		~NewWrapper()
		{
			this.Dispose(false);
		}


		/// <summary>
		/// 
		/// </summary>
		/// <param name="company"></param>
		/// <returns></returns>
		public bool LogonCompany(string company)
		{
			try
			{
				_jurisServer.Companies.CurrentCompany = string.Format("Company{0}", company);
				_jurisServer.DataConnections.Datamode = (JDataBase.enmDatamode)Enum.ToObject(typeof(JDataBase.enmDatamode), 0);
				_jurisServer.OpenDatabase(String.Empty, false);
			}
			catch (Exception exception)
			{
				WrapperException = exception;
				Gizmox.CSharp.Information.Err().CaptureException(exception);
				ErrLog(Information.Err().Number.ToString(CultureInfo.InvariantCulture), Information.Err().Description, string.Format("getbillimage:{0}", Gizmox.CSharp.Information.Err().Source));
				Gizmox.CSharp.Information.Err().Clear();
				return false;
			}

			return true;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="userID"></param>
		/// <returns></returns>
		public int LogonUser(string userID)
		{
			try
			{
				return (int)_jurisServer.CurrentUser.LogonTrustedAccount(userID);
			}
			catch (Exception exception)
			{
				WrapperException = exception;
				Gizmox.CSharp.Information.Err().Clear();
				return 0;
			}

		}


		public List<string> logOnAndDoWork(string company,List<Bill> BList, string textBox, string path, bool processExpense)
        {
			List<string> errors = new List<string>();
			if (LogonCompany(company))
			{
				foreach (Bill bb1 in BList)
				{
					if (!bb1.badBill)
					{
						try
						{
							string finalPath = Path.Combine(path, getFileName(bb1, textBox));
							if (!File.Exists(finalPath))
								GetBillImage(Convert.ToInt32(bb1.billNo.ToString()), finalPath);
							if (bb1.hasExpAttach && processExpense)
							{
								foreach (ExpAttachment eex in bb1.exps)
								{
									string expPath = finalPath.Replace(".pdf", "");
									expPath = expPath + "-" + eex.fileName;
									if (!File.Exists(expPath))
										File.WriteAllBytes(expPath, eex.fileData);
								}
							}

						}
						catch (Exception ccs)
						{
							WrapperException = ccs;
							Gizmox.CSharp.Information.Err().CaptureException(ccs);
							ErrLog(Information.Err().Number.ToString(CultureInfo.InvariantCulture), Information.Err().Description, string.Format("getbillimage:{0}", Gizmox.CSharp.Information.Err().Source));
							Gizmox.CSharp.Information.Err().Clear();
							errors.Add("Unable to print bill " + bb1.billNo + " because " + ccs.Message);
							//throw;
						}
					}

				}

			}
			else
            {
				errors.Add("There was an error logging into your books. Check your license");
			}
			return errors;
		}

		private string getFileName(Bill bb, string textBox)
		{
			string output = textBox.Replace(".pdf", "");
			output = output.Replace("pdf", "");

			output = output.Replace("[ClientNum]", bb.clientNo);
			output = output.Replace("[MatterNum]", bb.matterNo);
			output = output.Replace("[ClientName]", bb.clientName);
			output = output.Replace("[MatterName]", bb.matterName);
			output = output.Replace("[BillNum]", bb.billNo.ToString());
			output = output.Replace("[BillDate]", bb.billDate.Replace("/", "-"));
			output = output.Replace("[Clisys]", bb.clisys.ToString());
			output = output.Replace("[Matsys]", bb.matsys.ToString());
			output = output.Replace("[NowDate]", DateTime.Now.ToString("MM/dd/yyyy").Replace("/", "-"));
			output = output.Replace("[NowTime]", DateTime.Now.ToString("MM/dd/yyyy hh:mm tt").Replace("/", "-").Replace(":", " "));
			output = output + ".pdf";
			return output;
		}


		/// <summary>
		/// Gets the bill image.
		/// </summary>
		/// <param name="lBillNo"></param>
		/// <param name="sFileName"></param>
		/// <returns></returns>
		public int GetBillImage(int lBillNo, string sFileName)
		{
			int functionReturnValue = 0;
			var oPP = RefCounter.AddRef(default(PrintPrebill)); //[refcount];

			// JurisSVR.PrintPrebill 
			ActiveReportEx oAR = default(ActiveReportEx);
			PdfExport oPDF = default(PdfExport);
			int lResult = 0;
			string sPDF = String.Empty;
			bool success = false;
			try
			{
				RefCounter.SetRef(ref oPP, new PrintPrebill()); //[refcount]
				try
				{
					lResult = (int)oPP.PreviewArchiveByBillNumber(lBillNo, false);
					success = true;
				}
				catch (Exception vv)
                {
					
					WrapperException = vv;
					Gizmox.CSharp.Information.Err().CaptureException(vv);
					ErrLog(Information.Err().Number.ToString(CultureInfo.InvariantCulture), Information.Err().Description, string.Format("getbillimage:{0}", Gizmox.CSharp.Information.Err().Source));
					functionReturnValue = -1;
					Gizmox.CSharp.Information.Err().Clear();
					success = false;
					throw;
				}
				if (success)
				{
					if (lResult == 0)
					{
						if ((double)oPP.ArchiveType == 2)
						{
							sPDF = oPP.ArchiveImage;
							if (!string.IsNullOrEmpty(sPDF))
							{
								FileSystem.FileOpen(1, sFileName, OpenMode.Binary, OpenAccess.Default, OpenShare.Default, -1);
								FileSystem.FilePut(1, sPDF);
								FileSystem.FileClose(1);
							}
						}

						else

						{
							oPDF = new PdfExport();
							oAR = oPP.ReportObject;
							oAR.Run(false);

							//Gizmox Zoran: commented out filename is put direct in export method oPDF.FileName = sFileName;
							oPDF.Version = PdfVersion.Pdf13;
							oPDF.ImageQuality = (ImageQuality)Enum.ToObject(typeof(ImageQuality), 100);
							oPDF.Export(oAR.Document, sFileName);
							oPDF = null;
							oAR = null;
							RefCounter.SetRef(ref oPP, null); //[refcount]
						}
					}
				}
				RefCounter.RemoveRef(oPP); //[refcount]

				return lResult;
			}
			catch (Exception err)
			{
				WrapperException = err;
				Gizmox.CSharp.Information.Err().CaptureException(err);
				ErrLog(Information.Err().Number.ToString(CultureInfo.InvariantCulture), Information.Err().Description, string.Format("getbillimage:{0}", Gizmox.CSharp.Information.Err().Source));
				functionReturnValue = -1;
				Gizmox.CSharp.Information.Err().Clear();
			}
			RefCounter.RemoveRef(oPP); //[refcount]
			return functionReturnValue;
		}



		/// <summary>
		/// 
		/// </summary>
		/// <param name="errNbr"></param>
		/// <param name="errDesc"></param>
		/// <param name="errSub"></param>
		/// <returns></returns>
		public object ErrLog(string errNbr, string errDesc, string errSub)
		{
			try
			{
				int hFile = FileSystem.FreeFile();

				//FileSystem.FileOpen(hFile, @"c:\Intel\JurisError.log", OpenMode.Append, OpenAccess.Default, OpenShare.Default, -1);
				//FileSystem.Print(hFile, DateAndTime.Now);
				//FileSystem.PrintLine(hFile, string.Format("{0}: {1} - {2}", errNbr, errDesc, errSub));
				//FileSystem.FileClose(hFile);
			}
			catch (Exception vv)
            {
				
            }
			return default(object);
		}





	}
}
