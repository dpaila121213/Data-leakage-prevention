using System;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Net.Mail;

namespace DLPCheckO365
{
    public partial class o365ThisAddIn
    {
        public string PidTagSmtpAddress = "Actual SMTP address";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            log4net.Config.XmlConfigurator.Configure();
            log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
            this.Application.ItemSend += new ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
            this.Application.NewMailEx += new Outlook.ApplicationEvents_11_NewMailExEventHandler(Application_NewMailEx);
        }
        private static Random random = new Random();
        public static bool isManagerApproved;
        string pendingFolder = "Pending-Validation";
        string ApprovedFolder = "Approved-Validation";
        string RejectedFolder = "Rejected-Validation";
        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length).Select(s => s[random.Next(s.Length)]).ToArray());
        }
        private bool ConfigureGroupEmailBoxFolders(string groupEmailAddr, string prefix, string outlookSubFolder)
        {
            log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
            Outlook.Application app = new Outlook.Application();
            Outlook.NameSpace ns = app.GetNamespace("MAPI");
            string groupEmailBox = string.Empty;
            bool statut = false;
            groupEmailBox = groupEmailAddr;//"Boîte aux lettres - " + groupEmailAddr;
            Outlook.MAPIFolder MailBox;
            Outlook.MAPIFolder MailBoxSubFolder;
            try
            {
                try
                {
                    MailBox = ns.Folders[groupEmailBox];
                }
                catch
                {
                    groupEmailBox = prefix + groupEmailAddr;
                    MailBox = ns.Folders[groupEmailBox];
                }
                try
                {
                    MailBoxSubFolder = ns.Folders[groupEmailBox].Folders[outlookSubFolder];
                    statut = true;
                }
                catch (SystemException ex)
                {
                    MailBoxSubFolder = MailBox.Folders.Add(outlookSubFolder);
                    statut = true;
                }
            }
            catch (SystemException ex)
            {
                log.Info(groupEmailBox + " Mailbox not connected ");
                //MessageBox.Show(groupEmailBox + " Mailbox not connected ");
            }
            return statut;
        }

        private void ApprovalEmailItem(string subjectEmail, string toEmail, string bodyEmail, Outlook.MailItem extEmail, string groupEmailAddr, string subFolder)
        {
            log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
            try
            {
                Outlook.MailItem eMail = (Outlook.MailItem)
                this.Application.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Application app = new Outlook.Application();
                Outlook.NameSpace ns = app.GetNamespace("MAPI");
                string groupEmailBox = string.Empty;
                Outlook.MAPIFolder MailBox;
                //MessageBox.Show("EmailTrigger");
                bool isGroupEmailConnect = ConfigureGroupEmailBoxFolders(groupEmailAddr, "Mailbox - ", subFolder);
                try
                {
                    MailBox = ns.Folders[groupEmailAddr];
                    groupEmailBox = groupEmailAddr;
                }
                catch
                {
                    MailBox = ns.Folders["Mailbox - " + groupEmailAddr];
                    groupEmailBox = "Mailbox - " + groupEmailAddr;
                }
                //groupEmailBox = groupEmailAddr;
                
                if (isGroupEmailConnect != true)
                {
                    isGroupEmailConnect = ConfigureGroupEmailBoxFolders(groupEmailAddr, "Boîte aux lettres - ", subFolder);
                    groupEmailBox = "Boîte aux lettres - " + groupEmailAddr;
                    if (isGroupEmailConnect != true)
                    {
                        isGroupEmailConnect = ConfigureGroupEmailBoxFolders(groupEmailAddr, string.Empty, subFolder);
                        groupEmailBox = groupEmailAddr;
                    }
                }
                if (isGroupEmailConnect == true)
                {
                    log.Info("Connected to GroupMailBox");
                    //MessageBox.Show("connected");
                    Outlook.MAPIFolder outlookFolder = ns.Folders[groupEmailBox].Folders[subFolder];
                    eMail.Subject = subjectEmail;
                    eMail.Attachments.Add(extEmail, Outlook.OlAttachmentType.olEmbeddeditem);
                    eMail.To = toEmail;
                    eMail.Body = bodyEmail;
                    //MessageBox.Show("Drafted");
                    log.Info("Drafted");
                    try
                    {
                        ((Outlook._MailItem)eMail).Save();
                        ((Outlook._MailItem)eMail).Move(outlookFolder);
                    }
                    catch (SystemException ex)
                    {
                        log.Info("1" + ex.Message);
                        //MessageBox.Show("1" + ex.Message);
                    }
                }
                else
                {
                    log.Info("Mailbox - " + groupEmailAddr + " needs to be configured properly");
                    //MessageBox.Show("Mailbox - " + groupEmailAddr + " needs to be configured properly");
                }
            }
            catch (SystemException ex)
            {
                log.Info("0 - " + ex.Message);
                //MessageBox.Show("0-" + ex.Message);
            }
        }
        int CountRecipients(Outlook.AddressEntry entry)
        {
            int count = 1;
            if (entry.Members == null)
                return 1;
            else
            {
                foreach (Outlook.AddressEntry e in entry.Members)
                {
                    count = count + CountRecipients(e);
                }
            }
            return count;
        }
        void Application_ItemSend(object Item, ref bool Cancel)
        {
            log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
            Boolean isAnomaly = false, iEmailBlock = false, isAlert = true;
            string message = string.Empty;
            string caption = string.Empty;
            string fromAddr = string.Empty;
            string myEmailAddr = string.Empty;
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result;
            string strMgrEmail = string.Empty;
            //strMgrEmail = getEmail("MgrEmail");
            string strAttachList = string.Empty;
            try
            {
                Outlook.MailItem myItem = Item as Outlook.MailItem;
                Outlook.MailItem OrgEmailCopy = Item as Outlook.MailItem;
                string DistinctEmailAddr = string.Empty;
                string DistinctDomainNames = string.Empty, DistinctInternalDomainNames = string.Empty, DistinctExternalDomainNames = string.Empty;
                string currEmailAddr = string.Empty, currDomainName = string.Empty;
                string attachmentNames = string.Empty;
                int recipientCount = 0, attachmentCount = 0, distinctDomainCount = 0, distinctIntDomainCount = 0, distinctExtDomainCount = 0;
                string uniqueID = RandomString(50);
                try {
                    myEmailAddr = myItem.SendUsingAccount.SmtpAddress;
                } catch {
                    myEmailAddr = myItem.UserProperties["username"].Value;
                }
                
                fromAddr = myItem.SentOnBehalfOfName;
                if (myItem.HTMLBody.IndexOf("<table class") > myItem.HTMLBody.IndexOf("<img"))
                {
                    attachmentCount = attachmentCount + 1;
                }
                if (myItem != null)
                {
                    foreach (Outlook.Attachment attchmnt in myItem.Attachments)
                    {
                        if (!myItem.HTMLBody.ToLower().Contains("cid:" + attchmnt.FileName.ToLower()))
                        { 
                            attachmentCount = attachmentCount + 1;
                            attachmentNames = attachmentNames + attachmentCount.ToString() + ". " + attchmnt.FileName + "\n";
                            isAnomaly = true;
                        }
                    }
                    foreach (Outlook.Recipient recip in myItem.Recipients)
                    {
                        recipientCount += CountRecipients(recip.AddressEntry);
                        currEmailAddr = recip.PropertyAccessor.GetProperty(PidTagSmtpAddress).ToString();
                        if (currEmailAddr.Substring(0, 3).ToUpper() == "DG.")
                        {
                            recipientCount = recipientCount - 1;
                        }
                        if (DistinctEmailAddr.IndexOf(currEmailAddr) == -1)
                        {
                            DistinctEmailAddr = DistinctEmailAddr + currEmailAddr + ";";
                        }
                        currDomainName = currEmailAddr.Split('@')[1].ToLower();
                        if (currDomainName == "domain1.com" || currDomainName == "domain2.com" || currDomainName == "domain3.com" || currDomainName == "domain4.com" || currDomainName == "domain5.com" || currDomainName == "domain6.com" || currDomainName == "domain7.com")
                        {
                            if (DistinctInternalDomainNames.IndexOf(currDomainName) == -1)
                            {
                                distinctIntDomainCount = distinctIntDomainCount + 1;
                                DistinctInternalDomainNames = DistinctInternalDomainNames + currDomainName + ";";
                            }
                        }
                        else
                        {
                            if (DistinctExternalDomainNames.IndexOf(currDomainName) == -1)
                            {
                                distinctExtDomainCount = distinctExtDomainCount + 1;
                                DistinctExternalDomainNames = DistinctExternalDomainNames + currDomainName + ";";
                                isAnomaly = true;
                            }

                        }
                        if (DistinctDomainNames.IndexOf(currDomainName) == -1)
                        {
                            distinctDomainCount = distinctDomainCount + 1;
                            DistinctDomainNames = DistinctDomainNames + currDomainName + ";";
                        }
                        currEmailAddr = string.Empty;
                        currDomainName = string.Empty;
                    }
                    if (distinctExtDomainCount > 0 && attachmentCount > 0 && fromAddr.Length > 1)
                    {
                        iEmailBlock = true;
                        if (distinctExtDomainCount > 1)
                        {
                            MessageBox.Show("Current Email is being sent to " + distinctExtDomainCount + " different clients.");
                        }
                    }
                    if (DistinctEmailAddr.Length > 1)
                    {
                        DistinctEmailAddr = DistinctEmailAddr.Substring(0, DistinctEmailAddr.Length - 1);
                    }
                    if (DistinctDomainNames.Length > 1)
                    {
                        DistinctDomainNames = DistinctDomainNames.Substring(0, DistinctDomainNames.Length - 1);
                    }
                    if (DistinctExternalDomainNames.Length > 1)
                    {
                        DistinctExternalDomainNames = DistinctExternalDomainNames.Substring(0, DistinctExternalDomainNames.Length - 1);
                    }
                    if (DistinctInternalDomainNames.Length > 1)
                    {
                        DistinctInternalDomainNames = DistinctInternalDomainNames.Substring(0, DistinctInternalDomainNames.Length - 1);
                    }
                    if (recipientCount > 25)
                    {
                        isAnomaly = true;
                    }
                    //*******************************************************************************************************
                    if (isAnomaly)
                    {
                        if (distinctExtDomainCount > 0 && attachmentCount > 0 && iEmailBlock == true)
                        {
                            if (isManagerApproved == false)
                            {
                                message = "Identified receipient(s) sending outside SG Group : " + fromAddr;
                                message = message + "\n" + "\n" + "Requires 4-Eye Check approval";
                                isAlert = true;
                            }
                            else
                            {
                                message = "Received 4-Eye Check approval";
                                //isManagerApproved = false;
                                isAlert = false;
                            }
                            caption = "Email is going outside SG with attachment(s)";
                            result = DialogResult.No;
                            if (isAlert == true)
                            {
                                result = MessageBox.Show(String.Format(message, recipientCount), caption, buttons);
                            }
                            else
                            {
                                result = DialogResult.Yes;
                                isAlert = true;
                            }
                            if (result == DialogResult.No)
                            {
                                Cancel = true;
                                return;
                            }
                            else if (distinctExtDomainCount > 0 && iEmailBlock == true)
                            {
                                if (isManagerApproved == false)
                                {
                                    message = "Identified receipient(s) sending outside SG Group : " + fromAddr;
                                    message = message + "\n" + "\n" + "Requires 4-Eye Check approval";
                                    isAlert = true;
                                }
                                else
                                {
                                    message = "Received 4-Eye Check approval";
                                    isManagerApproved = false;
                                    isAlert = false;
                                }
                                caption = "Email is going outside SG with attachment(s)";
                                result = DialogResult.No;
                                if (isAlert == true)
                                {
                                    result = DialogResult.Yes;
                                }
                                else
                                {
                                    result = DialogResult.Yes;
                                    isAlert = true;
                                }
                                if (result == DialogResult.No)
                                {
                                    Cancel = true;
                                    return;
                                }
                            }
                            if (iEmailBlock == true)
                            {
                                try
                                {
                                    if (myItem.HTMLBody.Contains("$$EXTERNALEMAILVALIDATION$$") == false)
                                    {
                                        myItem.HTMLBody = myItem.HTMLBody + "$$EXTERNALEMAILBLOCKER$$" + uniqueID;
                                        myItem.UserProperties.Add("username", OlUserPropertyType.olText);
                                        myItem.UserProperties.Add("smtpdomain", OlUserPropertyType.olText);
                                        myItem.UserProperties["username"].Value = myItem.SendUsingAccount.UserName;
                                        myItem.UserProperties["smtpdomain"].Value = myItem.SendUsingAccount.SmtpAddress.ToString();
                                        myItem.Save();
                                        ApprovalEmailItem("4 Eye Check Required for Email from <<" + fromAddr + ">>**[" + uniqueID + "]**[Approve/Reject]", myEmailAddr, "Requires Approval - Do not change the subject line", OrgEmailCopy, fromAddr, pendingFolder);
                                        myItem.DeferredDeliveryTime = DateTime.Now.AddDays(11111);
                                        log.Info("Approval email saved:" + myItem.Subject );
                                    }
                                    else
                                    {
                                        myItem.HTMLBody = myItem.HTMLBody.Replace("$$EXTERNALEMAILVALIDATION$$", String.Empty);
                                    }
                                }
                                catch (SystemException ex)
                                {
                                    log.Error(ex.Message);
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                log.Info(ex.Message);
            }
            finally
            {

            }
        }
        void Application_NewMailEx(string EntryIDCollection)
        {
            log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
            Outlook.MailItem RecdMail;
            string subject = string.Empty, status = string.Empty;
            string txtBody = string.Empty;
            string fromAddr = string.Empty;
            string smtpaddress = string.Empty;
            string groupEmailAddr = string.Empty;
            string Fname = string.Empty, lname = string.Empty, name = string.Empty;
            bool isApproved = false;
            RecdMail = (Outlook.MailItem)this.Application.Session.GetItemFromID(EntryIDCollection, Type.Missing);
            fromAddr = RecdMail.SentOnBehalfOfName.ToUpper();
            log.Info("Approval email received from - " + fromAddr);
            int pFrom = 0, pTo = 0;
            //SmtpClient client = new SmtpClient();
                Outlook.NameSpace outlookNS = this.Application.GetNamespace("MAPI");
                Outlook.MAPIFolder mFolder = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox);
                foreach (Outlook.MailItem o in mFolder.Items)
                {
                    //MessageBox.Show(o.HTMLBody.Contains("$$EXTERNALEMAILBLOCKER$$" + RecdMail.Subject).ToString());
                    subject = RecdMail.Subject;
                    string mailsubject = string.Empty;
                    status = subject;
                    pFrom = subject.IndexOf("**[") + "**[".Length;
                    pTo = subject.LastIndexOf("]**");
                    try
                    {
                        subject = subject.Substring(pFrom, pTo - pFrom);
                        pFrom = status.IndexOf("<<") + "<<".Length;
                        pTo = status.LastIndexOf(">>");
                        groupEmailAddr = status.Substring(pFrom, pTo - pFrom);
                    }
                    catch
                    {
                        
                    }
                    try
                    {
                        name = o.UserProperties["username"].Value;
                        smtpaddress = o.UserProperties["smtpdomain"].Value;
                    }catch
                    {
                        name = o.SendUsingAccount.UserName;
                        smtpaddress = o.SendUsingAccount.SmtpAddress;
                    }
                        log.Info("Phase2");
                        //MessageBox.Show("Phase2");
                        Fname = name.Substring(0, name.IndexOf(".")).ToUpper();
                        lname = name.Substring(name.IndexOf(".") + 1, name.Length - name.IndexOf(".") - 1).ToUpper();
                
                  if (o.HTMLBody.Contains("$$EXTERNALEMAILBLOCKER$$"+ subject) == true && fromAddr.Split(' ')[0] != lname && fromAddr.Split(' ')[1] != Fname)
                //if (o.HTMLBody.Contains("$$EXTERNALEMAILBLOCKER$$" + subject) == true)
                    {
                        log.Info ("Rcv Mail Handler Fired" + o.Subject);
                        if (status.ToUpper().Contains("[APPROVED]") == true || status.ToUpper().Contains("[APPROVE]") == true)
                        {
                            o.HTMLBody = o.HTMLBody.Replace("$$EXTERNALEMAILBLOCKER$$" + subject, "$$EXTERNALEMAILVALIDATION$$");
                            //o.HTMLBody = o.HTMLBody.Replace("$$EXTERNALEMAILBLOCKER$$", "$$EXTERNALEMAILVALIDATION$$");
                            o.DeferredDeliveryTime = DateTime.Now;
                            o.Save();
                            isManagerApproved = true;
                           ApprovalEmailItem("Approved Email", smtpaddress, "Requires Approval - Do not change the subject line", o, groupEmailAddr, ApprovedFolder);
                            try
                            {
                            // MessageBox.Show("Send email");
                                mailsubject = o.Subject;
                                o.Send();
                                //MessageBox.Show("Sent email");
                                log.Info("sent" + mailsubject);
                                //Console.WriteLine("Success");
                            }
                            catch (System.Exception e)
                            {
                                log.Error(e.Message);
                            }

                        }
                        else if (status.ToUpper().Contains("[REJECT]") == true || status.ToUpper().Contains("[REJECTED]") == true)
                        {
                            MessageBox.Show("External email is Rejected");
                            ApprovalEmailItem("Rejected Email", smtpaddress, "Requires Approval - Do not change the subject line", o, groupEmailAddr, RejectedFolder);
                            o.Delete();

                        }
                        else if (status.ToUpper().Contains("[APPROVED]") == false && status.ToUpper().Contains("[APPROVE]") == false)
                        {
                            MessageBox.Show("Approve/Reject  - action has to be taken care, Please iniate the request by resending another email");
                        }
                    }
                    else if (fromAddr.Split(' ')[0] == lname && fromAddr.Split(' ')[1] == Fname)
                    {
                        MessageBox.Show("Approver and the initiator cannot be same, Please iniate the request by resending another email");
                    }
                }
            //}
            //catch (System.Exception ex)
            //{
            //    MessageBox.Show("3" + ex.Message);
            //}

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
