using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DLPCheckO365
{
    public partial class Refresh
    {
        
        private void Refresh_Load(object sender, RibbonUIEventArgs e)
        {
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
                this.onButton_Click);
        }
        private void onButton_Click(object sender, RibbonControlEventArgs e)
        { 
            log4net.Config.XmlConfigurator.Configure();
            log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
            Outlook.Application myApplication = new Outlook.Application();
            String oSubject = string.Empty;
            string iSubject = string.Empty;
            int pFrom = 0;
            int pTo = 0;
            Boolean isManagerApproved = false;
            string subject = string.Empty;
            string groupEmailAddr = string.Empty;
            string name = string.Empty;
            string smtpaddress = string.Empty;
            string Fname = string.Empty;
            string lname = string.Empty;
            string fromAddr = string.Empty;
            int mailCount = 0;
            int i = 0;
            string strAttachment = string.Empty;
            string pendingFolder = "Pending-Validation";
            string ApprovedFolder = "Approved-Validation";
            string RejectedFolder = "Rejected-Validation";
            Outlook.NameSpace outlookNS = myApplication.GetNamespace("MAPI");
            Outlook.MAPIFolder mFolder = myApplication.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox);
            Outlook.MAPIFolder iFolder = myApplication.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            mailCount = mFolder.Items.Count; 
            if (mailCount != 0) {
                foreach (Outlook.MailItem o in mFolder.Items)
                {
                    oSubject = o.Subject;
                    try {
                        Outlook.Items myItems = iFolder.Items;
                        //myItems.Sort("[ReceivedTime]", false);
                        foreach (Object oItem in iFolder.Items)
                        {
                            log.Info(i = i + 1);
                            if (oItem is Outlook.MailItem){ 
                                Outlook.MailItem iMail = (Outlook.MailItem)oItem;
                                if (iMail.Subject.Contains("4 Eye Check Required for Email from") == true)
                                {
                                        iSubject = iMail.Subject;
                                        pFrom = iSubject.IndexOf("**[") + "**[".Length;
                                        pTo = iSubject.LastIndexOf("]**");
                                        subject = iSubject.Substring(pFrom, pTo - pFrom);
                                        pFrom = iSubject.IndexOf("<<") + "<<".Length;
                                        pTo = iSubject.LastIndexOf(">>");
                                        groupEmailAddr = iSubject.Substring(pFrom, pTo - pFrom);
                                        name = iMail.SendUsingAccount.UserName;
                                        smtpaddress = iMail.SendUsingAccount.SmtpAddress;
                                        Fname = name.Substring(0, name.IndexOf(".")).ToUpper();
                                        lname = name.Substring(name.IndexOf(".") + 1, name.Length - name.IndexOf(".") - 1).ToUpper();
                                        fromAddr = iMail.SentOnBehalfOfName.ToUpper();
                                        if (o.HTMLBody.Contains("$$EXTERNALEMAILBLOCKER$$" + subject) == true && fromAddr.Split(' ')[0] != lname && fromAddr.Split(' ')[1] != Fname)
                                        //if (o.HTMLBody.Contains("$$EXTERNALEMAILBLOCKER$$" + subject) == true)
                                        {
                                            //log.Info("Received Mail Handler Fired");
                                            if (iSubject.ToUpper().Contains("[APPROVED]") == true || iSubject.ToUpper().Contains("[APPROVE]") == true)
                                            {
                                                o.HTMLBody = o.HTMLBody.Replace("$$EXTERNALEMAILBLOCKER$$" + subject, "$$EXTERNALEMAILVALIDATION$$");
                                                //o.HTMLBody = o.HTMLBody.Replace("$$EXTERNALEMAILBLOCKER$$", "$$EXTERNALEMAILVALIDATION$$");
                                                o.DeferredDeliveryTime = DateTime.Now;
                                                o.Save();
                                                isManagerApproved = true;
                                                //ApprovalEmailItem("Approved Email", smtpaddress, "Requires Approval - Do not change the subject line", o, groupEmailAddr, ApprovedFolder);
                                                try
                                                {
                                                    // MessageBox.Show("Send email");
                                                    o.Send();
                                                    //MessageBox.Show("Sent email");
                                                    log.Info("Success - Force");
                                                    break;
                                                    //Console.WriteLine("Success");
                                                }
                                                catch (System.Exception ex)
                                                {
                                                    log.Error(ex.Message);
                                                }

                                            }
                                            else if (iSubject.ToUpper().Contains("[REJECT]") == true || iSubject.ToUpper().Contains("[REJECTED]") == true)
                                            {
                                                log.Info("External email " + iSubject + " is Rejected");
                                                ApprovalEmailItem("Rejected Email", smtpaddress, "Requires Approval - Do not change the subject line", o, groupEmailAddr, RejectedFolder);
                                                o.Delete();
                                                break;
                                            }
                                            else if (iSubject.ToUpper().Contains("[APPROVED]") == false && iSubject.ToUpper().Contains("[APPROVE]") == false)
                                            {
                                                log.Info("Approve/Reject  - action has to be taken care, Please iniate the request by resending another email" + iSubject);
                                                break;
                                            }
                                        }
                                        else if(o.HTMLBody.Contains("$$EXTERNALEMAILVALIDATION$$")){
                                            o.Send();
                                        }
                                        else if (fromAddr.Split(' ')[0] == lname && fromAddr.Split(' ')[1] == Fname)
                                        {
                                            //MessageBox.Show("Approver and the initiator cannot be same, Please iniate the request by resending another email");
                                        }
                                  
                                }
                             }
                        }
                        log.Info("Force refreshed");
                    }
                    catch
                    {
                        log.Error("Outbox email scan failed");
                    }
                }
    }
            else
            {
                log.Info("Force refresh -- Zeroemails in outbox");
            }
 }

        private void CreateMailItem()
        {
            Outlook.Application Application = new Outlook.Application();
            Outlook.MailItem mailItem = (Outlook.MailItem)Application.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = "This is the subject";
            mailItem.To = "someone@example.com";
            mailItem.Body = "This is the message.";
            mailItem.Importance = Outlook.OlImportance.olImportanceLow;
            mailItem.Display(true);
        }
        public void ApprovalEmailItem(string subjectEmail, string toEmail, string bodyEmail, Outlook.MailItem extEmail, string groupEmailAddr, string subFolder)
        {
            Outlook.Application myApplication = new Outlook.Application();
            try
            {
                Outlook.MailItem eMail = (Outlook.MailItem)
                myApplication.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Application app = new Outlook.Application();
                Outlook.NameSpace ns = app.GetNamespace("MAPI");
                string groupEmailBox = string.Empty;
                //MessageBox.Show("EmailTrigger");
                bool isGroupEmailConnect = ConfigureGroupEmailBoxFolders(groupEmailAddr, "Mailbox - ", subFolder);
                groupEmailBox = "Mailbox - " + groupEmailAddr;
                if (isGroupEmailConnect != true)
                {
                    groupEmailBox = "Boîte aux lettres - " + groupEmailAddr;
                    if (isGroupEmailConnect != true)
                    {
                        isGroupEmailConnect = ConfigureGroupEmailBoxFolders(groupEmailAddr, string.Empty, subFolder);
                        groupEmailBox = groupEmailAddr;
                    }
                }
                if (isGroupEmailConnect == true)
                {
                    Outlook.MAPIFolder outlookFolder = ns.Folders[groupEmailBox].Folders[subFolder];
                    eMail.Subject = subjectEmail;
                    eMail.Attachments.Add(extEmail, Outlook.OlAttachmentType.olEmbeddeditem);
                    eMail.To = toEmail;
                    eMail.Body = bodyEmail;
                    try
                    {
                        ((Outlook._MailItem)eMail).Save();
                        ((Outlook._MailItem)eMail).Move(outlookFolder);
                    }
                    catch (SystemException ex)
                    {
                        //MessageBox.Show("1" + ex.Message);
                    }
                }
                else
                {
                    //MessageBox.Show("Mailbox - " + groupEmailAddr + " needs to be configured properly");
                }
            }
            catch (SystemException ex)
            {
                //MessageBox.Show("0-" + ex.Message);
            }
        }
        private bool ConfigureGroupEmailBoxFolders(string groupEmailAddr, string prefix, string outlookSubFolder)
        {
            Outlook.Application app = new Outlook.Application();
            Outlook.NameSpace ns = app.GetNamespace("MAPI");
            string groupEmailBox = string.Empty;
            bool statut = false;
            groupEmailBox = prefix + groupEmailAddr;//"Boîte aux lettres - " + groupEmailAddr;
            Outlook.MAPIFolder MailBox;
            Outlook.MAPIFolder MailBoxSubFolder;
            try
            {
                MailBox = ns.Folders[groupEmailBox];
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
                //MessageBox.Show(groupEmailBox + " Mailbox not connected ");
            }
            return statut;
        }
    


   }
}
