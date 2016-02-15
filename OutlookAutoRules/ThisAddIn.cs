using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Diagnostics;
using RibbonXOutlook14AddinCS;

namespace OutlookAutoRules
{
    public partial class ThisAddIn
    {
        //private Office.CommandBar menuBar;
        //private Office.CommandBarPopup newMenuBar;
        //private Office.CommandBarButton buttonOne;
        //private Office.CommandBarButton buttonTwo;

        Outlook.Application m_Application;

        private Outlook.MailItem item;
        private Office.CommandBarButton btnCheckForRule;
        private Outlook.Selection selection;
        private List<Outlook.Rule> RulesList = new List<Outlook.Rule>();
        private Outlook.Stores AllStores;
        private List<Office.CommandBarButton> btnRule = new List<Office.CommandBarButton>();
        internal static Office.IRibbonUI m_Ribbon;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
        
            return new RibbonXAddin(m_Application);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //AddMenuBar();
            m_Application = this.Application;

            Debugger.Launch();

            string test = "";
            LoadRules();
        }

        private void LoadRules()
        {
            try
            {
                Application.ItemContextMenuDisplay += new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(Application_ItemContextMenuDisplay);
                Application.ContextMenuClose += new Outlook.ApplicationEvents_11_ContextMenuCloseEventHandler(Application_ContextMenuClose);

                AllStores = Application.Session.Stores;
                foreach (Outlook.Store OS in AllStores)
                {
                    try
                    {
                        foreach (Outlook.Rule OR in OS.GetRules())
                        {
                            try
                            {
                                Debug.WriteLine(OR.Name);
                                RulesList.Add(OR);
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine(ex.Message);
                            }
                        }
                    }
                    catch
                    {
                        // error loading store
                    }
                }
            }
            catch (Exception ex)
            {
                //      MessageBox.Show(ex.Message);
            }
        }

        //This
        void Application_ContextMenuClose(Outlook.OlContextMenu ContextMenu)
        {
            selection = null;
            item = null;
            if (btnCheckForRule != null)
            {
                btnCheckForRule.Click -= new Office
                    ._CommandBarButtonEvents_ClickEventHandler(
                    btnRules_Click);
            }
            btnCheckForRule = null;

        }
        //
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //this
            Application.ItemContextMenuDisplay -= new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(Application_ItemContextMenuDisplay);
            Application.ContextMenuClose -= new Outlook.ApplicationEvents_11_ContextMenuCloseEventHandler(Application_ContextMenuClose);
            //
        }
        
        //This
        void Application_ItemContextMenuDisplay(Office.CommandBar CommandBar, Outlook.Selection Selection)
        {
            selection = Selection;
            if (GetMessageClass(selection[1])=="IPM.Note" &&selection.Count==1)
            {
                item = (Outlook.MailItem)selection[1];
                btnCheckForRule = (Office.CommandBarButton)CommandBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing);
                btnCheckForRule.Caption = "CheckRules";
                btnCheckForRule.Click += new Office._CommandBarButtonEvents_ClickEventHandler(btnCheckForRule_Click);
               
                foreach (Outlook.Rule R in Application.Session.DefaultStore.GetRules())
                {
                    Office.CommandBarButton test=     (Office.CommandBarButton)CommandBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing);
                    test.Caption = @"ADD2RULE: "+R.Name;
                    test.Click += new Office._CommandBarButtonEvents_ClickEventHandler(btnRules_Click);
                }  
            }
        }
        
        void btnRules_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            int rulecount = 1;
            int ruleindex = 1;
            string RuleName = Ctrl.Caption.Substring(10);//Load selected Rule Name from button caption
            bool AlreadyThere = false;
            Outlook.MailItem citem = Application.ActiveExplorer().Selection[1];//Load selected mail item
          
            Outlook.Rules MyRules = Application.Session.DefaultStore.GetRules();//Retrieve Rules
           
            Outlook.Folder Folder = Application.ActiveExplorer().CurrentFolder//Retrieve Folder
        as Outlook.Folder;
            string Email = citem.SenderEmailAddress;//Extract Selected Mail Item Email Address
            
            foreach (Outlook.Rule RL in MyRules)
            {
                
                if (RuleName == RL.Name)
                {
                    ruleindex = rulecount;//Assign indext to selected Rule
                }
                rulecount++;
            }

            foreach (Outlook.RuleCondition RC in MyRules[ruleindex].Conditions)
            {

                if (RC.Enabled) //Add selected item parts to respective conditions if condition enabled
                {
                    switch (RC.ConditionType)//When I put this in as switch condition vs automatically added all the case statements below!
                    {
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionAccount:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionAnyCategory:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionBody:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionBodyOrSubject:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionCategory:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionCc:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionDateRange:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionFlaggedForAction:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionFormName:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionFrom:

                            foreach (Outlook.Recipient ite in MyRules[ruleindex].Conditions.From.Recipients)
                            {
                                string rec = ite.Address;
                                string name = ite.Name;
                                if (rec == item.SenderEmailAddress || name==item.SenderName) AlreadyThere = true;

                            }


                            if (!AlreadyThere)
                            {
                                MyRules[ruleindex].Conditions.From.Recipients.Add(item.SenderEmailAddress);
                                MyRules[ruleindex].Conditions.From.Recipients.Add(item.SenderName);
                                MyRules[ruleindex].Conditions.From.Recipients.GetEnumerator();//This here thingy keeps the new address and rule condition from returning void and hince being added multiple times and not executing when rule is run. Otherwise even though the new condition recipient shows up in the wizard, but has no effect when wizard is run.
                                
                            }
                            else System.Windows.Forms.MessageBox.Show(item.SenderEmailAddress + @" already in sender list!",
                                   "Error Adding To Rule", System.Windows.Forms.MessageBoxButtons.OK);

                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionFromAnyRssFeed:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionFromRssFeed:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionHasAttachment:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionImportance:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionLocalMachineOnly:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionMeetingInviteOrUpdate:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionMessageHeader:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionNotTo:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionOOF:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionOnlyToMe:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionOtherMachine:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionProperty:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionRecipientAddress:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionSenderAddress:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionSenderInAddressBook:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionSensitivity:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionSentTo:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionSizeRange:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionSubject:
                            //Next to do#########################################
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionTo:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionToOrCc:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionUnknown:
                            break;
                        default:
                            break;
                    }
                    
                    
                }
            }
            
            
            MyRules.Save(false);
           
            
            MyRules[ruleindex].Execute(true, Folder, Type.Missing, Type.Missing);
            
        }

        void btnCheckForRule_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Outlook.MailItem sitem = Application.ActiveExplorer().Selection[1];//Load selected mail item
            Outlook.Rules TheRules = Application.Session.DefaultStore.GetRules();//Retrieve Rules
            string TheEmail = sitem.SenderEmailAddress;
            string TheName = sitem.SenderName;
            string TheDisplayName = sitem.SenderName;
            int recipientindex=1;
            int deleteADRindex = 0;
            bool deleteADRrecipient = false;
            int deleteNAMEindex = 0;
            bool deleteNAMErecipient = false;
            bool NoMatchFound = true;
            int RecipientCount = 0;
            foreach (Outlook.Rule RL in TheRules)
            {
                recipientindex = 1;
                RecipientCount = RL.Conditions.From.Recipients.Count;
                
                foreach (Outlook.Recipient s in RL.Conditions.From.Recipients)//Check by sender address or name
                {
                    if (s.Name == TheName)
                    {
                        DialogResult Namefound;
                        Namefound = System.Windows.Forms.MessageBox.Show(TheName + @" exist in Rule <" + RL.Name + @">" + Environment.NewLine + "Remove from Rule?",
                                    "Rule Found", System.Windows.Forms.MessageBoxButtons.YesNo);
                        NoMatchFound = false;
                        if (Namefound == DialogResult.Yes)
                        {
                            deleteNAMErecipient = true;
                            deleteNAMEindex = recipientindex;
                        }
                    }

                    if (s.Address == TheEmail)//skip if already being removed based on name
                    {
                        DialogResult ADRfound;
                       ADRfound=  System.Windows.Forms.MessageBox.Show(TheEmail + @" exist in Rule <"+RL.Name+@">"+Environment.NewLine+"Remove from Rule?",
                                   "Rule Found", System.Windows.Forms.MessageBoxButtons.YesNo);
                       NoMatchFound = false;
                       if (ADRfound == DialogResult.Yes)
                       {
                           deleteADRrecipient = true;
                           deleteADRindex = recipientindex;
                       }


                    }
                    
                    recipientindex++;
                }
                if (deleteADRrecipient)
                {
                    deleteADRindex = deleteADRindex - (RecipientCount - RL.Conditions.From.Recipients.Count);//Adjust index to account for any deletions made to Recipient list after index was set
                    RL.Conditions.From.Recipients.Remove(deleteADRindex);
                    RL.Conditions.From.Recipients.GetEnumerator();
                    deleteADRrecipient = false;
                    
                }
                if(deleteNAMErecipient)
                {
                    deleteNAMEindex = deleteNAMEindex - (RecipientCount - RL.Conditions.From.Recipients.Count);//Adjust index to account for any deletions made to Recipient list after index was set
                    RL.Conditions.From.Recipients.Remove(deleteNAMEindex);
                    RL.Conditions.From.Recipients.GetEnumerator();
                    deleteNAMErecipient = false;

                }
            }
            if (NoMatchFound) System.Windows.Forms.MessageBox.Show("Not in any Rules",
                                    "Rule Check", System.Windows.Forms.MessageBoxButtons.OK);
            else TheRules.Save(false);
        }
        private string GetMessageClass(object item)
        {
            object[] args = new Object[] { };
            Type t = item.GetType();
            return t.InvokeMember("messageClass", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.GetField | System.Reflection.BindingFlags.GetProperty, null, item, args).ToString();
        }

        // OnMyButtonClick routine handles all button click events
        // and displays IRibbonControl.Context in message box
        public void OnMyButtonClick(Office.IRibbonControl control)
        {
            string msg = string.Empty;
            if (control.Context is Outlook.AttachmentSelection)
            {
                msg = "Context=AttachmentSelection" + "\n";
                Outlook.AttachmentSelection attachSel =
                    control.Context as Outlook.AttachmentSelection;
                foreach (Outlook.Attachment attach in attachSel)
                {
                    msg = msg + attach.DisplayName + "\n";
                }
            }
            else if (control.Context is Outlook.Folder)
            {
                msg = "Context=Folder" + "\n";
                Outlook.Folder folder =
                    control.Context as Outlook.Folder;
                msg = msg + folder.Name;
            }
            else if (control.Context is Outlook.Selection)
            {
                msg = "Context=Selection" + "\n";
                Outlook.Selection selection =
                    control.Context as Outlook.Selection;
                if (selection.Count == 1)
                {
                    OutlookItem olItem =
                        new OutlookItem(selection[1]);
                    msg = msg + olItem.Subject
                        + "\n" + olItem.LastModificationTime;
                }
                else
                {
                    msg = msg + "Multiple Selection Count="
                        + selection.Count;
                }
            }
            else if (control.Context is Outlook.OutlookBarShortcut)
            {
                msg = "Context=OutlookBarShortcut" + "\n";
                Outlook.OutlookBarShortcut shortcut =
                    control.Context as Outlook.OutlookBarShortcut;
                msg = msg + shortcut.Name;
            }
            else if (control.Context is Outlook.Store)
            {
                msg = "Context=Store" + "\n";
                Outlook.Store store =
                    control.Context as Outlook.Store;
                msg = msg + store.DisplayName;
            }
            else if (control.Context is Outlook.View)
            {
                msg = "Context=View" + "\n";
                Outlook.View view =
                    control.Context as Outlook.View;
                msg = msg + view.Name;
            }
            else if (control.Context is Outlook.Inspector)
            {
                msg = "Context=Inspector" + "\n";
                Outlook.Inspector insp =
                    control.Context as Outlook.Inspector;
                if (insp.AttachmentSelection.Count >= 1)
                {
                    Outlook.AttachmentSelection attachSel =
                        insp.AttachmentSelection;
                    foreach (Outlook.Attachment attach in attachSel)
                    {
                        msg = msg + attach.DisplayName + "\n";
                    }
                }
                else
                {
                    OutlookItem olItem =
                        new OutlookItem(insp.CurrentItem);
                    msg = msg + olItem.Subject;
                }
            }
            else if (control.Context is Outlook.Explorer)
            {
                msg = "Context=Explorer" + "\n";
                Outlook.Explorer explorer =
                    control.Context as Outlook.Explorer;
                if (explorer.AttachmentSelection.Count >= 1)
                {
                    Outlook.AttachmentSelection attachSel =
                        explorer.AttachmentSelection;
                    foreach (Outlook.Attachment attach in attachSel)
                    {
                        msg = msg + attach.DisplayName + "\n";
                    }
                }
                else
                {
                    Outlook.Selection selection =
                        explorer.Selection;
                    if (selection.Count == 1)
                    {
                        OutlookItem olItem =
                            new OutlookItem(selection[1]);
                        msg = msg + olItem.Subject
                            + "\n" + olItem.LastModificationTime;
                    }
                    else
                    {
                        msg = msg + "Multiple Selection Count="
                            + selection.Count;
                    }
                }
            }
            else if (control.Context is Outlook.NavigationGroup)
            {
                msg = "Context=NavigationGroup" + "\n";
                Outlook.NavigationGroup navGroup =
                    control.Context as Outlook.NavigationGroup;
                msg = msg + navGroup.Name;
            }
            else if (control.Context is
                Microsoft.Office.Core.IMsoContactCard)
            {
                msg = "Context=IMsoContactCard" + "\n";
                Office.IMsoContactCard card =
                    control.Context as Office.IMsoContactCard;
                if (card.AddressType ==
                    Office.MsoContactCardAddressType.
                    msoContactCardAddressTypeOutlook)
                {
                    // IMSOContactCard.Address is AddressEntry.ID
                    Outlook.AddressEntry addr =
                        Globals.ThisAddIn.Application.Session.GetAddressEntryFromID(
                        card.Address);
                    if (addr != null)
                    {
                        msg = msg + addr.Name;
                    }
                }
            }
            else if (control.Context is Outlook.NavigationModule)
            {
                msg = "Context=NavigationModule";
            }
            else if (control.Context == null)
            {
                msg = "Context=Null";
            }
            else
            {
                msg = "Context=Unknown";
            }
            MessageBox.Show(msg,
                "RibbonXOutlook14AddinCS",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
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
