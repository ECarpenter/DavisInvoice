using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Windows.Forms;
using MFilesAPI;

namespace DavisInvoice
{
    public partial class DavisInvoice : Form
    {


        public DavisInvoice()
        {
            InitializeComponent();
            //connect to mFiles
            var mFilesApp = new MFilesClientApplication();
            

            //open vault
            var vaultConnect = new VaultConnection();
            vaultConnect = mFilesApp.GetVaultConnectionsWithGUID("{" + Properties.Settings.Default.vaultGUID + "}").Cast<VaultConnection>().FirstOrDefault();

            var currVault = new Vault();
            currVault = vaultConnect.BindToVault(this.Handle, true, false);

            //export
            button1.Click += delegate (object sender, EventArgs e) { button1_Click(sender, e, currVault, mFilesApp); };
            //import
            button2.Click += delegate (object sender, EventArgs e) { button2_Click(sender, e, currVault, mFilesApp); };
            

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        //Export Button
        private void button1_Click(object sender, EventArgs e, Vault currVault, MFilesClientApplication mFilesApp)
        {
            //MessageBox.Show(currVault.Name);

            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "XML Files (*.xml)|*.xml";


            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                //pull invoices using search conditions
                var searchConditions = new SearchConditions();

                //is it not deleted
                var isNotDeleted = new SearchCondition();
                isNotDeleted.Expression.DataStatusValueType = MFStatusType.MFStatusTypeDeleted;
                isNotDeleted.Expression.DataStatusValueDataFunction = MFDataFunction.MFDataFunctionNoOp;
                isNotDeleted.ConditionType = MFConditionType.MFConditionTypeNotEqual;
                isNotDeleted.TypedValue.SetValue(MFDataType.MFDatatypeBoolean, true);
                searchConditions.Add(-1, isNotDeleted);

                //is it part of the Invoice workflow
                var isInvoice = new SearchCondition();
                isInvoice.Expression.DataPropertyValuePropertyDef = (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefWorkflow;
                isInvoice.ConditionType = MFConditionType.MFConditionTypeEqual;
                isInvoice.TypedValue.SetValue(MFDataType.MFDatatypeLookup, Properties.Settings.Default.invoiceWorkflow);
                searchConditions.Add(-1, isInvoice);

                //is it in the accounting state
                var isAccounting = new SearchCondition();
                isAccounting.Expression.DataPropertyValuePropertyDef = (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefState;
                isAccounting.ConditionType = MFConditionType.MFConditionTypeEqual;
                isAccounting.TypedValue.SetValue(MFDataType.MFDatatypeLookup, Properties.Settings.Default.stateAccounting);
                searchConditions.Add(-1, isAccounting);

                //Perform search
                var invoices = currVault.ObjectSearchOperations.SearchForObjectsByConditions(searchConditions, MFSearchFlags.MFSearchFlagNone, false);


                //start output file
                XElement output = new XElement("YsiTran");
                XElement payables = new XElement("Payables");

                //get post month 
                var postMonthForm = new PostMonth();
                postMonthForm.ShowDialog();

                //loop through invoices collecting at import workflow state, build XML file from the inside out. 
                int count = 0;
                foreach (ObjectVersion invoice in invoices)
                {
                    XElement payable = new XElement("Payable");
                    double totalAmount = 0;
                    count++;

                    var propValues = new PropertyValues();
                    var currPropertyValue = new PropertyValue();
                    var objLedger = new ObjectVersion();
                    propValues = currVault.ObjectPropertyOperations.GetProperties(invoice.ObjVer);

                    XElement details = new XElement("Details");

                    //Get Ledger Entry reference
                    currPropertyValue = propValues.SearchForProperty(Properties.Settings.Default.propLedgerEntry);
                    
                    if (currPropertyValue.TypedValue.DataType == MFDataType.MFDatatypeMultiSelectLookup)
                    {
                        var lookups = new Lookups();

                        lookups = currPropertyValue.TypedValue.GetValueAsLookups();

                        int i = 0;
                        foreach (Lookup lookup in lookups)
                        {
                            XElement detail = new XElement("Detail");

                            var propDef = new PropertyDef();
                            propDef = currVault.PropertyDefOperations.GetPropertyDef(currPropertyValue.PropertyDef);
                            var valListObjType = new ObjType();
                            valListObjType = currVault.ValueListOperations.GetValueList(propDef.ValueList);
                        
                            if (valListObjType.RealObjectType)
                            {
                                i++;
                                //Get Ledgery Entry Object 
                                var objDetail = new ObjVer();
                                objDetail.SetIDs(valListObjType.ID, lookup.Item, lookup.Version);
                                var detailValues = new PropertyValues();
                                var detailValue = new PropertyValue();
                                detailValues = currVault.ObjectPropertyOperations.GetProperties(objDetail);
                                //MessageBox.Show(i.ToString());
                                //Get Account
                                detailValue = detailValues.SearchForProperty(Properties.Settings.Default.propAccount);
                                if (detailValue.TypedValue.DataType == MFDataType.MFDatatypeMultiSelectLookup)
                                {
                                    Lookup lookupAccount = new Lookup();
                                    lookupAccount = detailValue.TypedValue.GetValueAsLookup();

                                    propDef = currVault.PropertyDefOperations.GetPropertyDef(detailValue.PropertyDef);
                                    valListObjType = currVault.ValueListOperations.GetValueList(propDef.ValueList);

                                    if (valListObjType.RealObjectType)
                                    {
                                        //Get Account Number 
                                        var objAccount = new ObjVer();
                                        objAccount.SetIDs(valListObjType.ID, lookupAccount.Item, lookupAccount.Version);
                                        var accountValues = new PropertyValues();
                                        var accountValue = new PropertyValue();
                                        accountValues = currVault.ObjectPropertyOperations.GetProperties(objAccount);
                                        accountValue = accountValues.SearchForProperty(Properties.Settings.Default.propGLCode);
                                        XElement account = new XElement("AccountId");
                                        account.SetValue(accountValue.GetValueAsLocalizedText());
                                        detail.Add(account);
                                    }
                                }

                                //get Description-Notes
                                detailValue = detailValues.SearchForProperty(Properties.Settings.Default.propDescription);
                                XElement notes = new XElement("Notes");
                                notes.SetValue(detailValue.GetValueAsLocalizedText());
                                detail.Add(notes);

                                //get Amount
                                detailValue = detailValues.SearchForProperty(Properties.Settings.Default.propGLAmount);
                                XElement amount = new XElement("Amount");
                                amount.SetValue(detailValue.GetValueAsLocalizedText());
                                detail.Add(amount);
                                totalAmount += Convert.ToDouble(detailValue.GetValueAsLocalizedText());
                            }

                            XElement propertyID = new XElement("PropertyId");
                            detail.Add(propertyID);
                            details.Add(detail);
                            
                        }
                    }

                    //Get Property ID
                    currPropertyValue = propValues.SearchForProperty(Properties.Settings.Default.propProperty);
                    if (currPropertyValue.TypedValue.DataType == MFDataType.MFDatatypeMultiSelectLookup)
                    {
                        //Getlookup of property to find object
                        var lookup = new Lookup();
                        lookup = currPropertyValue.TypedValue.GetValueAsLookup();
                        var propDef = new PropertyDef();
                        propDef = currVault.PropertyDefOperations.GetPropertyDef(currPropertyValue.PropertyDef);
                        var valListObjType = new ObjType();
                        valListObjType = currVault.ValueListOperations.GetValueList(propDef.ValueList);

                        if (valListObjType.RealObjectType)
                        {
                            //Get property ID 
                            var objProperty = new ObjVer();
                            objProperty.SetIDs(valListObjType.ID, lookup.Item, lookup.Version);
                            var propertyValues = new PropertyValues();
                            var propertyValue = new PropertyValue();
                            propertyValues = currVault.ObjectPropertyOperations.GetProperties(objProperty);
                            propertyValue = propertyValues.SearchForProperty(Properties.Settings.Default.propPropertyID);
                            
                            

                            IEnumerable < XElement > ieDetails = from el in details.Elements() select el;


                            //loop through items

                            foreach (XElement detail in ieDetails)
                            {
                                //Check that a check has been cut in Yardi.
                                if (detail.Elements("PropertyId").Any())
                                {
                                    detail.Element("PropertyId").SetValue(propertyValue.GetValueAsLocalizedText());
                                }
                            }
                        }
                    }

                    //Get Vendor ID
                    currPropertyValue = propValues.SearchForProperty(Properties.Settings.Default.propVendor);
                    if (currPropertyValue.TypedValue.DataType == MFDataType.MFDatatypeMultiSelectLookup)
                    {
                        //Getlookup of vendor to find object
                        var lookup = new Lookup();
                        lookup = currPropertyValue.TypedValue.GetValueAsLookup();
                        var propDef = new PropertyDef();
                        propDef = currVault.PropertyDefOperations.GetPropertyDef(currPropertyValue.PropertyDef);
                        var valListObjType = new ObjType();
                        valListObjType = currVault.ValueListOperations.GetValueList(propDef.ValueList);

                        if (valListObjType.RealObjectType)
                        {
                            //Get Vendor ID 
                            var objProperty = new ObjVer();
                            objProperty.SetIDs(valListObjType.ID, lookup.Item, lookup.Version);
                            var vendorValues = new PropertyValues();
                            var vendorValue = new PropertyValue();
                            vendorValues = currVault.ObjectPropertyOperations.GetProperties(objProperty);
                            vendorValue = vendorValues.SearchForProperty(Properties.Settings.Default.propYardiCode);
                            XElement propertyID = new XElement("PersonId");
                            propertyID.SetValue(vendorValue.GetValueAsLocalizedText());
                            payable.Add(propertyID);
                        }
                    }

                    // Add details to payable
                    payable.Add(details);

                    //Add Post Month
                    XElement postMonth = new XElement("PostMonth");
                    postMonth.SetValue(postMonthForm.StrPostMonth);
                    payable.Add(postMonth);

                    //Get link to object
                    XElement link = new XElement("Notes");
                    string strLink = "https://openurl.m-files.com/2.0/OpenMFilesUrl.html?MFilesURL=m-files://show/" + Properties.Settings.Default.vaultGUID + "/" + invoice.ObjVer.Type.ToString() + "-" + invoice.ObjVer.ID.ToString();
                    link.SetValue(strLink);
                    payable.Add(link);

                    //get Invoice Number
                    currPropertyValue = propValues.SearchForProperty(Properties.Settings.Default.propInvoiceNumber);
                    XElement invoiceNumber = new XElement("InvoiceNumber");
                    invoiceNumber.SetValue(currPropertyValue.GetValueAsLocalizedText());
                    payable.Add(invoiceNumber);

                    //get Invoice date
                    currPropertyValue = propValues.SearchForProperty(Properties.Settings.Default.propInvoiceDate);
                    XElement invoiceDate = new XElement("InvoiceDate");
                    invoiceDate.SetValue(currPropertyValue.GetValueAsLocalizedText());
                    payable.Add(invoiceDate);

                    //get Due Date
                    currPropertyValue = propValues.SearchForProperty(Properties.Settings.Default.propDueDate);
                    XElement dueDate = new XElement("DueDate");
                    dueDate.SetValue(currPropertyValue.GetValueAsLocalizedText());
                    payable.Add(dueDate);

                    //Set Total
                    XElement total = new XElement("TotalAmount");
                    total.SetValue(totalAmount.ToString());


                    payables.Add(payable);

                    //change workflow state 
                    propValues.SearchForProperty((int)MFBuiltInPropertyDef.MFBuiltInPropertyDefState).TypedValue.SetValue(MFDataType.MFDatatypeLookup, Properties.Settings.Default.stateProcessing);
                    currVault.ObjectPropertyOperations.SetAllProperties(invoice.ObjVer, true, propValues);

                }



                output.Add(payables);
                output.Save(saveFile.FileName);
                MessageBox.Show(count.ToString() + " Files Exported!");
            }
        }
        


    

        private void button1_Click(object sender, EventArgs e)
        {

        }

        //import button
        private void button2_Click(object sender, EventArgs e)
        {
            
        }
        private void button2_Click(object sender, EventArgs e, Vault currVault, MFilesClientApplication mFilesApp)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter ="XML Files (*.xml) | *.xml";

            if (openFile.ShowDialog() == DialogResult.OK)
            {

                //open XML file
                XElement inputFile = XElement.Load(openFile.FileName);
                IEnumerable<XElement> inputxml = from el in inputFile.Element("Payables").Elements() select el;


                //loop through items
                int count = 0;
                foreach (XElement payable in inputxml)
                {
                    //Check that a check has been cut in Yardi.
                    if (payable.Element("Details").Element("Detail").Elements("CheckNumber").Any())
                    {
                        string checkNumber = payable.Element("Details").Element("Detail").Element("CheckNumber").Value;
                        //MessageBox.Show(checkNumber);

                        //Find Invoice in mFiles
                        //pull invoices using search conditions
                        var searchConditions = new SearchConditions();

                        //is it not deleted
                        var isNotDeleted = new SearchCondition();
                        isNotDeleted.Expression.DataStatusValueType = MFStatusType.MFStatusTypeDeleted;
                        isNotDeleted.Expression.DataStatusValueDataFunction = MFDataFunction.MFDataFunctionNoOp;
                        isNotDeleted.ConditionType = MFConditionType.MFConditionTypeNotEqual;
                        isNotDeleted.TypedValue.SetValue(MFDataType.MFDatatypeBoolean, true);
                        searchConditions.Add(-1, isNotDeleted);

                        //is it part of the Invoice workflow
                        var isInvoice = new SearchCondition();
                        isInvoice.Expression.DataPropertyValuePropertyDef = (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefWorkflow;
                        isInvoice.ConditionType = MFConditionType.MFConditionTypeEqual;
                        isInvoice.TypedValue.SetValue(MFDataType.MFDatatypeLookup, Properties.Settings.Default.invoiceWorkflow);
                        searchConditions.Add(-1, isInvoice);

                        //is it in the payment processing state
                        var isAccounting = new SearchCondition();
                        isAccounting.Expression.DataPropertyValuePropertyDef = (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefState;
                        isAccounting.ConditionType = MFConditionType.MFConditionTypeEqual;
                        isAccounting.TypedValue.SetValue(MFDataType.MFDatatypeLookup, Properties.Settings.Default.stateProcessing);
                        searchConditions.Add(-1, isAccounting);

                        //is it the correct payable
                        var isPayable = new SearchCondition();
                        isPayable.Expression.DataPropertyValuePropertyDef = Properties.Settings.Default.propInvoiceNumber;
                        isPayable.ConditionType = MFConditionType.MFConditionTypeEqual;
                        isPayable.TypedValue.SetValue(MFDataType.MFDatatypeText, payable.Element("InvoiceNumber").Value);
                        searchConditions.Add(-1, isPayable);

                        //Perform search
                        var invoices = currVault.ObjectSearchOperations.SearchForObjectsByConditions(searchConditions, MFSearchFlags.MFSearchFlagNone, false);

                        foreach (ObjectVersion invoice in invoices)
                        {
                            var propValues = new PropertyValues();
                            var currPropertyValue = new PropertyValue();
                            propValues = currVault.ObjectPropertyOperations.GetProperties(invoice.ObjVer);

                            //currPropertyValue = propValues.SearchForProperty(Properties.Settings.Default.propCheckNumber);
                            count++;
                            propValues.SearchForProperty(Properties.Settings.Default.propCheckNumber).TypedValue.SetValue(MFDataType.MFDatatypeText, checkNumber);
                            //change workflow state 
                            propValues.SearchForProperty((int)MFBuiltInPropertyDef.MFBuiltInPropertyDefState).TypedValue.SetValue(MFDataType.MFDatatypeLookup, Properties.Settings.Default.stateComplete);
                            currVault.ObjectPropertyOperations.SetAllProperties(invoice.ObjVer, true, propValues);
                        }
                    }

                }
                MessageBox.Show(count.ToString() + " Invoices update in mFiles!");
            }
        }
    }
}
