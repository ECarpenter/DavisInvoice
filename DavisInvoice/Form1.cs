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
            vaultConnect = mFilesApp.GetVaultConnection(Properties.Settings.Default.VaultName);

            var currVault = new Vault();
            currVault = vaultConnect.BindToVault(this.Handle, true, false);

            //export
            button1.Click += delegate (object sender, EventArgs e) { button1_Click(sender, e, currVault, mFilesApp); };

            //pull up invoices

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        //Export Button
        private void button1_Click(object sender, EventArgs e, Vault currVault, MFilesClientApplication mFilesApp)
        {
            //MessageBox.Show(currVault.Name);

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
            foreach (ObjectVersion invoice in invoices)
            {
                XElement payable = new XElement("Payable");
                double totalAmount = 0;

                var objtype = default(ObjType);
                objtype = currVault.ObjectTypeOperations.GetObjectType(invoice.ObjVer.Type);
                MessageBox.Show(objtype.NameSingular);

                var propValues = new PropertyValues();
                var currPropertyValue = new PropertyValue();
                var objLedger = new ObjectVersion();
                propValues = currVault.ObjectPropertyOperations.GetProperties(invoice.ObjVer);

                XElement details = new XElement("Details");
                XElement detail = new XElement("Detail");

                //Get Ledger Entry reference
                currPropertyValue = propValues.SearchForProperty(Properties.Settings.Default.propLedgerEntry);
                if (currPropertyValue.TypedValue.DataType == MFDataType.MFDatatypeMultiSelectLookup) 
                {
                    var lookup = new Lookup();
                    lookup = currPropertyValue.TypedValue.GetValueAsLookup();

                    var propDef = new PropertyDef();
                    propDef = currVault.PropertyDefOperations.GetPropertyDef(currPropertyValue.PropertyDef);
                    var valListObjType = new ObjType();
                    valListObjType = currVault.ValueListOperations.GetValueList(propDef.ValueList);

                    if (valListObjType.RealObjectType)
                    {
                        //Get Ledgery Entry Object 
                        var objDetail = new ObjVer();
                        objDetail.SetIDs(valListObjType.ID, lookup.Item, lookup.Version);
                        var detailValues = new PropertyValues();
                        var detailValue = new PropertyValue();
                        detailValues = currVault.ObjectPropertyOperations.GetProperties(objDetail);

                        //Get Account
                        detailValue = detailValues.SearchForProperty(Properties.Settings.Default.propAccount);
                        if (detailValue.TypedValue.DataType == MFDataType.MFDatatypeMultiSelectLookup)
                        {
                            lookup = detailValue.TypedValue.GetValueAsLookup();

                            propDef = currVault.PropertyDefOperations.GetPropertyDef(detailValue.PropertyDef);
                            valListObjType = currVault.ValueListOperations.GetValueList(propDef.ValueList);

                            if (valListObjType.RealObjectType)
                            {
                                //Get Account Number 
                                var objAccount = new ObjVer();
                                objAccount.SetIDs(valListObjType.ID, lookup.Item, lookup.Version);
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
                        payable.Add(notes);

                        //get Amount
                        detailValue = detailValues.SearchForProperty(Properties.Settings.Default.propGLAmount);
                        XElement amount = new XElement("Amount");
                        amount.SetValue(detailValue.GetValueAsLocalizedText());
                        detail.Add(amount);
                        totalAmount += Convert.ToDouble(detailValue.GetValueAsLocalizedText());

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
                        XElement propertyID = new XElement("PropertyId");
                        propertyID.SetValue(propertyValue.GetValueAsLocalizedText());
                        detail.Add(propertyID);
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
                details.Add(detail);
                payable.Add(details);

                //Add Post Month
                XElement postMonth = new XElement("PostMonth");
                postMonth.SetValue(postMonthForm.StrPostMonth);
                payable.Add(postMonth);

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


            }

            
             output.Add(payables);
             output.Save("ThirdImport.XML");


            //End Loop

            //combine xaml objects

            //save export file

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        //import button
        private void button2_Click(object sender, EventArgs e)
        {
            //open XML file

            //loop through items

            //find mFiles object with matching invoice number

            //add check number to mFiles Object

            //change state to complete

            //End Loop
        }
    }
}
