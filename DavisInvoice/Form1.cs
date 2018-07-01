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
                var objtype = default(ObjType);
                objtype = currVault.ObjectTypeOperations.GetObjectType(invoice.ObjVer.Type);
                MessageBox.Show(objtype.NameSingular);

                var propValues = new PropertyValues();
                var currPropertyValue = new PropertyValue();
                var objLedger = new ObjectVersion();
                propValues = currVault.ObjectPropertyOperations.GetProperties(invoice.ObjVer);

                //Get Ledger Entry reference
                XElement details = new XElement("Details");
                XElement detail = new XElement("Detail");
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
                                
                            }


                        }

                    }


                }


                MessageBox.Show(currPropertyValue.GetValueAsLocalizedText());
                //Start Detail
                XElement Detail = new XElement("Detail");
                

                XElement payable = new XElement("Payable");
                payable.SetElementValue("PostMonth", postMonthForm.StrPostMonth);
                payables.Add(payable);


            }

            

            //add data to xml object

            //change workflow state 

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
