using System.IO;

namespace SimpleFormUiApi
{
    public class SimpleForm
    {
        private SAPbouiCOM.Application SBO_Application;
        private SAPbouiCOM.Form oForm;

        public SimpleForm()
        {
            //*************************************************************
            // set SBO_Application with an initialized application object
            //*************************************************************

            SetApplication();

            //*************************************************************
            // Create the simple form
            //*************************************************************

            CreateMySimpleForm();

            oForm.Visible = true;

            //*************************************************************
            // Save the form to an Xml document
            //*************************************************************

            //SaveAsXML();
        }

        private void SetApplication()
        {

            SAPbouiCOM.SboGuiApi SboGuiApi = new SAPbouiCOM.SboGuiApi();

            var sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";

            //if (Environment.GetCommandLineArgs().Length > 1)
            //{
            //    sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            //}
            //else
            //{
            //    sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(0));
            //}

            try
            {
                SboGuiApi.Connect(sConnectionString);
            }
            catch
            {
                System.Environment.Exit(0);
            }

            SBO_Application = SboGuiApi.GetApplication(0);
        }


        private void CreateMySimpleForm()
        {
            SAPbouiCOM.Item oItem = null;

            // *******************************************
            // we will use the following objects to set
            // the specific values of every item
            // we add.
            // this is the best way to do so
            //*********************************************

            SAPbouiCOM.Button oButton = null;
            SAPbouiCOM.StaticText oStaticText = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oComboBox = null;

            // add a new form
            SAPbouiCOM.FormCreationParams oCreationParams = null;

            oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oCreationParams.UniqueID = "MySimpleForm";

            oForm = SBO_Application.Forms.AddEx(oCreationParams);

            // add a User Data Source to the form
            oForm.DataSources.UserDataSources.Add("EditSource", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("CombSource", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

            // set the form properties
            oForm.Title = "Simple Form";
            oForm.Left = 400;
            oForm.Top = 100;
            oForm.ClientHeight = 80;
            oForm.ClientWidth = 350;

            //*****************************************
            // Adding Items to the form
            // and setting their properties
            //*****************************************


            //**********************
            // Adding an Ok button
            //*********************

            // We get automatic event handling for
            // the Ok and Cancel Buttons by setting
            // their UIDs to 1 and 2 respectively

            oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 6;
            oItem.Width = 65;
            oItem.Top = 51;
            oItem.Height = 19;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Ok";

            //************************
            // Adding a Cancel button
            //***********************

            oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 75;
            oItem.Width = 65;
            oItem.Top = 51;
            oItem.Height = 19;

            oButton = ((SAPbouiCOM.Button)(oItem.Specific));

            oButton.Caption = "Cancel";

            //************************
            // Adding a Rectangle
            //***********************

            oItem = oForm.Items.Add("Rect1", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oItem.Left = 0;
            oItem.Width = 344;
            oItem.Top = 1;
            oItem.Height = 49;

            //***************************
            // Adding a Static Text item
            //***************************

            oItem = oForm.Items.Add("StaticTxt1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = 7;
            oItem.Width = 148;
            oItem.Top = 8;
            oItem.Height = 14;

            oItem.LinkTo = "EditText1";

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Static Text 1";

            //**********************************
            // Adding another Static Text item
            //**********************************

            oItem = oForm.Items.Add("StaticTxt2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = 7;
            oItem.Width = 148;
            oItem.Top = 24;
            oItem.Height = 14;

            oItem.LinkTo = "ComboBox1";

            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));

            oStaticText.Caption = "Static Text 2";

            //*************************
            // Adding a Text Edit item
            //*************************

            oItem = oForm.Items.Add("EditText1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = 157;
            oItem.Width = 163;
            oItem.Top = 8;
            oItem.Height = 14;

            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));

            // bind the text edit item to the defined used data source
            oEditText.DataBind.SetBound(true, "", "EditSource");

            oEditText.String = "Edit Text 1";

            //*************************
            // Adding a Combo Box item
            //*************************

            oItem = oForm.Items.Add("ComboBox1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oItem.Left = 157;
            oItem.Width = 163;
            oItem.Top = 24;
            oItem.Height = 14;

            oItem.DisplayDesc = false;

            oComboBox = ((SAPbouiCOM.ComboBox)(oItem.Specific));

            // bind the Combo Box item to the defined used data source
            oComboBox.DataBind.SetBound(true, "", "CombSource");

            oComboBox.ValidValues.Add("1", "Combo Value 1");
            oComboBox.ValidValues.Add("2", "Combo Value 2");
            oComboBox.ValidValues.Add("3", "Combo Value 3");


            //*************************
            // set the form as visible
            //*************************

        }


        private void SaveAsXML()
        {
            //**********************************************************************
            //
            // always use XML to work with user forms.
            // after creating your form save it as an XML file
            //
            //**********************************************************************

            System.Xml.XmlDocument oXmlDoc = null;

            oXmlDoc = new System.Xml.XmlDocument();

            string sXmlString = null;

            // get the form as an XML string
            sXmlString = oForm.GetAsXML();

            // load the form's XML string to the
            // XML document object
            oXmlDoc.LoadXml(sXmlString);

            // save the XML Document
            string sPath = null;

            var enderecoENomeDoArquivo = Path.GetTempFileName();
            sPath = System.IO.Directory.GetParent(enderecoENomeDoArquivo).ToString();

            oXmlDoc.Save((sPath + @"\MySimpleForm.xml"));

        }
    }
}
