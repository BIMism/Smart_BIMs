using System;
using System.Collections.Generic;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Smart_BIMs.UI
{
    public partial class ExcelManagerWindow : Window
    {
        private Document _doc;
        private ViewSchedule _schedule;
        public List<FieldItem> AvailableFields { get; set; }
        
        public bool DoExport { get; private set; }
        public bool DoSync { get; private set; }

        public ExcelManagerWindow(Document doc, ViewSchedule schedule)
        {
            InitializeComponent();
            _doc = doc;
            _schedule = schedule;
            txtTitle.Text = "Manage: " + schedule.Name;

            LoadFields();
        }

        private void LoadFields()
        {
            AvailableFields = new List<FieldItem>();
            ScheduleDefinition def = _schedule.Definition;
            
            HashSet<ElementId> existingIds = new HashSet<ElementId>();
            for (int i = 0; i < def.GetFieldCount(); i++)
            {
                existingIds.Add(def.GetField(i).ParameterId);
            }

            IList<SchedulableField> allFields = def.GetSchedulableFields();
            foreach (SchedulableField sf in allFields)
            {
                bool isIn = existingIds.Contains(sf.ParameterId);
                AvailableFields.Add(new FieldItem {
                    Name = sf.GetName(_doc),
                    Schedulable = sf,
                    IsSelected = isIn,
                    IsInSchedule = isIn
                });
            }

            // Put existing headers first
            AvailableFields.Sort((a, b) => b.IsInSchedule.CompareTo(a.IsInSchedule));

            lstFields.ItemsSource = AvailableFields;
        }

        private void btnAddNewParam_Click(object sender, RoutedEventArgs e)
        {
            string pName = txtNewParamName.Text.Trim();
            if (string.IsNullOrEmpty(pName)) { MessageBox.Show("Please enter a parameter name."); return; }
            int pTypeIdx = cmbNewParamType.SelectedIndex;

            try
            {
                using (Transaction trans = new Transaction(_doc, "Create Shared Parameter"))
                {
                    trans.Start();
                    Autodesk.Revit.ApplicationServices.Application app = _doc.Application;
                    string spFile = app.SharedParametersFilename;
                    if (string.IsNullOrEmpty(spFile) || !System.IO.File.Exists(spFile))
                    {
                        string tempSP = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "SmartBIM_SharedParams.txt");
                        string content = "# This is a Revit shared parameter file.\r\n*META\tVERSION\tMINVERSION\r\nMETA\t2\t1\r\n*GROUP\tID\tNAME\r\n*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\tHIDEWHENNOVALUE\r\n";
                        System.IO.File.WriteAllText(tempSP, content);
                        app.SharedParametersFilename = tempSP;
                    }
                    DefinitionFile defFile = app.OpenSharedParameterFile();
                    if (defFile == null) { MessageBox.Show("Could not initialize Shared Parameters file."); trans.RollBack(); return; }

                    DefinitionGroup grp = defFile.Groups.get_Item("SmartBIM Data") ?? defFile.Groups.Create("SmartBIM Data");

                    Definition existingDef = grp.Definitions.get_Item(pName);
                    if (existingDef == null)
                    {
                        Autodesk.Revit.DB.ForgeTypeId dataType = Autodesk.Revit.DB.SpecTypeId.String.Text;
                        if (pTypeIdx == 1) dataType = Autodesk.Revit.DB.SpecTypeId.Number;
                        else if (pTypeIdx == 2) dataType = Autodesk.Revit.DB.SpecTypeId.Int.Integer;
                        else if (pTypeIdx == 3) dataType = Autodesk.Revit.DB.SpecTypeId.Length;
                        Autodesk.Revit.DB.ExternalDefinitionCreationOptions opt = new Autodesk.Revit.DB.ExternalDefinitionCreationOptions(pName, dataType);
                        existingDef = grp.Definitions.Create(opt);
                    }

                    CategorySet catSet = app.Create.NewCategorySet();
                    if (_schedule.Definition.CategoryId != ElementId.InvalidElementId)
                    {
                        catSet.Insert(Category.GetCategory(_doc, _schedule.Definition.CategoryId));
                    }
                    else
                    {
                        var elems = new FilteredElementCollector(_doc, _schedule.Id).ToElements();
                        foreach(var el in elems) { if (el.Category != null) catSet.Insert(el.Category); }
                    }

                    if (catSet.IsEmpty) { MessageBox.Show("Could not resolve categories for the schedule."); trans.RollBack(); return; }

                    InstanceBinding binding = app.Create.NewInstanceBinding(catSet);
                    _doc.ParameterBindings.Insert(existingDef, binding, Autodesk.Revit.DB.GroupTypeId.Data);
                    
                    trans.Commit();
                }

                LoadFields();
                foreach(var fi in AvailableFields) { if (fi.Name == pName) fi.IsSelected = true; }
                lstFields.ItemsSource = null;
                lstFields.ItemsSource = AvailableFields;
                
                txtNewParamName.Text = "";
                MessageBox.Show($"Parameter '{pName}' successfully added to the Model!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex) { MessageBox.Show("Failed to create parameter: " + ex.Message); }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            DoExport = true;
            this.Close();
        }

        private void BtnSync_Click(object sender, RoutedEventArgs e)
        {
            DoSync = true;
            this.Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }

    public class FieldItem
    {
        public string Name { get; set; }
        public SchedulableField Schedulable { get; set; }
        public bool IsSelected { get; set; }
        public bool IsInSchedule { get; set; }
    }
}
