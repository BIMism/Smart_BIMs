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
