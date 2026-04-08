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
            for (int i = 0; i < def.GetFieldCount(); i++)
            {
                ScheduleField field = def.GetField(i);
                AvailableFields.Add(new FieldItem { Name = field.GetName(), Field = field, IsSelected = true });
            }
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
        public ScheduleField Field { get; set; }
        public bool IsSelected { get; set; }
    }
}
