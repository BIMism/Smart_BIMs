using System.Collections.Generic;
using System.Windows;
using Autodesk.Revit.DB;

namespace Smart_BIMs.UI
{
    public partial class AddFieldWindow : Window
    {
        public List<FieldItem> AvailableFields { get; set; }
        public List<FieldItem> SelectedFields { get; set; } = new List<FieldItem>();
        private Document _doc;
        private ViewSchedule _schedule;

        public AddFieldWindow(Document doc, ViewSchedule schedule)
        {
            InitializeComponent();
            _doc = doc;
            _schedule = schedule;
            LoadFields();
        }

        private void LoadFields()
        {
            AvailableFields = new List<FieldItem>();
            ScheduleDefinition def = _schedule.Definition;
            
            HashSet<ElementId> existingIds = new HashSet<ElementId>();
            for (int i = 0; i < def.GetFieldCount(); i++) existingIds.Add(def.GetField(i).ParameterId);

            foreach (SchedulableField sf in def.GetSchedulableFields())
            {
                if (!existingIds.Contains(sf.ParameterId))
                {
                    AvailableFields.Add(new FieldItem {
                        Name = sf.GetName(_doc),
                        Schedulable = sf,
                        IsSelected = false,
                        IsInSchedule = false
                    });
                }
            }
            AvailableFields.Sort((a,b) => string.Compare(a.Name, b.Name));
            lstFields.ItemsSource = AvailableFields;
        }

        private void BtnPush_Click(object sender, RoutedEventArgs e)
        {
            foreach(var item in AvailableFields) {
                if (item.IsSelected) SelectedFields.Add(item);
            }
            if (SelectedFields.Count == 0) {
                MessageBox.Show("Please select at least one field to push.");
                return;
            }
            this.DialogResult = true;
            this.Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
