using System.Collections.Generic;
using System.Windows;
using Autodesk.Revit.DB;

namespace Smart_BIMs.UI
{
    public partial class EasyScheduleWindow : Window
    {
        private List<Category> _allCategories;

        public EasyScheduleWindow(List<Category> categories)
        {
            InitializeComponent();
            _allCategories = categories;
            CategoryListBox.ItemsSource = _allCategories;
        }

        public List<Category> GetSelectedCategories()
        {
            List<Category> selected = new List<Category>();
            foreach (var item in CategoryListBox.SelectedItems)
            {
                selected.Add((Category)item);
            }
            return selected;
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
