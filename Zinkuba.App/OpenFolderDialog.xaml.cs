using System.Windows;

namespace Zinkuba.App
{
    /// <summary>
    /// Interaction logic for OpenFolderDialog.xaml
    /// </summary>
    public partial class OpenFolderDialog : Window
    {
        public OpenFolderDialog()
        {
            InitializeComponent();
        }

        public string Folder { get; private set; }

        private void Submit(object sender, RoutedEventArgs e)
        {
            Folder = FolderPath.Text;
            Close();
        }
    }
}
