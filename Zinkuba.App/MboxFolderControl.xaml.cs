using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using Rendezz.UI;
using Zinkuba.App.Annotations;
using Zinkuba.App.Folder;
using Zinkuba.App.Mailbox;

namespace Zinkuba.App
{
    /// <summary>
    /// Interaction logic for Mbox.xaml
    /// </summary>
    public partial class MboxFolderControl : UserControl, IReflectedObject<MboxFolder>
    {
        private MboxDataContext _dataContext;
        private MboxFolder _mboxFolder;
        public Action<MboxFolder> RemoveFolderFunction;

        public MboxFolderControl(MboxFolder mboxFolder)
        {
            InitializeComponent();
            _mboxFolder = mboxFolder;
            _dataContext = new MboxDataContext(_mboxFolder);
            DataContext = _dataContext;
        }

        private void RemoveItem(object sender, RoutedEventArgs e)
        {
            if (RemoveFolderFunction != null)
            {
                RemoveFolderFunction(_mboxFolder);
            }
        }

        public MboxFolder MirrorSource { get { return _mboxFolder; } }

        private void BrowseForMbox(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            //dlg.DefaultExt = ".mbox";
            //dlg.Filter = "Mbox Files (*.mbox;mbox)|";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                SetMboxFile(filename);
            }
        }

        public void SetMboxFile(string filename)
        {
            _dataContext.MboxPath = filename;
            _dataContext.FolderPath = Regex.Match(filename, @".*\\(.*)\.mbox").Groups[1].Value;
        }
    }

    public class MboxDataContext : INotifyPropertyChanged
    {
        private MboxFolder _folder;

        public MboxDataContext(MboxFolder folder)
        {
            _folder = folder;
        }

        public String MboxPath { get { return _folder.MboxPath; } set { _folder.MboxPath = value; OnPropertyChanged("MboxPath"); }}
        public String FolderPath { get { return _folder.FolderPath; } set { _folder.FolderPath = value; OnPropertyChanged("FolderPath");} }


        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName = null)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
