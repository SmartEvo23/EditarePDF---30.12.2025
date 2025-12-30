using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace EditarePDF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenPdf_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Add logic to open a PDF file
        }

        private void Eraser_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement eraser functionality here
        }

        private void ConvertToWord_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement PDF to Word conversion logic here
            MessageBox.Show("Convertire Word funcționalitate nu este încă implementată.");
        }

        private void EditorCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            // TODO: Add your mouse down logic here
        }

        private void EditorCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            // TODO: Add your mouse move logic here
        }
    }
}