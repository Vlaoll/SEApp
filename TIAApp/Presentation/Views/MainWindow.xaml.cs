using seConfSW.Presentation.ViewModels;
using System.Windows;

namespace seConfSW.Presentation.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow(MainWindowViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
            Closed += (s, e) => viewModel.CloseWindows(); // Вызов метода закрытия при закрытии окна
        }
    }
}