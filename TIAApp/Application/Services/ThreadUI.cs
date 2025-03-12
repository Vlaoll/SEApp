using System;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Controls;
using System.Windows.Media;


namespace seConfSW
{
    public class ThreadUI : DispatcherObject
    {
        public void messageText(string message, string color = "Black")
        {
            try
            {
                this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (Action)(() =>
                {
                    var mainWindow = (Application.Current.MainWindow as MainWindow);
                    //mainWindow.txtMessage.Document.Blocks.Clear();
                    //mainWindow.txtMessage.Foreground = new BrushConverter().ConvertFromString(color) as SolidColorBrush; 



                   //mainWindow.txtMessage.AppendText(message);                    
                }));
            }
            catch (ArgumentException)
            {
                MessageBox.Show("The exception forwarder from secondary UI thread!",
               "Dispatcher Error", MessageBoxButton.OK, MessageBoxImage.Stop);
                return;
            }
        }

        public void shutDown()
        {
            this.Dispatcher.InvokeShutdown();
        }
    }
}
