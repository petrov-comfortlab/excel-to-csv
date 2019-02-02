using System;
using System.Windows;
using System.Windows.Input;

namespace ExcelToCsv.CommandHandlers
{
    public class CommandParameterHandler : ICommand
    {
        public Action<object> CommandAction { set; get; }
        public Func<bool> CanExecuteFunc { set; get; }

        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }

        public bool CanExecute(object parameter)
        {
            return CanExecuteFunc == null || CanExecuteFunc();
        }

        public void Execute(object parameter)
        {
            try
            {
                CommandAction(parameter);
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}