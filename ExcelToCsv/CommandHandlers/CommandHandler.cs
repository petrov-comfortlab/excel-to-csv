using System;
using System.Windows.Input;

namespace ExcelToCsv.CommandHandlers
{
    public class CommandHandler : ICommand
    {
        public Action CommandAction { set; get; }
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
            CommandAction();
        }
    }
}