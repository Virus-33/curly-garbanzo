using System.Windows.Input;
using System;

namespace Logic.Utility
{
    public class Command : ICommand
    {        
        Action<object> act;

        Func<object, bool> canAct;

        public Command(Action<object> act, Func<object, bool> canAct = null)
        {
            this.act = act;
            this.canAct = canAct;
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return this.canAct == null || this.canAct(parameter);
        }

        public void Execute(object parameter)
        {
            this.act(parameter);
        }
    }
}
