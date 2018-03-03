using System.Windows.Input;

namespace QueryRunner
{
    public static class ViewCommands
    {
        public static readonly RoutedUICommand Exit = new RoutedUICommand
            (
                text: "E_xit",
                name: "Exit",
                ownerType: typeof(ViewCommands),
                inputGestures: new InputGestureCollection()
                {
                    new KeyGesture(Key.X, ModifierKeys.Alt)
                }
            );

        public static readonly RoutedUICommand Close = new RoutedUICommand
            (
                text: "_Close",
                name: "Close",
                ownerType: typeof(ViewCommands),
                inputGestures: new InputGestureCollection()
                {
                    new KeyGesture(Key.C, ModifierKeys.Alt)
                }
            );
    }
}
