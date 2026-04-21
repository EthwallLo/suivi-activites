using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MonTableurApp.Models;
using MonTableurApp.ViewModels;

namespace MonTableurApp.Views
{
    public partial class VueAgendaView : UserControl
    {
        public VueAgendaView()
        {
            InitializeComponent();
        }

        private void AgendaList_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed)
            {
                return;
            }

            DependencyObject? source = e.OriginalSource as DependencyObject;
            while (source != null)
            {
                if (source is FrameworkElement segmentElement && segmentElement.DataContext is AgendaTaskSegment segment)
                {
                    DragDrop.DoDragDrop(segmentElement, segment.SourceTask, DragDropEffects.Move);
                    return;
                }

                if (source is FrameworkElement element && element.DataContext is AgendaTaskItem task)
                {
                    DragDrop.DoDragDrop(element, task, DragDropEffects.Move);
                    return;
                }

                source = System.Windows.Media.VisualTreeHelper.GetParent(source);
            }
        }

        private void DayPlan_Drop(object sender, DragEventArgs e)
        {
            if (DataContext is not MainViewModel viewModel)
            {
                return;
            }

            if (!e.Data.GetDataPresent(typeof(AgendaTaskItem)))
            {
                return;
            }

            if (sender is not FrameworkElement element || element.Tag is not AgendaWorkDay day)
            {
                return;
            }

            if (e.Data.GetData(typeof(AgendaTaskItem)) is AgendaTaskItem task)
            {
                Point position = e.GetPosition(element);
                viewModel.MoveAgendaTaskToDay(task, day, position.Y);
            }
        }

        private void Backlog_Drop(object sender, DragEventArgs e)
        {
            if (DataContext is not MainViewModel viewModel)
            {
                return;
            }

            if (!e.Data.GetDataPresent(typeof(AgendaTaskItem)))
            {
                return;
            }

            if (e.Data.GetData(typeof(AgendaTaskItem)) is AgendaTaskItem task)
            {
                viewModel.MoveAgendaTaskToBacklog(task);
            }
        }

        private void VueAgendaView_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Z || Keyboard.Modifiers != ModifierKeys.Control)
            {
                return;
            }

            if (e.OriginalSource is TextBox)
            {
                return;
            }

            if (DataContext is not MainViewModel viewModel || !viewModel.CanUndoAgenda)
            {
                return;
            }

            viewModel.UndoAgendaLastAction();
            e.Handled = true;
        }
    }
}
