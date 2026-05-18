using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
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
            if (FindVisualParent<ButtonBase>(source) is not null ||
                FindVisualParent<TextBox>(source) is not null)
            {
                return;
            }

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

        private void PlanifierEssai_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel viewModel ||
                (sender as FrameworkElement)?.DataContext is not AgendaTaskItem task)
            {
                return;
            }

            viewModel.PlanAgendaTaskAtFirstAvailableSlot(task);
        }

        private void RetirerEssaiAgenda_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel viewModel)
            {
                return;
            }

            AgendaTaskItem? task = (sender as FrameworkElement)?.DataContext switch
            {
                AgendaTaskSegment segment => segment.SourceTask,
                AgendaTaskItem agendaTask => agendaTask,
                _ => null
            };

            if (task is not null)
            {
                viewModel.MoveAgendaTaskToBacklog(task);
                e.Handled = true;
            }
        }

        private void SegmentStartTime_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            BeginSegmentTimeEdit(sender, editStartTime: true, e);
        }

        private void SegmentEndTime_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            BeginSegmentTimeEdit(sender, editStartTime: false, e);
        }

        private void SegmentStartTimeEdit_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox textBox &&
                textBox.DataContext is AgendaTaskSegment segment &&
                segment.IsEditingStartTime)
            {
                CommitSegmentTimeEdit(textBox, editStartTime: true);
            }
        }

        private void SegmentEndTimeEdit_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox textBox &&
                textBox.DataContext is AgendaTaskSegment segment &&
                segment.IsEditingEndTime)
            {
                CommitSegmentTimeEdit(textBox, editStartTime: false);
            }
        }

        private void SegmentTimeEdit_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (sender is not TextBox textBox)
            {
                return;
            }

            if (e.Key == Key.Enter)
            {
                if (textBox.DataContext is AgendaTaskSegment segment)
                {
                    CommitSegmentTimeEdit(textBox, segment.IsEditingStartTime);
                    Keyboard.ClearFocus();
                    e.Handled = true;
                }

                return;
            }

            if (e.Key == Key.Escape)
            {
                CancelSegmentTimeEdit(textBox);
                Keyboard.ClearFocus();
                e.Handled = true;
            }
        }

        private void BeginSegmentTimeEdit(object sender, bool editStartTime, MouseButtonEventArgs e)
        {
            if (e.ClickCount < 2 ||
                sender is not FrameworkElement element ||
                element.DataContext is not AgendaTaskSegment segment ||
                !segment.CanEditTimes)
            {
                return;
            }

            segment.IsEditingStartTime = editStartTime;
            segment.IsEditingEndTime = !editStartTime;
            e.Handled = true;

            DependencyObject? editHost = FindVisualParent<Grid>(element);
            Dispatcher.BeginInvoke(new Action(() =>
            {
                TextBox? editor = FindVisualChild<TextBox>(editHost);
                if (editor is null)
                {
                    return;
                }

                editor.Focus();
                editor.SelectAll();
                Keyboard.Focus(editor);
            }), System.Windows.Threading.DispatcherPriority.Input);
        }

        private void CommitSegmentTimeEdit(TextBox textBox, bool editStartTime)
        {
            if (textBox.DataContext is not AgendaTaskSegment segment)
            {
                return;
            }

            segment.IsEditingStartTime = false;
            segment.IsEditingEndTime = false;

            bool updated = DataContext is MainViewModel viewModel &&
                (editStartTime
                    ? viewModel.TryUpdateAgendaTaskSegmentStartTime(segment, textBox.Text)
                    : viewModel.TryUpdateAgendaTaskSegmentEndTime(segment, textBox.Text));

            if (!updated)
            {
                textBox.Text = editStartTime ? segment.StartTimeText : segment.EndTimeText;
            }
        }

        private static void CancelSegmentTimeEdit(TextBox textBox)
        {
            if (textBox.DataContext is not AgendaTaskSegment segment)
            {
                return;
            }

            textBox.Text = segment.IsEditingStartTime ? segment.StartTimeText : segment.EndTimeText;
            segment.IsEditingStartTime = false;
            segment.IsEditingEndTime = false;
        }

        private void ClearAgendaSearch_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.ClearAgendaSearch();
            }
        }

        private void UndoAgenda_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel && viewModel.CanUndoAgenda)
            {
                viewModel.UndoAgendaLastAction();
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

        private static T? FindVisualParent<T>(DependencyObject? child) where T : DependencyObject
        {
            while (child is not null)
            {
                if (child is T match)
                {
                    return match;
                }

                child = System.Windows.Media.VisualTreeHelper.GetParent(child);
            }

            return null;
        }

        private static T? FindVisualChild<T>(DependencyObject? parent) where T : DependencyObject
        {
            if (parent is null)
            {
                return null;
            }

            int childCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                DependencyObject child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child is T match)
                {
                    return match;
                }

                T? nestedMatch = FindVisualChild<T>(child);
                if (nestedMatch is not null)
                {
                    return nestedMatch;
                }
            }

            return null;
        }
    }
}
