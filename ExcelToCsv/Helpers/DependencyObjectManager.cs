using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media;

namespace ExcelToCsv.Helpers
{
    public class DependencyObjectManager
    {
        public static List<DependencyObject> GetChildren(DependencyObject reference)
        {
            if (reference == null)
                return new List<DependencyObject>();

            var children = new List<DependencyObject>();
            var count = VisualTreeHelper.GetChildrenCount(reference);

            for (var i = 0; i < count; i++)
            {
                children.Add(VisualTreeHelper.GetChild(reference, i));
            }

            return children;
        }

        public static DependencyObject GetChildByName(DependencyObject reference, string name)
        {
            var queue = GetQueue(reference);

            while (queue.Any())
            {
                var child = queue.Dequeue();

                if (((FrameworkElement)child).Name == name)
                    return child;

                AddChildrenOfChildToQueue(child, queue);
            }

            return null;
        }

        public static DependencyObject GetFirstChildByType(DependencyObject reference, Type type)
        {
            var queue = GetQueue(reference);

            while (queue.Any())
            {
                var child = queue.Dequeue();

                if (((FrameworkElement)child).GetType() == type)
                    return child;

                AddChildrenOfChildToQueue(child, queue);
            }

            return null;
        }

        private static Queue<DependencyObject> GetQueue(DependencyObject reference) => new Queue<DependencyObject>(GetChildren(reference));

        private static void AddChildrenOfChildToQueue(DependencyObject child, Queue<DependencyObject> queue)
        {
            var childrenOfChild = GetChildren(child);

            foreach (var childOfChild in childrenOfChild)
                queue.Enqueue(childOfChild);
        }
    }
}