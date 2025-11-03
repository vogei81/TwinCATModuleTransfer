using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using TCatSysManagerLib;
using TwinCATModuleTransfer.ViewModels;

namespace TwinCATModuleTransfer.Services
{
    public static class ProjectTreeService
    {
        public static SelectionTreeItem BuildTreeFrom(ITcSmTreeItem root, string displayOverride, bool selectable, int depthLimit)
        {
            var node = new SelectionTreeItem
            {
                DisplayName = displayOverride ?? SafeName(root),
                Path = SafePath(root),
                IsSelectable = selectable
            };

            if (depthLimit <= 0) return node;

            try
            {
                foreach (var child in EnumerateChildren(root))
                {
                    node.Children.Add(BuildTreeFrom(child, null, true, depthLimit - 1));
                }
            }
            catch
            {
                // manche virtuellen Knoten werfen sporadisch – dann lassen wir es bei der Ebene.
            }
            return node;
        }

        /// <summary>
        /// Robuste Auflistung der Kinder eines TreeItems:
        /// - nutzt ChildCount/Child[i], fängt COM- und allgemeine Ausnahmen ab
        /// - hat Fallback, falls ChildCount nicht zuverlässig ist
        /// </summary>
        public static IEnumerable<ITcSmTreeItem> EnumerateChildren(ITcSmTreeItem parent)
        {
            int count = 0;
            bool countOk = true;
            try { count = parent.ChildCount; }
            catch (COMException) { countOk = false; }
            catch { countOk = false; }

            if (countOk && count > 0)
            {
                for (int i = 1; i <= count; i++)
                {
                    ITcSmTreeItem child = null;
                    try { child = parent.Child[i]; }
                    catch (COMException) { child = null; }
                    catch { child = null; }
                    if (child != null) yield return child;
                }
                yield break;
            }

            // Fallback: versuche sequentiell, bis nichts mehr kommt (harte Obergrenze, um Hänger zu vermeiden)
            for (int i = 1; i <= 1024; i++)
            {
                ITcSmTreeItem child = null;
                try { child = parent.Child[i]; }
                catch (COMException) { child = null; }
                catch { child = null; }

                if (child == null) break;
                yield return child;
            }
        }

        public static void CheckAll(SelectionTreeItem node, bool check)
        {
            node.IsChecked = check && node.IsSelectable;
            foreach (var c in node.Children) CheckAll(c, check);
        }

        public static IEnumerable<SelectionTreeItem> EnumerateChecked(SelectionTreeItem node)
        {
            if (node.IsChecked) yield return node;
            foreach (var c in node.Children)
            {
                foreach (var cc in EnumerateChecked(c))
                    yield return cc;
            }
        }

        private static string SafeName(ITcSmTreeItem n)
        {
            try { return n.Name ?? string.Empty; } catch { return string.Empty; }
        }
        private static string SafePath(ITcSmTreeItem n)
        {
            try { return n.PathName ?? string.Empty; } catch { return string.Empty; }
        }
    }
}
