using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using TCatSysManagerLib;
using TwinCATModuleTransfer.ViewModels;

namespace TwinCATModuleTransfer.Services
{
    /// <summary>
    /// Robuste Erkennung von NC/Motion-Bereichen im TwinCAT-Baum.
    /// - Probiert bekannte Wurzeln (TMC/TM/NC/Motion).
    /// - Sucht gezielt nach "Hotspots" (Axes/Axis/Channels/Tasks/NC/Motion/CNC/Interpolation).
    /// - BFS mit Tiefe-Limit und COM-Fehler-Toleranz.
    /// </summary>
    public static class NcDiscoveryService
    {
        /// <summary>
        /// Liefert auswählbare NC-/Motion-Teilbäume (als SelectionTreeItem), die in die UI übernommen werden können.
        /// </summary>
        /// <param name="sys">SystemManager</param>
        /// <param name="maxDepthPerRoot">Tiefe für den UI-Baum</param>
        public static IList<SelectionTreeItem> DiscoverNcRoots(ITcSysManager sys, int maxDepthPerRoot = 6)
        {
            var result = new List<SelectionTreeItem>();
            var pathsTried = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 1) HARDCODED-KANDIDATEN (häufigste Wurzeln/Unterpfade)
            var candidates = new[]
            {
                "TMC", "TMC^Axes", "TMC^Axis", "TMC^Tasks", "TMC^Channels",
                "TM",  "TM^Axes",  "TM^Axis",  "TM^Tasks",  "TM^Channels",
                "NC",  "NC^Axes",  "NC^Axis",  "NC^Tasks",  "NC^Channels",
                "Motion", "Motion^Axes", "Motion^Axis"
            };

            foreach (var c in candidates)
            {
                var root = TryLookup(sys, c);
                if (root == null) continue;

                if (!pathsTried.Contains(root.PathName))
                {
                    pathsTried.Add(root.PathName);
                    result.Add(ProjectTreeService.BuildTreeFrom(root, $"[NC/Motion] {SafeName(root)}", true, maxDepthPerRoot));
                }
            }

            // 2) HEURISTISCHE SUCHE von bekannten möglichen Wurzeln aus
            foreach (var startKey in new[] { "TMC", "TM", "NC", "Motion" })
            {
                var start = TryLookup(sys, startKey);
                if (start == null) continue;

                foreach (var hs in FindNcHotspots(start))
                {
                    if (!pathsTried.Contains(hs.PathName))
                    {
                        pathsTried.Add(hs.PathName);
                        result.Add(ProjectTreeService.BuildTreeFrom(hs, $"[NC] {SafeName(hs)}", true, maxDepthPerRoot));
                    }
                }
            }

            // 3) Optional: letzter Fallback – noch breiter unter "TMC" oder "TM" schauen
            if (result.Count == 0)
            {
                foreach (var fb in new[] { "TMC", "TM", "NC" })
                {
                    var r = TryLookup(sys, fb);
                    if (r == null) continue;

                    foreach (var n in BroadScan(r, 4)) // breiter, aber limitiert
                    {
                        var nm = SafeName(n);
                        if (LooksLikeNcNode(nm) && !pathsTried.Contains(n.PathName))
                        {
                            pathsTried.Add(n.PathName);
                            result.Add(ProjectTreeService.BuildTreeFrom(n, $"[NC?] {nm}", true, maxDepthPerRoot));
                        }
                    }
                    if (result.Count > 0) break;
                }
            }

            return result;
        }

        private static IEnumerable<ITcSmTreeItem> BroadScan(ITcSmTreeItem root, int maxDepth)
        {
            var q = new Queue<Tuple<ITcSmTreeItem, int>>();
            q.Enqueue(Tuple.Create(root, 0));

            while (q.Count > 0)
            {
                var t = q.Dequeue();
                var node = t.Item1;
                var depth = t.Item2;
                yield return node;

                if (depth >= maxDepth) continue;

                int childCount = 0;
                try { childCount = node.ChildCount; } catch (COMException) { childCount = 0; }

                for (int i = 1; i <= childCount; i++)
                {
                    ITcSmTreeItem child = null;
                    try { child = node.Child[i]; } catch (COMException) { }
                    if (child != null) q.Enqueue(Tuple.Create(child, depth + 1));
                }
            }
        }

        private static List<ITcSmTreeItem> FindNcHotspots(ITcSmTreeItem start)
        {
            var hits = new List<ITcSmTreeItem>();
            var names = new[]
            {
                "Axes","Axis","Axis_","Axis ","Axis1","Axis 1","AxisGroup",
                "Channels","Channel",
                "Tasks","Task",
                "NC","Motion","CNC","Interpolation"
            };

            foreach (var node in BroadScan(start, 3)) // kompakt halten
            {
                var nm = SafeName(node);
                // Direkter Namenstreffer?
                if (names.Any(k => nm.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0))
                    hits.Add(node);
                // Heuristik: "Axis" im Namen -> Kandidat
                else if (nm.IndexOf("axis", StringComparison.OrdinalIgnoreCase) >= 0)
                    hits.Add(node);
                // Heuristik: "NC" im Namen -> Kandidat
                else if (nm.IndexOf("nc", StringComparison.OrdinalIgnoreCase) >= 0)
                    hits.Add(node);
            }

            // Eindeutige Pfade
            var unique = new Dictionary<string, ITcSmTreeItem>(StringComparer.OrdinalIgnoreCase);
            foreach (var h in hits)
            {
                var p = SafePath(h);
                if (!unique.ContainsKey(p))
                    unique[p] = h;
            }
            return unique.Values.ToList();
        }

        private static ITcSmTreeItem TryLookup(ITcSysManager sys, string path)
        {
            try { return sys.LookupTreeItem(path); }
            catch { return null; }
        }

        private static string SafeName(ITcSmTreeItem n)
        {
            try { return n.Name ?? string.Empty; } catch { return string.Empty; }
        }

        private static string SafePath(ITcSmTreeItem n)
        {
            try { return n.PathName ?? string.Empty; } catch { return string.Empty; }
        }

        private static bool LooksLikeNcNode(string name)
        {
            var n = name ?? string.Empty;
            return n.IndexOf("axis", StringComparison.OrdinalIgnoreCase) >= 0
                || n.IndexOf("axes", StringComparison.OrdinalIgnoreCase) >= 0
                || n.IndexOf("task", StringComparison.OrdinalIgnoreCase) >= 0
                || n.IndexOf("channel", StringComparison.OrdinalIgnoreCase) >= 0
                || n.IndexOf("nc", StringComparison.OrdinalIgnoreCase) >= 0
                || n.IndexOf("motion", StringComparison.OrdinalIgnoreCase) >= 0
                || n.IndexOf("cnc", StringComparison.OrdinalIgnoreCase) >= 0
                || n.IndexOf("interpolation", StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}