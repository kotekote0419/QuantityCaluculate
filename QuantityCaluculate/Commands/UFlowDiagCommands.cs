using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PlantInstance;

namespace UFlowPlant3D.Commands
{
    public class UFlowDiagCommands
    {
        [CommandMethod("UFLOW_DUMP_PNP_TABLES")]
        public void DumpPnPTables()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            var dlm = PlantApplication.CurrentProject?.ProjectParts["Piping"]?.DataLinksManager;
            if (dlm == null)
            {
                ed.WriteMessage("\n[UFLOW] DataLinksManagerが取得できません。");
                return;
            }

            object pnpDb = GetPnPDatabase(dlm);
            if (pnpDb == null)
            {
                ed.WriteMessage("\n[UFLOW] PnPDatabaseが取得できません。");
                return;
            }

            object tables = GetProp(pnpDb, "Tables") ?? Invoke0(pnpDb, "GetTables");
            if (tables == null)
            {
                ed.WriteMessage("\n[UFLOW] Tablesが取得できません。");
                return;
            }

            ed.WriteMessage("\n[UFLOW] --- PnP Tables ---");
            foreach (var name in EnumerateTableNames(tables))
                ed.WriteMessage($"\n  - {name}");
        }

        private static object GetPnPDatabase(DataLinksManager dlm)
            => GetProp(dlm, "PnPDatabase") ?? Invoke0(dlm, "GetPnPDatabase");

        private static object GetProp(object obj, string name)
        {
            if (obj == null) return null;
            try
            {
                var pi = obj.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
                return pi?.GetValue(obj);
            }
            catch { return null; }
        }

        private static object Invoke0(object obj, string method)
        {
            if (obj == null) return null;
            try
            {
                var mi = obj.GetType().GetMethod(method, BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null);
                return mi?.Invoke(obj, null);
            }
            catch { return null; }
        }

        private static IEnumerable<string> EnumerateTableNames(object tables)
        {
            // IDictionary
            if (tables is IDictionary dict)
            {
                foreach (DictionaryEntry de in dict)
                    if (de.Key != null) yield return de.Key.ToString();
                yield break;
            }

            // IEnumerable
            if (tables is IEnumerable e)
            {
                foreach (var it in e)
                {
                    if (it is DictionaryEntry de && de.Key != null)
                        yield return de.Key.ToString();
                    else
                    {
                        var name = GetProp(it, "Name")?.ToString()
                                ?? GetProp(it, "Key")?.ToString();
                        if (!string.IsNullOrWhiteSpace(name))
                            yield return name;
                    }
                }
            }
        }
    }
}
