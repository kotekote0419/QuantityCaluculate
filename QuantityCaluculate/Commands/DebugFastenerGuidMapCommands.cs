using System;
using System.Collections.Generic;
using System.Linq;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using Autodesk.ProcessPower.DataLinks;
using UFlowPlant3D.Services;

namespace UFlow
{
    public class DebugFastenerGuidMapCommands
    {
        [CommandMethod("UFLOW_DEBUG_FIND_FASTENER_BY_PICKED_CONNECTOR")]
        public void UFLOW_DEBUG_FIND_FASTENER_BY_PICKED_CONNECTOR()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            DataLinksManager dlm;
            try { dlm = DataLinksManager.GetManager(db); }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] DLM get failed: {ex.Message}");
                return;
            }

            // pick connector
            var peo = new PromptEntityOptions("\n[UFLOW][DBG] Pick Connector (gasket/boltset-like): ");
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            int connRowId = -1;
            try { connRowId = dlm.FindAcPpRowId(per.ObjectId); }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] FindAcPpRowId failed: {ex.Message}");
                return;
            }

            var connProps = SafeGetAllProps(dlm, connRowId);
            string guid = GetFirstString(connProps, "PnPGuid", "RowGuid", "GUID", "Guid");

            ed.WriteMessage($"\n[UFLOW][DBG] PickedRowId={connRowId}, PnPGuid='{guid}'");

            if (string.IsNullOrWhiteSpace(guid))
            {
                ed.WriteMessage("\n[UFLOW][DBG] GUID not found on connector row. Stop.");
                return;
            }

            // Fastener rowIds (PnPDB Fasteners table)
            List<int> fastenerRows;
            try
            {
                fastenerRows = FastenerCollector.CollectFastenerRowIds(dlm, ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] FastenerCollector failed: {ex.Message}");
                return;
            }

            // Find matches: any property value equals guid
            var matches = new List<int>();

            foreach (int fr in fastenerRows)
            {
                var props = SafeGetAllProps(dlm, fr);
                if (props == null || props.Count == 0) continue;

                string connIdStr = connRowId.ToString();

                bool hit = props.Values.Any(v =>
                {
                    if (string.IsNullOrWhiteSpace(v)) return false;
                    var s = v.Trim();
                    if (string.Equals(s, guid, StringComparison.OrdinalIgnoreCase)) return true;
                    if (string.Equals(s, connIdStr, StringComparison.OrdinalIgnoreCase)) return true;
                    // 数値だけ抜き出して比較
                    var digits = new string(s.Where(char.IsDigit).ToArray());
                    return digits == connIdStr;
                });

                if (hit) matches.Add(fr);

            }

            ed.WriteMessage($"\n[UFLOW][DBG] FastenerRows matched by GUID = {matches.Count}");

            // show some detail
            foreach (int fr in matches.Take(20))
            {
                var props = SafeGetAllProps(dlm, fr);
                string rowGuid = GetFirstString(props, "RowGuid", "PnPGuid", "Guid");
                string className = GetFirstString(props, "ClassName", "PnPClassName");
                ed.WriteMessage($"\n[UFLOW][DBG]  FastenerRowId={fr}, Class='{className}', Guid='{rowGuid}'");
            }

            if (matches.Count > 20)
                ed.WriteMessage($"\n[UFLOW][DBG]  ... (more {matches.Count - 20})");
        }

        private static Dictionary<string, string> SafeGetAllProps(DataLinksManager dlm, int rowId)
        {
            try
            {
                var obj = dlm.GetAllProperties(rowId, true);
                if (obj == null) return new Dictionary<string, string>();

                // IDictionary<string, object> っぽいものが返る想定（環境差を吸収）
                var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                foreach (var kv in obj)
                {
                    // kv は KeyValuePair<string, object> で来る環境が多い
                    var keyProp = kv.GetType().GetProperty("Key");
                    var valProp = kv.GetType().GetProperty("Value");
                    if (keyProp == null || valProp == null) continue;

                    var k = keyProp.GetValue(kv)?.ToString();
                    var v = valProp.GetValue(kv)?.ToString();

                    if (!string.IsNullOrWhiteSpace(k))
                        dict[k] = v ?? "";
                }

                return dict;
            }
            catch
            {
                return new Dictionary<string, string>();
            }
        }

        private static string GetFirstString(Dictionary<string, string> props, params string[] keys)
        {
            if (props == null) return "";
            foreach (var k in keys)
            {
                var hit = props.Keys.FirstOrDefault(x => x.Equals(k, StringComparison.OrdinalIgnoreCase));
                if (hit != null)
                {
                    var v = props[hit];
                    if (!string.IsNullOrWhiteSpace(v)) return v.Trim();
                }
            }

            // 最後の保険：Guidっぽい値を拾う
            foreach (var v in props.Values)
            {
                if (!string.IsNullOrWhiteSpace(v) && v.Length >= 32 && v.Contains("-"))
                    return v.Trim();
            }

            return "";
        }
    }
}
