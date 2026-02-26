using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

namespace UFlowPlant3D.Commands
{
    /// <summary>
    /// デバッグ用：PnP3dObjectsMgd.dll がロードされても PnP3dObject 型が見つからない原因を特定するための診断コマンド。
    /// - ロード済みアセンブリ一覧から "PnP3d" を含む型を列挙
    /// - ReflectionTypeLoadException の LoaderExceptions を表示
    /// </summary>
    public class PnP3dTypeDiagCommands
    {
        [CommandMethod("UFLOW_DIAG_PNP3D_TYPES")]
        public void DumpPnP3dTypes()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            // 1) まず PnP3dObjectsMgd.dll を明示ロード（前回と同じ場所）
            string dllPath = @"C:\Program Files\Autodesk\AutoCAD 2026\PLNT3D\PnP3dObjectsMgd.dll";
            try
            {
                if (File.Exists(dllPath))
                {
                    Assembly.LoadFrom(dllPath);
                    ed.WriteMessage($"\n[UFLOW][DIAG] Loaded: {dllPath}");
                }
                else
                {
                    ed.WriteMessage($"\n[UFLOW][DIAG] DLL not found: {dllPath}");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DIAG] LoadFrom failed: {ex.GetType().Name}: {ex.Message}");
            }

            // 2) ロード済みアセンブリから PnP3dObjects* を探す
            var asms = AppDomain.CurrentDomain.GetAssemblies()
                .Where(a =>
                {
                    try { return (a.GetName().Name ?? "").IndexOf("PnP3d", StringComparison.OrdinalIgnoreCase) >= 0; }
                    catch { return false; }
                })
                .ToList();

            ed.WriteMessage($"\n[UFLOW][DIAG] Assemblies containing 'PnP3d': {asms.Count}");
            foreach (var a in asms)
            {
                string name = "";
                string loc = "";
                try { name = a.FullName; } catch { }
                try { loc = a.Location; } catch { }
                ed.WriteMessage($"\n  - {name}");
                if (!string.IsNullOrWhiteSpace(loc)) ed.WriteMessage($"\n      {loc}");
            }

            // 3) 型一覧：PnP3d を含む型を最大200件表示
            int shown = 0;
            foreach (var asm in asms)
            {
                Type[] types = null;
                try
                {
                    types = asm.GetTypes();
                }
                catch (ReflectionTypeLoadException rtle)
                {
                    types = rtle.Types.Where(t => t != null).ToArray();

                    ed.WriteMessage($"\n[UFLOW][DIAG] ReflectionTypeLoadException in {asm.GetName().Name}:");
                    foreach (var le in rtle.LoaderExceptions.Take(20))
                    {
                        if (le == null) continue;
                        ed.WriteMessage($"\n    LoaderException: {le.GetType().Name}: {le.Message}");
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n[UFLOW][DIAG] GetTypes failed in {asm.GetName().Name}: {ex.GetType().Name}: {ex.Message}");
                    continue;
                }

                foreach (var t in types)
                {
                    if (t == null) continue;

                    string fn = t.FullName ?? "";
                    if (fn.IndexOf("PnP3d", StringComparison.OrdinalIgnoreCase) < 0) continue;

                    ed.WriteMessage($"\n[UFLOW][DIAG] TYPE: {fn}");
                    shown++;
                    if (shown >= 200) break;
                }
                if (shown >= 200) break;
            }

            // 4) 期待の型名候補を探す
            var candidateNames = new[] { "PnP3dObject", "Pnp3dObject", "PnP3dPort", "Port" };
            foreach (var cn in candidateNames)
            {
                var hits = FindTypesContaining(asms, cn).Take(30).ToList();
                ed.WriteMessage($"\n[UFLOW][DIAG] Types containing '{cn}': {hits.Count}");
                foreach (var h in hits)
                    ed.WriteMessage($"\n    {h}");
            }
        }

        private static IEnumerable<string> FindTypesContaining(List<Assembly> asms, string token)
        {
            foreach (var asm in asms)
            {
                Type[] types;
                try { types = asm.GetTypes(); }
                catch (ReflectionTypeLoadException rtle) { types = rtle.Types.Where(t => t != null).ToArray(); }
                catch { continue; }

                foreach (var t in types)
                {
                    if (t == null) continue;
                    var fn = t.FullName ?? "";
                    if (fn.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0)
                        yield return fn;
                }
            }
        }
    }
}
