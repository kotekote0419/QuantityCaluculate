using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

namespace UFlowPlant3D.Commands
{
    /// <summary>
    /// UFLOW_DIAG_DUMP_PORTS
    /// Plant3Dの「PnP3dObject」ではなく、PnP3dObjectsMgd の Part/Pipe/InlineAsset から Port を取得する版。
    /// （Plant3Dのサンプルでも Part.GetPorts(PortType.All) が使われる） 
    /// </summary>
    public class PortDumpCommands
    {
        [CommandMethod("UFLOW_DIAG_DUMP_PORTS")]
        public void DumpPorts()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            var peo = new PromptEntityOptions("\n[UFLOW][DIAG] Port一覧を確認するエンティティ（Pipe/Part/Fastenerなど）を選択：");
            peo.AllowNone = false;
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            var oid = per.ObjectId;

            // CSV出力先
            string outPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                $"PortDump_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            );

            try
            {
                // PnP3dObjectsMgd を確実にロード
                EnsureAssemblyLoaded(ed, @"C:\Program Files\Autodesk\AutoCAD 2026\PLNT3D\PnP3dObjectsMgd.dll");

                var ports = GetPortsByPartReflection(ed, db, oid).ToList();

                ed.WriteMessage($"\n[UFLOW][DIAG] Ports count={ports.Count}");
                foreach (var p in ports)
                    ed.WriteMessage($"\n    {p.Name} : {p.Position}");

                var s1 = ports.FirstOrDefault(p => string.Equals(p.Name, "S1", StringComparison.OrdinalIgnoreCase));
                var s2 = ports.FirstOrDefault(p => string.Equals(p.Name, "S2", StringComparison.OrdinalIgnoreCase));
                ed.WriteMessage($"\n[UFLOW][DIAG] S1={(s1 == null ? "not found" : s1.Position)} / S2={(s2 == null ? "not found" : s2.Position)}");

                // CSV
                using (var sw = new StreamWriter(outPath, false, System.Text.Encoding.UTF8))
                {
                    sw.WriteLine("Handle,ObjectType,PortName,Position");
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var ent = tr.GetObject(oid, OpenMode.ForRead) as Entity;
                        string handle = ent?.Handle.ToString() ?? "";
                        string objType = ent?.GetType().FullName ?? "";

                        foreach (var p in ports)
                        {
                            sw.WriteLine(string.Join(",",
                                CsvEsc(handle),
                                CsvEsc(objType),
                                CsvEsc(p.Name),
                                CsvEsc(p.Position)
                            ));
                        }
                        tr.Commit();
                    }
                }

                ed.WriteMessage($"\n[UFLOW][DIAG] PortDump CSV 出力: {outPath}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DIAG][ERROR] Portダンプに失敗: {ex.GetType().Name}: {ex.Message}");
                ed.WriteMessage("\n[UFLOW][DIAG] 追加ヒント：");
                ed.WriteMessage("\n  - この環境では PnP3dObject 型は存在しない可能性があります。PnP3dObjects.Part から Ports を取ります。");
                ed.WriteMessage("\n  - Pipe/Elbow/Valve/Flange などで先に試してください（Connectorは対象外のことがあります）。");
            }
        }

        // ----------------------------
        // helpers
        // ----------------------------

        private sealed class PortInfo
        {
            public string Name { get; set; } = "";
            public string Position { get; set; } = "";
        }

        private static void EnsureAssemblyLoaded(Editor ed, string dllPath)
        {
            try
            {
                // すでにロード済みなら何もしない
                if (AppDomain.CurrentDomain.GetAssemblies().Any(a =>
                {
                    try { return string.Equals(a.GetName().Name, "PnP3dObjectsMgd", StringComparison.OrdinalIgnoreCase); }
                    catch { return false; }
                }))
                {
                    return;
                }

                if (File.Exists(dllPath))
                {
                    Assembly.LoadFrom(dllPath);
                    ed.WriteMessage($"\n[UFLOW][DIAG] Loaded assembly: {dllPath}");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DIAG] Assembly load failed: {ex.GetType().Name}: {ex.Message}");
            }
        }

        private static IEnumerable<PortInfo> GetPortsByPartReflection(Editor ed, Database db, ObjectId oid)
        {
            // 型をロード済みアセンブリから取得
            var partType = FindLoadedType("Autodesk.ProcessPower.PnP3dObjects.Part");
            if (partType == null)
                throw new InvalidOperationException("PnP3dObjects.Part type not found. (PnP3dObjectsMgd.dll not loaded?)");

            var portTypeEnum = FindLoadedType("Autodesk.ProcessPower.PnP3dObjects.PortType");
            var portClassType = FindLoadedType("Autodesk.ProcessPower.PnP3dObjects.Port");
            if (portTypeEnum == null || portClassType == null)
                throw new InvalidOperationException("PnP3dObjects.Port/PortType not found.");

            // PortType.All を作る（なければ Both / 0）
            object portTypeAll = CreateEnumValue(portTypeEnum, new[] { "All", "Both", "AllPorts" });

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var dbo = tr.GetObject(oid, OpenMode.ForRead, false);
                if (dbo == null)
                    throw new InvalidOperationException("Selected object is null.");

                // dbo が Part 互換か？（Pipe/InlineAsset等は Part 派生のことが多い）
                if (!partType.IsAssignableFrom(dbo.GetType()))
                {
                    // ConnectorはこのAPIで取れることもあるが、取れないならここで打ち切り
                    throw new InvalidOperationException($"Selected object is not a PnP3dObjects.Part. actual={dbo.GetType().FullName}");
                }

                // Part.GetPorts(PortType)
                var miGetPorts = dbo.GetType().GetMethod("GetPorts", BindingFlags.Public | BindingFlags.Instance, null, new[] { portTypeEnum }, null)
                              ?? partType.GetMethod("GetPorts", BindingFlags.Public | BindingFlags.Instance, null, new[] { portTypeEnum }, null);

                if (miGetPorts == null)
                    throw new InvalidOperationException("GetPorts(PortType) method not found on Part.");

                object portsObj = miGetPorts.Invoke(dbo, new[] { portTypeAll });
                if (portsObj == null)
                    yield break;

                // PortCollection は IEnumerable のはず
                if (!(portsObj is IEnumerable enumPorts))
                    throw new InvalidOperationException($"Ports is not IEnumerable. type={portsObj.GetType().FullName}");

                foreach (var port in enumPorts)
                {
                    if (port == null) continue;

                    string name = TryGetPropString(port, "Name")
                               ?? TryGetPropString(port, "PortName")
                               ?? port.ToString();

                    string pos = TryGetPointLikeString(port, new[] { "Position", "Location", "Point", "Center" });

                    yield return new PortInfo { Name = name, Position = pos };
                }

                tr.Commit();
            }
        }

        private static Type FindLoadedType(string fullName)
        {
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                try
                {
                    var t = asm.GetType(fullName, throwOnError: false, ignoreCase: false);
                    if (t != null) return t;
                }
                catch { }
            }
            return null;
        }

        private static object CreateEnumValue(Type enumType, string[] preferredNames)
        {
            if (enumType == null || !enumType.IsEnum) return 0;

            var names = Enum.GetNames(enumType);
            foreach (var pn in preferredNames)
            {
                var hit = names.FirstOrDefault(n => string.Equals(n, pn, StringComparison.OrdinalIgnoreCase));
                if (hit != null) return Enum.Parse(enumType, hit);
            }

            // fallback: 0
            return Enum.ToObject(enumType, 0);
        }

        private static string TryGetPropString(object obj, string propName)
        {
            try
            {
                var pi = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                return pi?.GetValue(obj)?.ToString();
            }
            catch { return null; }
        }

        private static string TryGetPointLikeString(object obj, string[] propCandidates)
        {
            foreach (var pn in propCandidates)
            {
                try
                {
                    var pi = obj.GetType().GetProperty(pn, BindingFlags.Public | BindingFlags.Instance);
                    if (pi == null) continue;
                    var v = pi.GetValue(obj);
                    if (v == null) continue;
                    return v.ToString();
                }
                catch { }
            }
            return "";
        }

        private static string CsvEsc(string s)
        {
            s ??= "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n"))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }
    }
}
