#pragma warning disable CA1031 // Do not catch general exception types

using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace PrismTaskPanes.Core.Extensions
{
    [ComVisible(false)]
    internal static class ComExtensions
    {
        #region Private Fields

        private const string ClassesSeparator = "_";

        #endregion Private Fields

        #region Public Methods

        public static string GetProgId(Type hostType, Type contentType)
        {
            var result = string.Join(
                ClassesSeparator,
                hostType.FullName,
                contentType.FullName);

            return result;
        }

        public static void Register(string progId, Type type)
        {
            var given = Type.GetTypeFromProgID(progId);

            if (given == default)
            {
                var clsId = Guid.NewGuid().AsString();

                RegistryKey progIdKeyGetter(string name) => GetKeyProgId(
                    progId: progId,
                    name: name);

                RegisterProgId(
                    keyGetter: progIdKeyGetter,
                    type: type,
                    clsId: clsId);

                RegistryKey clsIdKeyGetter(string name) => GetKeyClsId(
                    clsId: clsId,
                    name: name);

                RegisterClsId(
                    keyGetter: clsIdKeyGetter,
                    type: type,
                    progId: progId);
            }
        }

        public static void Unregister(string progId)
        {
            var given = Type.GetTypeFromProgID(progId);

            if (given != default)
            {
                using (var key = GetKeyClassesRoot())
                {
                    key?.DeleteSubKeyTree(
                        subkey: progId,
                        throwOnMissingSubKey: false);
                }

                using (var key = GetKeyClsIdRoot())
                {
                    key?.DeleteSubKeyTree(
                        subkey: given.GUID.AsString(),
                        throwOnMissingSubKey: false);
                }
            }
        }

        #endregion Public Methods

        #region Private Methods

        private static string AsString(this Guid guid)
        {
            return guid.ToString("B").ToUpperInvariant();
        }

        private static string GetAssemblyPath(this Type type)
        {
            return new Uri(type.Assembly.Location).LocalPath;
        }

        private static RegistryKey GetKeyClassesRoot(string name = default)
        {
            var path = Path.Combine(
                "Software",
                "Classes",
                name ?? string.Empty);

            var result = GetSubKey(
                root: Registry.LocalMachine,
                name: path);

            return result;
        }

        private static RegistryKey GetKeyClsId(string clsId, string name = default)
        {
            var path = Path.Combine(
                clsId,
                name ?? string.Empty);

            var result = GetKeyClsIdRoot(path);

            return result;
        }

        private static RegistryKey GetKeyClsIdRoot(string name = default)
        {
            var path = Path.Combine(
                "WOW6432Node",
                "CLSID",
                name ?? string.Empty);

            var result = GetKeyClassesRoot(path);

            return result;
        }

        private static RegistryKey GetKeyProgId(string progId, string name = default)
        {
            var path = Path.Combine(
                progId,
                name ?? string.Empty);

            var result = GetKeyClassesRoot(path);

            return result;
        }

        private static string GetRuntimeVersion(this Type type)
        {
            var definition = "v4.0.30319"; // use CLR4 as the default

            try
            {
                var mscorlib = type.Assembly.GetReferencedAssemblies()
                    .FirstOrDefault(a => a.Name == "mscorlib");

                if (mscorlib != default
                    && mscorlib.Version.Major < 4)
                {
                    return "v2.0.50727"; // use CLR2
                }
            }
            catch
            { }

            return definition;
        }

        private static RegistryKey GetSubKey(RegistryKey root, string name)
        {
            name = name ?? string.Empty;

            var result = root.OpenSubKey(
                name: name,
                writable: true);

            if (result == default)
            {
                var parent = Path.GetDirectoryName(name);

                if (string.IsNullOrEmpty(parent))
                {
                    root.CreateSubKey(
                        subkey: name);
                }
                else
                {
                    using (RegistryKey key = GetSubKey(
                        root: root,
                        name: parent))
                    {
                        key.CreateSubKey(
                            subkey: Path.GetFileName(name));
                    }
                }

                result = root.OpenSubKey(
                    name: name,
                    writable: true);
            }

            return result;
        }

        private static void RegisterClsId(Func<string, RegistryKey> keyGetter, Type type, string progId)
        {
            using (var key = keyGetter.Invoke(default))
            {
                key?.SetValue(
                    name: default,
                    value: type.FullName);
            }

            using (var key = keyGetter.Invoke("ProgId"))
            {
                key?.SetValue(
                    name: default,
                    value: progId);
            }

            var categoriesPath = Path.Combine(
                path1: "Implemented Categories",
                path2: "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}");

            using (var key = keyGetter.Invoke(categoriesPath))
            { }

            var assemblyPath = type.GetAssemblyPath();
            var runtimeVersion = type.GetRuntimeVersion();

            using (var key = keyGetter.Invoke("InprocServer32"))
            {
                key?.SetValue(
                    name: default,
                    value: "mscoree.dll");

                key?.SetValue(
                    name: "Assembly",
                    value: type.Assembly.FullName);

                key?.SetValue(
                    name: "Class",
                    value: type.FullName);

                key?.SetValue(
                    name: "ThreadingModel",
                    value: "Both");

                key?.SetValue(
                    name: "CodeBase",
                    value: assemblyPath);

                key?.SetValue(
                    name: "RuntimeVersion",
                    value: runtimeVersion);
            }

            var versionPath = Path.Combine(
                path1: "InprocServer32",
                path2: type.Assembly.GetName().Version.ToString());

            using (var key = keyGetter.Invoke(versionPath))
            {
                key?.SetValue(
                    name: "Assembly",
                    value: type.Assembly.FullName);

                key?.SetValue(
                    name: "Class",
                    value: type.FullName);

                key?.SetValue(
                    name: "CodeBase",
                    value: assemblyPath);

                key?.SetValue(
                    name: "RuntimeVersion",
                    value: runtimeVersion);
            }
        }

        private static void RegisterProgId(Func<string, RegistryKey> keyGetter, Type type, string clsId)
        {
            using (var key = keyGetter.Invoke(default))
            {
                key?.SetValue(
                    name: default,
                    value: (string)type.FullName,
                    valueKind: RegistryValueKind.String);
            }

            using (var key = keyGetter.Invoke("CLSID"))
            {
                key?.SetValue(
                    name: default,
                    value: (string)clsId,
                    valueKind: RegistryValueKind.String);
            }
        }

        #endregion Private Methods
    }
}

#pragma warning restore CA1031 // Do not catch general exception types