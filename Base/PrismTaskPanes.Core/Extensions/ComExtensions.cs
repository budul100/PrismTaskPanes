#pragma warning disable CA1031 // Do not catch general exception types

using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;

namespace PrismTaskPanes.Core.Extensions
{
    internal static class ComExtensions
    {
        #region Public Methods

        public static string GetProgId(Type hostType, Type contentType)
        {
            var result = string.Join(
                "_",
                hostType.FullName,
                contentType.FullName);

            return result;
        }

        public static string Register(string progId, Type type)
        {
            var clsId = CreateClsId();

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

            return progId;
        }

        public static void Unregister(string progId)
        {
            using (var key = GetKeyClassesRoot())
            {
                key?.DeleteSubKeyTree(
                    subkey: progId,
                    throwOnMissingSubKey: false);
            }

            var clsId = GetClsId(progId);

            if (!string.IsNullOrWhiteSpace(clsId))
            {
                using (var key = GetKeyClsIdRoot())
                {
                    key?.DeleteSubKeyTree(
                        subkey: clsId,
                        throwOnMissingSubKey: false);
                }
            }
        }

        #endregion Public Methods

        #region Private Methods

        private static string CreateClsId()
        {
            return Guid.NewGuid().ToString("B").ToUpperInvariant();
        }

        private static string GetAssemblyPath(this Type type)
        {
            return new Uri(type.Assembly.Location).LocalPath;
        }

        private static string GetClsId(string progId)
        {
            var result = default(string);

            using (var key = GetKeyProgId(
                progId: progId,
                name: "CLSID"))
            {
                result = key?.GetValue(default)?.ToString();
            }

            return result;
        }

        private static RegistryKey GetKeyClassesRoot(string name = default)
        {
            var key = Registry.ClassesRoot.OpenSubKey(
                name: name,
                writable: true);

            if (key != default)
            {
                return key;
            }

            var parentPath = Path.GetDirectoryName(name);

            if (string.IsNullOrEmpty(parentPath))
            {
                return Registry.CurrentUser.CreateSubKey(name);
            }

            using (var parentKey = GetKeyClassesRoot(parentPath))
            {
                return parentKey?.CreateSubKey(Path.GetFileName(name));
            }
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
                path2: type.Assembly.FullName);

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
                    value: type.FullName);
            }

            using (var key = keyGetter.Invoke("CLSID"))
            {
                key?.SetValue(
                    name: default,
                    value: clsId);
            }
        }

        #endregion Private Methods
    }
}

#pragma warning restore CA1031 // Do not catch general exception types