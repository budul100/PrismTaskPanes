using Microsoft.Win32;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace PrismTaskPanes.Extensions
{
    [ComVisible(false)]
    internal static class ComExtensions
    {
        #region Private Fields

        private const string ClassesSeparator = "_";

        #endregion Private Fields

        #region Public Methods

        public static string GetProgId(this Type hostType, Type contentType)
        {
            var result = string.Join(
                ClassesSeparator,
                hostType.FullName,
                contentType.FullName);

            return result;
        }

        public static void Register(string progId, Type type, string directory = default)
        {
            var clsId = type.GUID;

            RegisterProgId(
                clsId: clsId,
                progId: progId);

            RegisterClsId(
                clsId: clsId,
                type: type,
                directory: directory);
        }

        public static void Unregister(string progId)
        {
            var given = Type.GetTypeFromProgID(progId);

            if (given != default)
            {
                using var key = GetKeyClassesRoot();

                key?.DeleteSubKeyTree(
                    subkey: progId,
                    throwOnMissingSubKey: false);
            }
        }

        #endregion Public Methods

        #region Private Methods

        private static string AsString(this Guid guid)
        {
            return guid.ToString("B").ToUpperInvariant();
        }

        private static string GetComHostPath(this Type type, string directory)
        {
            var result = default(string);

            var assemblyName = Path.GetFileNameWithoutExtension(type.Assembly.Location);
            var assemblyExtension = Path.GetExtension(type.Assembly.Location);

            var fileName = $"{assemblyName}.comhost{assemblyExtension}";

            if (!string.IsNullOrWhiteSpace(directory))
            {
                result = Path.Combine(
                    path1: directory,
                    path2: fileName);
            }

            if (!File.Exists(result))
            {
                result = Path.Combine(
                    path1: Path.GetDirectoryName(type.Assembly.Location),
                    path2: fileName);
            }

            return result;
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

        private static RegistryKey GetKeyClsId(Guid clsId, string name = default)
        {
            var path = Path.Combine(
                clsId.AsString(),
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

        private static RegistryKey GetSubKey(RegistryKey root, string name)
        {
            name ??= string.Empty;

            var result = root.OpenSubKey(
                name: name,
                writable: true);

            if (result == default)
            {
                var parent = Path.GetDirectoryName(name);

                if (string.IsNullOrEmpty(parent))
                {
                    _ = root.CreateSubKey(
                        subkey: name);
                }
                else
                {
                    using var key = GetSubKey(
                        root: root,
                        name: parent);
                    _ = key.CreateSubKey(
                        subkey: Path.GetFileName(name));
                }

                result = root.OpenSubKey(
                    name: name,
                    writable: true);
            }

            return result;
        }

        private static void RegisterClsId(Guid clsId, Type type, string directory)
        {
            RegistryKey baseKey(string name)
            {
                return GetKeyClsId(
                    clsId: clsId,
                    name: name);
            }

            using (var key = baseKey(default))
            {
                key?.SetValue(
                    name: default,
                    value: "CoreCLR COMHost Server");
            }

            using (var key = baseKey("ProgId"))
            {
                key?.SetValue(
                    name: default,
                    value: type.FullName);
            }

            using (var key = baseKey("InprocServer32"))
            {
                key?.SetValue(
                    name: default,
                    value: type.GetComHostPath(directory));

                key?.SetValue(
                    name: "ThreadingModel",
                    value: "Both");
            }
        }

        private static void RegisterProgId(Guid clsId, string progId)
        {
            RegistryKey baseKey(string name)
            {
                return GetKeyProgId(
                    progId: progId,
                    name: name);
            }

            using (var key = baseKey(default))
            {
                key?.SetValue(
                    name: default,
                    value: progId);
            }

            using (var key = baseKey("CLSID"))
            {
                key?.SetValue(
                    name: default,
                    value: clsId.AsString());
            }
        }

        #endregion Private Methods
    }
}