using Microsoft.Win32;
using PrismTaskPanes.Controls;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace PrismTaskPanes.Extensions
{
    [ComVisible(false)]
    public static class HostExtensions
    {
        #region Public Methods

        public static void RegisterTaskPaneHost(this Type type, string sourceDirectory = default)
        {
            if (type is null)
            {
                throw new ArgumentNullException(nameof(type));
            }

            var progId = typeof(PrismTaskPanesHost).FullName;
            var given = Type.GetTypeFromProgID(progId);

            if (given == default)
            {
                var sourcePathComHost = GetSourcePath<PrismTaskPanesHost>(
                    directory: sourceDirectory,
                    extension: "comhost.dll");

                RegisterProgId(
                    clsId: typeof(PrismTaskPanesHost).GUID,
                    progId: progId);

                RegisterClsId(
                    clsId: typeof(PrismTaskPanesHost).GUID,
                    progId: progId,
                    comHostPath: sourcePathComHost);
            }
        }

        public static void UnregisterTaskPaneHost(this Type type)
        {
            if (type is null)
            {
                throw new ArgumentNullException(nameof(type));
            }

            var progId = typeof(PrismTaskPanesHost).FullName;
            var given = Type.GetTypeFromProgID(progId);

            if (given != default)
            {
                using var clsIdRoot = GetRootClsId();

                clsIdRoot?.DeleteSubKeyTree(
                    subkey: typeof(PrismTaskPanesHost).GUID.ToString(),
                    throwOnMissingSubKey: false);

                using var progIdRoot = GetRootProgId();

                progIdRoot?.DeleteSubKeyTree(
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

        private static RegistryKey GetKeyClsId(this Guid clsId, string name = default)
        {
            var path = Path.Combine(
                clsId.AsString(),
                name ?? string.Empty);

            var result = GetRootClsId(path);

            return result;
        }

        private static RegistryKey GetKeyProgId(this string progId, string name = default)
        {
            var path = Path.Combine(
                progId,
                name ?? string.Empty);

            var result = GetRootProgId(path);

            return result;
        }

        private static RegistryKey GetRootClsId(string name = default)
        {
            var path = Path.Combine(
                "WOW6432Node",
                "CLSID",
                name ?? string.Empty);

            var result = GetRootProgId(path);

            return result;
        }

        private static RegistryKey GetRootProgId(string name = default)
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

        private static string GetSourcePath<T>(string directory, string extension)
        {
            var result = default(string);

            var type = typeof(T);

            var assemblyName = Path.GetFileNameWithoutExtension(type.Assembly.Location);
            var fileName = $"{assemblyName}.{extension}";

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

            if (!File.Exists(result))
            {
                throw new FileNotFoundException(
                    message: "The PrismTaskPanes component cannot be found.",
                    fileName: result);
            }

            return result;
        }

        private static RegistryKey GetSubKey(this RegistryKey root, string name)
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
                    root.CreateSubKey(
                        subkey: name);
                }
                else
                {
                    using var key = GetSubKey(
                        root: root,
                        name: parent);

                    key.CreateSubKey(
                        subkey: Path.GetFileName(name));
                }

                result = root.OpenSubKey(
                    name: name,
                    writable: true);
            }

            return result;
        }

        private static void RegisterClsId(this Guid clsId, string progId, string comHostPath)
        {
            RegistryKey baseKey(string name)
            {
                return clsId.GetKeyClsId(name);
            }

            using (var key = baseKey(default))
            {
                key?.SetValue(
                    name: default,
                    value: "CoreCLR COMHost Server");
            }

            using (var key = baseKey("ProgID"))
            {
                key?.SetValue(
                    name: default,
                    value: progId);
            }

            using (var key = baseKey("InProcServer32"))
            {
                key?.SetValue(
                    name: default,
                    value: comHostPath);

                key?.SetValue(
                    name: "ThreadingModel",
                    value: "Both");
            }
        }

        private static void RegisterProgId(this Guid clsId, string progId)
        {
            RegistryKey baseKey(string name)
            {
                return progId.GetKeyProgId(name);
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