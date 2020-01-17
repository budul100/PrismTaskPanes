using NetOffice.OfficeApi.Enums;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.Settings;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace PrismTaskPanes.TaskPanes
{
    internal class SettingsRepository
    {
        #region Private Fields

        private readonly IEnumerable<PrismTaskPaneAttribute> attributes;
        private readonly List<PrismTaskPaneSettings> settings;

        private readonly string settingsPath;

        #endregion Private Fields

        #region Public Constructors

        public SettingsRepository(IEnumerable<PrismTaskPaneAttribute> attributes, string settingsPath)
        {
            this.attributes = attributes;

            this.settingsPath = settingsPath;
            settings = ReadSettings().ToList();
        }

        #endregion Public Constructors

        #region Public Methods

        public IEnumerable<PrismTaskPaneSettings> Get(int documentHash)
        {
            var result = GetAllSettings(documentHash).ToArray();

            return result;
        }

        public PrismTaskPaneSettings Get(int attributeHash, int documentHash)
        {
            var result = GetCurrentSettings(
                attributeHash: attributeHash,
                documentHash: documentHash);

            return result;
        }

        public void Save()
        {
            WriteSettings();
        }

        public void Set(int attributeHash, int documentHash, bool visible,
            int width, int height, MsoCTPDockPosition dockPosition)
        {
            UpdateSettings(
                attributeHash: attributeHash,
                documentHash: documentHash,
                visible: visible,
                width: width,
                height: height,
                dockPosition: dockPosition);
        }

        #endregion Public Methods

        #region Private Methods

        private IEnumerable<PrismTaskPaneSettings> GetAllSettings(int documentHash)
        {
            foreach (var attribute in attributes)
            {
                var result = GetCurrentSettings(
                    attributeHash: attribute.Hash,
                    documentHash: documentHash);

                yield return result;
            }
        }

        private PrismTaskPaneSettings GetCurrentSettings(int attributeHash, int documentHash)
        {
            var currentAttribute = attributes
                .FirstOrDefault(a => a.Hash == attributeHash);

            if (currentAttribute == null) throw new ApplicationException(
                $"There is no respective Prism Task Pane defined.");

            var result = settings?
                .Where(s => s.AttributeHash == attributeHash)
                .Where(s => s.DocumentHash == documentHash).LastOrDefault();

            if (result == null)
            {
                result = new PrismTaskPaneSettings
                {
                    AttributeHash = attributeHash,
                    DockPosition = currentAttribute.DockPosition,
                    DocumentHash = documentHash,
                    Height = currentAttribute.Height,
                    Visible = currentAttribute.Visible,
                    Width = currentAttribute.Width,
                };

                settings.Add(result);
            }

            result.DockRestriction = currentAttribute.DockRestriction;
            result.InvisibleAtStart = currentAttribute.InvisibleAtStart;
            result.NavigationKey = currentAttribute.NavigationKey;
            result.NavigationValue = currentAttribute.NavigationValue;
            result.RegionContext = currentAttribute.RegionContext;
            result.RegionName = currentAttribute.RegionName;
            result.Title = currentAttribute.Title;
            result.View = currentAttribute.View;

            return result;
        }

        private PrismTaskPaneSettingsCollection GetSettingsCollection()
        {
            var relevantSettings = settings
                .GroupBy(s => new { s.AttributeHash, s.DocumentHash })
                .Select(g => g.Last()).ToArray();

            var collection = new PrismTaskPaneSettingsCollection
            {
                Settings = relevantSettings,
            };

            return collection;
        }

        private IEnumerable<PrismTaskPaneSettings> ReadSettings()
        {
            if (File.Exists(settingsPath))
            {
                var collection = default(PrismTaskPaneSettingsCollection);
                var serializer = new XmlSerializer(typeof(PrismTaskPaneSettingsCollection));

                try
                {
                    using (StreamReader file = new StreamReader(settingsPath))
                    {
                        collection = serializer.Deserialize(file) as PrismTaskPaneSettingsCollection;
                    }
                }
                catch { }

                if (collection?.Settings?.Any() ?? false)
                {
                    foreach (var setting in collection.Settings)
                    {
                        yield return setting;
                    }
                }
            }
        }

        private void UpdateSettings(int attributeHash, int documentHash, bool visible, int width, int height,
            MsoCTPDockPosition dockPosition)
        {
            var currentSettings = GetCurrentSettings(
                attributeHash: attributeHash,
                documentHash: documentHash);

            currentSettings.DockPosition = dockPosition;
            currentSettings.Height = height;
            currentSettings.DocumentHash = documentHash;
            currentSettings.Visible = visible;
            currentSettings.Width = width;
        }

        private void WriteSettings()
        {
            var collection = GetSettingsCollection();

            if (!Directory.Exists(Path.GetDirectoryName(settingsPath)))
                Directory.CreateDirectory(Path.GetDirectoryName(settingsPath));

            var writer = new XmlSerializer(collection.GetType());
            using (var file = File.Create(settingsPath))
            {
                writer.Serialize(
                    stream: file,
                    o: collection);
            }
        }

        #endregion Private Methods
    }
}