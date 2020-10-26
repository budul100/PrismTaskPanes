#pragma warning disable CA1031 // Keine allgemeinen Ausnahmetypen abfangen

using NetOffice.OfficeApi.Enums;
using PrismTaskPanes.Attributes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Serialization;

namespace PrismTaskPanes.Settings
{
    internal class TaskPaneSettingsRepository
    {
        #region Private Fields

        private readonly IList<PrismTaskPaneAttribute> attributes = new List<PrismTaskPaneAttribute>();
        private readonly IList<TaskPaneSettings> configurations;

        private readonly string configurationsPath;

        #endregion Private Fields

        #region Public Constructors

        public TaskPaneSettingsRepository(string configurationsPath)
        {
            this.configurationsPath = configurationsPath;

            try
            {
                configurations = ReadConfigurations().ToList();
            }
            catch
            {
                configurations = new List<TaskPaneSettings>();
            }
        }

        #endregion Public Constructors

        #region Public Methods

        public void AddAttributes(IEnumerable<PrismTaskPaneAttribute> attributes)
        {
            foreach (var attribute in attributes)
            {
                this.attributes.Add(attribute);
            }
        }

        public IEnumerable<TaskPaneSettings> Get(string documentHash)
        {
            var result = GetAllConfigurations(documentHash).ToArray();

            return result;
        }

        public TaskPaneSettings Get(string receiverHash, string documentHash)
        {
            var result = GetCurrentConfiguration(
                receiverHash: receiverHash,
                documentHash: documentHash);

            return result;
        }

        public void Save()
        {
            WriteConfigurations();
        }

        public void Set(string receiverHash, string documentHash, bool visible,
            int width, int height, MsoCTPDockPosition dockPosition)
        {
            UpdateConfiguration(
                attributeHash: receiverHash,
                documentHash: documentHash,
                visible: visible,
                width: width,
                height: height,
                dockPosition: dockPosition);
        }

        #endregion Public Methods

        #region Private Methods

        private IEnumerable<TaskPaneSettings> GetAllConfigurations(string documentHash)
        {
            foreach (var attribute in attributes)
            {
                var result = GetCurrentConfiguration(
                    receiverHash: attribute.ReceiverHash,
                    documentHash: documentHash);

                yield return result;
            }
        }

        private TaskPaneSettings GetCurrentConfiguration(string receiverHash, string documentHash)
        {
            var currentAttribute = attributes
                .FirstOrDefault(a => a.ReceiverHash == receiverHash);

            if (currentAttribute == default)
            {
                throw new ApplicationException("There is no respective Prism Task Pane defined.");
            }

            var result = configurations?
                .Where(s => s.ReceiverHash == receiverHash)
                .Where(s => s.DocumentHash == documentHash).LastOrDefault();

            if (result == default)
            {
                result = new TaskPaneSettings
                {
                    ReceiverHash = receiverHash,
                    DockPosition = currentAttribute.DockPosition,
                    DocumentHash = documentHash,
                    Height = currentAttribute.Height,
                    Visible = currentAttribute.Visible,
                    Width = currentAttribute.Width,
                };

                configurations.Add(result);
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

        private IEnumerable<TaskPaneSettings> ReadConfigurations()
        {
            if (File.Exists(configurationsPath))
            {
                var configurations = Enumerable.Empty<TaskPaneSettings>();

                using (var file = new FileStream(
                    path: configurationsPath,
                    mode: FileMode.Open))
                {
                    using (var reader = XmlReader.Create(file))
                    {
                        var serializer = new XmlSerializer(configurations.GetType());
                        configurations = serializer.Deserialize(reader) as TaskPaneSettings[];
                    }
                }

                if (configurations?.Any() ?? false)
                {
                    foreach (var configuration in configurations)
                    {
                        yield return configuration;
                    }
                }
            }
        }

        private void UpdateConfiguration(string attributeHash, string documentHash, bool visible, int width, int height,
            MsoCTPDockPosition dockPosition)
        {
            var currentConfiguration = GetCurrentConfiguration(
                receiverHash: attributeHash,
                documentHash: documentHash);

            currentConfiguration.DockPosition = dockPosition;
            currentConfiguration.Height = height;
            currentConfiguration.DocumentHash = documentHash;
            currentConfiguration.Visible = visible;
            currentConfiguration.Width = width;
        }

        private void WriteConfigurations()
        {
            var relevantConfigurations = configurations
                .GroupBy(s => new { s.ReceiverHash, s.DocumentHash })
                .Select(g => g.Last()).ToArray();

            if (!Directory.Exists(Path.GetDirectoryName(configurationsPath)))
                Directory.CreateDirectory(Path.GetDirectoryName(configurationsPath));

            var writer = new XmlSerializer(relevantConfigurations.GetType());
            using (var file = File.Create(configurationsPath))
            {
                writer.Serialize(
                    stream: file,
                    o: relevantConfigurations);
            }
        }

        #endregion Private Methods
    }
}

#pragma warning restore CA1031 // Keine allgemeinen Ausnahmetypen abfangen