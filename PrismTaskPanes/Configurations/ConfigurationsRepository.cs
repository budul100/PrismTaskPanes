using NetOffice.OfficeApi.Enums;
using PrismTaskPanes.Attributes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace PrismTaskPanes.Configurations
{
    internal class ConfigurationsRepository
    {
        #region Private Fields

        private readonly IList<PrismTaskPaneAttribute> attributes = new List<PrismTaskPaneAttribute>();
        private readonly IList<Configuration> configurations;

        private readonly string configurationsPath;

        #endregion Private Fields

        #region Public Constructors

        public ConfigurationsRepository(string configurationsPath)
        {
            this.configurationsPath = configurationsPath;
            configurations = ReadConfigurations().ToList();
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

        public IEnumerable<Configuration> Get(int documentHash)
        {
            var result = GetAllConfigurations(documentHash).ToArray();

            return result;
        }

        public Configuration Get(int receiverHash, int documentHash)
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

        public void Set(int receiverHash, int documentHash, bool visible,
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

        private IEnumerable<Configuration> GetAllConfigurations(int documentHash)
        {
            foreach (var attribute in attributes)
            {
                var result = GetCurrentConfiguration(
                    receiverHash: attribute.ReceiverHash,
                    documentHash: documentHash);

                yield return result;
            }
        }

        private Configuration GetCurrentConfiguration(int receiverHash, int documentHash)
        {
            var currentAttribute = attributes
                .FirstOrDefault(a => a.ReceiverHash == receiverHash);

            if (currentAttribute == null) throw new ApplicationException(
                $"There is no respective Prism Task Pane defined.");

            var result = configurations?
                .Where(s => s.ReceiverHash == receiverHash)
                .Where(s => s.DocumentHash == documentHash).LastOrDefault();

            if (result == null)
            {
                result = new Configuration
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

        private IEnumerable<Configuration> ReadConfigurations()
        {
            if (File.Exists(configurationsPath))
            {
                var configurations = Array.Empty<Configuration>();
                var serializer = new XmlSerializer(configurations.GetType());

                try
                {
                    using (StreamReader file = new StreamReader(configurationsPath))
                    {
                        configurations = (Configuration[])(serializer.Deserialize(file));
                    }
                }
                catch { }

                if (configurations?.Any() ?? false)
                {
                    foreach (var configuration in configurations)
                    {
                        yield return configuration;
                    }
                }
            }
        }

        private void UpdateConfiguration(int attributeHash, int documentHash, bool visible, int width, int height,
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