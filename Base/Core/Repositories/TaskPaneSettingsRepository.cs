using PrismTaskPanes.Attributes;
using PrismTaskPanes.Enums;
using PrismTaskPanes.Exceptions;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Serialization;

namespace PrismTaskPanes.Settings
{
    [ComVisible(false)]
    public class TaskPaneSettingsRepository
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
            configurations = ReadConfigurations();
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

        public void Set(string receiverHash, string documentHash, bool visible, int width, int height,
            DockPosition dockPosition)
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
                throw new ConfigurationMissingException();
            }

            var result = configurations?
                .Where(s => s.ReceiverHash == receiverHash)
                .LastOrDefault(s => s.DocumentHash == documentHash);

            if (result == default)
            {
                result = new TaskPaneSettings
                {
                    DockPosition = currentAttribute.DockPosition,
                    DocumentHash = documentHash,
                    Height = currentAttribute.Height,
                    ReceiverHash = receiverHash,
                    Visible = currentAttribute.Visible,
                    Width = currentAttribute.Width,
                };

                configurations.Add(result);
            }

            result.DockRestriction = currentAttribute.DockRestriction;
            result.InvisibleAtStart = currentAttribute.InvisibleAtStart;
            result.NavigationValue = currentAttribute.NavigationValue;
            result.RegionContext = currentAttribute.RegionContext;
            result.RegionName = currentAttribute.RegionName;
            result.ScrollBarHorizontal = currentAttribute.ScrollBarHorizontal;
            result.ScrollBarVertical = currentAttribute.ScrollBarVertical;
            result.Title = currentAttribute.Title;
            result.View = currentAttribute.View;

            return result;
        }

        private IList<TaskPaneSettings> ReadConfigurations()
        {
            var result = default(IList<TaskPaneSettings>);

            if (File.Exists(configurationsPath))
            {
                using var file = new FileStream(
                    path: configurationsPath,
                    mode: FileMode.Open);

                using var reader = XmlReader.Create(file);

                try
                {
                    var serializer = new XmlSerializer(typeof(List<TaskPaneSettings>));
                    result = serializer.Deserialize(reader) as List<TaskPaneSettings>;
                }
                catch
                { }
            }

            return result ?? new List<TaskPaneSettings>();
        }

        private void UpdateConfiguration(string attributeHash, string documentHash, bool visible, int width, int height,
            DockPosition dockPosition)
        {
            var currentConfiguration = GetCurrentConfiguration(
                receiverHash: attributeHash,
                documentHash: documentHash);

            currentConfiguration.DockPosition = dockPosition;
            currentConfiguration.DocumentHash = documentHash;
            currentConfiguration.Height = height;
            currentConfiguration.Visible = visible;
            currentConfiguration.Width = width;
        }

        private void WriteConfigurations()
        {
            var relevantConfigurations = configurations
                .GroupBy(s => (s.ReceiverHash, s.DocumentHash))
                .Select(g => g.Last()).ToArray();

            if (!Directory.Exists(Path.GetDirectoryName(configurationsPath)))
            {
                _ = Directory.CreateDirectory(Path.GetDirectoryName(configurationsPath));
            }

            var writer = new XmlSerializer(relevantConfigurations.GetType());

            using var file = File.Create(configurationsPath);

            writer.Serialize(
                stream: file,
                o: relevantConfigurations);
        }

        #endregion Private Methods
    }
}