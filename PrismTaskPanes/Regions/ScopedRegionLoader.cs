using CommonServiceLocator;
using Prism.Regions;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PrismTaskPanes.Regions
{
    internal class ScopedRegionLoader :
        RegionNavigationContentLoader, IRegionNavigationContentLoader
    {
        #region Private Fields

        private readonly IServiceLocator serviceLocator;

        #endregion Private Fields

        #region Public Constructors

        public ScopedRegionLoader(IServiceLocator serviceLocator) :
            base(serviceLocator)
        {
            this.serviceLocator = serviceLocator;
        }

        #endregion Public Constructors

        #region Public Methods

        public new object LoadContent(IRegion region, NavigationContext navigationContext)
        {
            if (region == null)
                throw new ArgumentNullException(nameof(region));
            if (navigationContext == null)
                throw new ArgumentNullException(nameof(navigationContext));

            var candidateTargetContract = GetContractFromNavigationContext(navigationContext);

            var candidates = GetCandidatesFromRegion(
                region: region,
                candidateNavigationContract: candidateTargetContract);

            var result = candidates
                .Where((v) => v.IsCandidate(navigationContext))
                .FirstOrDefault();

            if (result == default)
            {
                result = CreateNewRegionItem(candidateTargetContract);

                // Check if scoped region is required

                var info = (result as IScopedRegion)
                    ?? result.GetScopedRegionInfo();

                if (info != default)
                {
                    region.Add(
                        view: result,
                        viewName: info.ViewName,
                        createRegionManagerScope: info.CreateRegionManagerScope);
                }
                else
                {
                    region.Add(result);
                }
            }

            return result;
        }

        #endregion Public Methods

        #region Protected Methods

        /// <summary>
        /// Returns the set of candidates that may satisfiy this navigation request.
        /// </summary>
        /// <param name="region">The region containing items that may satisfy the navigation request.</param>
        /// <param name="candidateNavigationContract">The candidate navigation target.</param>
        /// <returns>An enumerable of candidate objects from the <see cref="IRegion"/></returns>
        protected override IEnumerable<object> GetCandidatesFromRegion(IRegion region, string candidateNavigationContract)
        {
            if (string.IsNullOrWhiteSpace(candidateNavigationContract))
                throw new ArgumentNullException(nameof(candidateNavigationContract));

            var result = base.GetCandidatesFromRegion(
                region: region,
                candidateNavigationContract: candidateNavigationContract);

            if (!result.Any())
            {
                // Autofac based implementation
                //
                // First try friendly name registration.
                // var matchingRegistration = context.GetService(typeof(KeyedService)).ComponentRegistry.Registrations
                //    .FirstOrDefault(r => r.Services
                //        .OfType<KeyedService>()
                //        .Any(s => s.ServiceKey.Equals(candidateNavigationContract)));
                //
                // If not found, try type registration
                // if (matchingRegistration == null)
                //    matchingRegistration = context.ComponentRegistry.Registrations
                //        .FirstOrDefault(r => candidateNavigationContract
                //            .Equals(r.Activator.LimitType.Name, StringComparison.Ordinal));

                var matchingInstances = serviceLocator
                    .GetInstance<object>(candidateNavigationContract);

                if (matchingInstances == null)
                    return Array.Empty<object>();

                var typeCandidateName = matchingInstances.GetType().FullName;

                result = base.GetCandidatesFromRegion(
                    region: region,
                    candidateNavigationContract: typeCandidateName);
            }

            return result;
        }

        #endregion Protected Methods
    }
}