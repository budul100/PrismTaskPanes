using Prism.Ioc;
using Prism.Regions;
using PrismTaskPanes.Exceptions;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace PrismTaskPanes.Regions
{
    [ComVisible(false)]
    public class ScopedRegionLoader
        : RegionNavigationContentLoader, IRegionNavigationContentLoader
    {
        #region Private Fields

        private readonly IContainerExtension container;

        #endregion Private Fields

        #region Public Constructors

        public ScopedRegionLoader(IContainerExtension container)
            : base(container)
        {
            this.container = container;
        }

        #endregion Public Constructors

        #region Public Methods

        public new object LoadContent(IRegion region, NavigationContext navigationContext)
        {
            if (region == default)
            {
                throw new ArgumentNullException(nameof(region));
            }

            if (navigationContext == default)
            {
                throw new ArgumentNullException(nameof(navigationContext));
            }

            var candidateTargetContract = GetContractFromNavigationContext(navigationContext);

            var candidates = GetCandidatesFromRegion(
                region: region,
                candidateNavigationContract: candidateTargetContract).ToArray();

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
            if (region == default)
            {
                throw new ArgumentNullException(nameof(region));
            }

            if (string.IsNullOrEmpty(candidateNavigationContract))
            {
                throw new ArgumentNullException(nameof(candidateNavigationContract));
            }

            var result = base.GetCandidatesFromRegion(
                region: region,
                candidateNavigationContract: candidateNavigationContract);

            if (!result.Any())
            {
                var matchingInstances = container
                    .Resolve<object>(candidateNavigationContract);

                if (matchingInstances == default)
                    return Enumerable.Empty<object>();

                var typeCandidateName = matchingInstances.GetType().FullName;

                try
                {
                    result = base.GetCandidatesFromRegion(
                        region: region,
                        candidateNavigationContract: typeCandidateName);
                }
                catch (NullReferenceException)
                {
                    throw new RegionNotLoadedException(
                        region: region);
                }
            }

            return result;
        }

        #endregion Protected Methods
    }
}