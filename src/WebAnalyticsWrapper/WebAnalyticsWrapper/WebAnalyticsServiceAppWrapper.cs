using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Microsoft.Office.Server.WebAnalytics.ProcessedDataRetriever;
using Microsoft.Office.Server.WebAnalytics.Reporting;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace WebAnalyticsWrapper
{
    /// <summary>
    /// Wraps calls to the FrontEndDataRetriever of the SharePoint 2010 WebAnalytics ServiceApplication
    /// see http://msdn.microsoft.com/en-us/library/ff385276(office.12).aspx for the docs to the reports & parameters
    /// </summary>
    public class WebAnalyticsServiceAppWrapper
    {
        private readonly AggregationContext aggregationContext;

        public WebAnalyticsServiceAppWrapper(SPWeb web)
        {
            aggregationContext = AggregationContext.GetContext(web);
        }
        public WebAnalyticsServiceAppWrapper(SPSite site)
        {
            aggregationContext = AggregationContext.GetContext(site);
        }
        public WebAnalyticsServiceAppWrapper(SPWebApplication webApp)
        {
            aggregationContext = AggregationContext.GetContext(webApp);
        }

        public DataTable GetGenericReport(string reportName, IDictionary<string, object> reportParams)
        {
            var viewParamsList = GetViewParamsList(reportParams);
            var dataPacket = FrontEndDataRetriever.QueryData(
                aggregationContext,
                null, // Reserved. MUST be NULL.
                reportName,
                viewParamsList,
                null, // TODO: Where
                null, // TODO: sortORder 
                1, int.MaxValue, // only used, when at least one sortOrder is set 
                false);

            if (dataPacket == null)
            {
                return null;
            }
            return dataPacket.DataTable;
        }

        #region GetBestBetSuggestions
        /*
         * TODO: fn_WA_GetBestBetSuggestions
         * Query and URL best bets recommended by the protocol
         */

        #endregion

        #region GetBestBetUsage
        /*
         * TODO: fn_WA_GetBestBetUsage
         * Best bet queries, query frequency, best bet URL, best bet URL click frequency and percentage of clicks of best bet URL vs. overall clicks.
         */
        #endregion

        #region GetClickthroughChanges
        /*
         * TODO: fn_WA_GetClickthroughChanges
         * Pages most visited along with their previous rank and current and previous frequency
         */

        #endregion

        #region GetInventory
        /*
         * TODO: fn_WA_GetInventory
         * Top site templates, product versions, languages, list templates
         */

        #endregion

        #region GetInventoryPerDay
        /*
         * TODO: fn_WA_GetInventoryPerDay
         * Number of sites (2), site collections, lists, libraries and storage size per day
         */

        #endregion

        #region GetLast24HourClickthroughChanges
        /*
         * TODO: fn_WA_GetLast24HourClickthroughChanges
         * Pages most visited in the last 24 hours along with their previous rank and current and previous frequency
         */
        #endregion

        #region GetLast24HourSearchQueryChanges
        /*
         * TODO: fn_WA_GetLast24HourSearchQueryChanges
         * Search queries most issued in the last 24 hours along with their previous rank and current and previous frequency
         */
        #endregion

        #region GetLast24HourUserDepartments
        /*
         * TODO: fn_WA_GetLast24HourUserDepartments
         * User departments logged in the last 24 hours. User department is the organizational department information of a user as stored in profile database of profile service.
         */
        #endregion

        #region GetLast24HourUserTitles
        /*
         * TODO: fn_WA_GetLast24HourUserTitles
         * User titles logged in the last 24 hours. User title is the organizational title information of a user as stored in profile database of profile service.
         */
        #endregion

        #region GetNumberOfClickthroughs
        /*
         * TODO: fn_WA_GetNumberOfClickthroughs
         * Total number of page views grouped per day or grouped by URL
         */
        #endregion

        #region GetNumberOfFailedSearchQueriesPerDay
        /*
         * TODO: fn_WA_GetNumberOfFailedSearchQueriesPerDay
         * Total number of queries per day that didn’t give satisfactory results. A query gives unsatisfactory results when it gives no results or the results it returns get little or no clicks.
         */
        #endregion

        #region GetNumberOfSearchQueries
        /*
         * TODO: fn_WA_GetNumberOfSearchQueries
         * Total number of search queries grouped per day or grouped by search query
         */
        #endregion

        #region GetNumberOfSearchQueriesPerDay
        /*
         * TODO: fn_WA_GetNumberOfSearchQueriesPerDay
         * Total number of search queries per day
         */
        #endregion

        #region GetSearchQueryChanges
        /*
         * TODO: fn_WA_GetSearchQueryChanges
         * Search queries most issued along with their current and previous frequency and previous rank
         */
        #endregion

        #region GetTopBrowsers
        /*
         * TODO: fn_WA_GetTopBrowsers
         * Top browsers
         */
        #endregion

        #region GetTopDestinations
        /*
         * TODO: fn_WA_GetTopDestinations
         * Top URLs that are outside the entity for which data is being requested and are referred by the entity for which data is being requested. The source and destination entities are the site (2)/ site collection / web application. For example this refers to the scenario when the URLs from a site (2) point to the destination site (2).
         */
        #endregion

        #region GetTopFailedSearchQueries
        /*
         * TODO: fn_WA_GetTopFailedSearchQueries
         * Search queries most issued that didn’t give satisfactory results. A query gives unsatisfactory results when it gives no results or the results it returns get little or no clicks.
         *
         */
        #endregion

        #region GetTopPages
        /*
         * TODO: fn_WA_GetTopPages
         * Pages most visited
         */
        #endregion

        #region GetTopReferrers
        /*
         * TODO: fn_WA_GetTopReferrers
         * Top URLs that are outside the entity for which data is being requested and refer the entity for which data is being requested
         */
        #endregion

        #region GetTopSearchQueries
        /*
         * TODO: fn_WA_GetTopSearchQueries
         * Search queries most issued
         */
        #endregion

        #region GetTopVisitors
        /*
         * TODO: fn_WA_GetTopVisitors
         * Top visitors
         */
        #endregion

        #region GetUserDepartments
        /*
         * TODO: fn_WA_GetUserDepartments
         * User department names. User department is the organizational department information of a user as stored in profile database of profile service.
         */
        #endregion

        #region GetUserTitles
        /*
         * TODO: fn_WA_GetUserTitles
         * User titles. User title  is the organizational title information of a user as stored in profile database of profile service.
         */
        #endregion

        #region GetSummary
        private DataTable GetSummary(DateTime currentPeriodStartDate, DateTime previousPeriodStartDate, int days, bool? includeSubWebs)
        {
            var parms = new Dictionary<string, object>
            {
                { "CurrentStartDateId", GetDateId(currentPeriodStartDate) },
                { "PreviousStartDateId", GetDateId(previousPeriodStartDate) },
                { "Duration", days }
            };
            if (includeSubWebs.HasValue)
            {
                parms.Add("IncludeSubSites", includeSubWebs.Value);
            }
            return GetGenericReport("fn_WA_GetSummary", parms);
        }
        /// <summary>
        /// Summary report for the entity comprising of Traffic, Search and Inventory Data
        /// TODO: Docs
        /// </summary>
        /// <param name="currentPeriodStartDate"></param>
        /// <param name="previousPeriodStartDate"></param>
        /// <param name="days"></param>
        /// <returns></returns>
        public DataTable GetSummary(DateTime currentPeriodStartDate, DateTime previousPeriodStartDate, int days)
        {
            return GetSummary(currentPeriodStartDate, previousPeriodStartDate, days, null);
        }

        /// <summary>
        /// Summary report for the entity comprising of Traffic, Search and Inventory Data
        /// TODO: Docs
        /// </summary>
        /// <param name="currentPeriodStartDate"></param>
        /// <param name="previousPeriodStartDate"></param>
        /// <param name="days"></param>
        /// <param name="includeSubWebs"></param>
        /// <returns></returns>
        public DataTable GetSummary(DateTime currentPeriodStartDate, DateTime previousPeriodStartDate, int days, bool includeSubWebs)
        {
            return GetSummary(currentPeriodStartDate, previousPeriodStartDate, days, (bool?)includeSubWebs);
        }
        #endregion

        #region GetTrafficVolumePerDay
        private DataTable GetTrafficVolumePerDay(DateTime startDate, DateTime endDate, TrafficVolumeMetricType type, bool? includeSubWebs)
        {
            var parms = new Dictionary<string, object>
            {
                {"StartDateId", GetDateId(startDate)},
                {"EndDateId", GetDateId(endDate)},
                {"MetricType", (int) type},
            };
            if (includeSubWebs.HasValue)
            {
                parms.Add("IncludeSubSites", includeSubWebs.Value);
            }

            return GetGenericReport("fn_WA_GetTrafficVolumePerDay", parms);
        }

        /// <summary>
        /// Page views per day
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="includeSubWebs"></param>
        /// <returns></returns>
        public DataTable GetPageViewsTrafficVolumePerDay(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.PageViews, includeSubWebs);
        }

        /// <summary>
        /// Page views per day
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public DataTable GetPageViewsTrafficVolumePerDay(DateTime startDate, DateTime endDate)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.PageViews, null);
        }

        /// <summary>
        /// Page views per day
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="includeSubWebs"></param>
        /// <returns></returns>
        public DataTable GetExternalDestinationsTrafficVolumePerDay(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.ExternalDestinations, includeSubWebs);
        }

        /// <summary>
        /// Page views per day
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public DataTable GetExternalDestinationsTrafficVolumePerDay(DateTime startDate, DateTime endDate)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.ExternalDestinations, null);
        }

        /// <summary>
        /// Page views per day
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="includeSubWebs"></param>
        /// <returns></returns>
        public DataTable GetExternalReferrersTrafficVolumePerDay(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.ExternalReferrers, includeSubWebs);
        }

        /// <summary>
        /// Page views per day
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public DataTable GetExternalReferrersTrafficVolumePerDay(DateTime startDate, DateTime endDate)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.ExternalReferrers, null);
        }

        /// <summary>
        /// Page views per day
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="includeSubWebs"></param>
        /// <returns></returns>
        public DataTable GetSearchQueriesTrafficVolumePerDay(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.SearchQueries, includeSubWebs);
        }

        /// <summary>
        /// Page views per day
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public DataTable GetSearchQueriesTrafficVolumePerDay(DateTime startDate, DateTime endDate)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.SearchQueries, null);
        }

        /// <summary>
        /// Page views per day
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="includeSubWebs"></param>
        /// <returns></returns>
        public DataTable GetUniqueVisitorsTrafficVolumePerDay(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.UniqueVisitors, includeSubWebs);
        }

        /// <summary>
        /// Page views per day
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public DataTable GetUniqueVisitorsTrafficVolumePerDay(DateTime startDate, DateTime endDate)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.UniqueVisitors, null);
        }
        #endregion

        #region GetTotalTrafficVolume
        private DataTable GetTotalTrafficVolume(DateTime startDate, DateTime endDate, TrafficVolumeMetricType type, bool? includeSubWebs)
        {
            var parms = new Dictionary<string, object>
            {
                {"StartDateId", GetDateId(startDate)},
                {"EndDateId", GetDateId(endDate)},
                {"MetricType", (int) type},
            };
            if (includeSubWebs.HasValue)
            {
                parms.Add("IncludeSubSites", includeSubWebs.Value);
            }

            return GetGenericReport("fn_WA_GetTotalTrafficVolume", parms);
        }

        /// <summary>
        /// Total number of page views
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="includeSubWebs"></param>
        /// <returns></returns>
        public DataTable GetPageViewsTotalTrafficVolume(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.PageViews, includeSubWebs);
        }

        /// <summary>
        /// Total number of page views
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public DataTable GetPageViewsTotalTrafficVolume(DateTime startDate, DateTime endDate)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.PageViews, null);
        }

        /// <summary>
        /// Total number of page views
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="includeSubWebs"></param>
        /// <returns></returns>
        public DataTable GetExternalDestinationsTotalTrafficVolume(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.ExternalDestinations, includeSubWebs);
        }

        /// <summary>
        /// Total number of page views
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public DataTable GetExternalDestinationsTotalTrafficVolume(DateTime startDate, DateTime endDate)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.ExternalDestinations, null);
        }

        /// <summary>
        /// Total number of page views
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="includeSubWebs"></param>
        /// <returns></returns>
        public DataTable GetExternalReferrersTotalTrafficVolume(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.ExternalReferrers, includeSubWebs);
        }

        /// <summary>
        /// Total number of page views
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public DataTable GetExternalReferrersTotalTrafficVolume(DateTime startDate, DateTime endDate)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.ExternalReferrers, null);
        }

        /// <summary>
        /// Total number of page views
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="includeSubWebs"></param>
        /// <returns></returns>
        public DataTable GetSearchQueriesTotalTrafficVolume(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.SearchQueries, includeSubWebs);
        }

        /// <summary>
        /// Total number of page views
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public DataTable GetSearchQueriesTotalTrafficVolume(DateTime startDate, DateTime endDate)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.SearchQueries, null);
        }

        /// <summary>
        /// Total number of page views
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="includeSubWebs"></param>
        /// <returns></returns>
        public DataTable GetUniqueVisitorsTotalTrafficVolume(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.UniqueVisitors, includeSubWebs);
        }

        /// <summary>
        /// Total number of page views
        /// TODO: Docs
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public DataTable GetUniqueVisitorsTotalTrafficVolume(DateTime startDate, DateTime endDate)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.UniqueVisitors, null);
        }

        #endregion

        #region Helpers

        private List<ViewParameterValue> GetViewParamsList(IDictionary<string, object> reportParams)
        {
            return reportParams.Select(kvp => new ViewParameterValue(kvp.Key, kvp.Value)).ToList();
        }

        private int GetDateId(DateTime date)
        {
            return ((date.Year*100) + date.Month)*100 + date.Day;
        }

        #endregion
    }
}
