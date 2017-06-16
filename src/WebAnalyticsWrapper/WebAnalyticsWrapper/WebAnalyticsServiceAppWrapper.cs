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
        public DataTable GetSummary(DateTime currentPeriodStartDate, DateTime previousPeriodStartDate, int days)
        {
            return GetSummary(currentPeriodStartDate, previousPeriodStartDate, days, null);
        }
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

        public DataTable GetPageViewsTrafficVolumePerDay(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.PageViews, includeSubWebs);
        }
        public DataTable GetPageViewsTrafficVolumePerDay(DateTime startDate, DateTime endDate)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.PageViews, null);
        }
        public DataTable GetExternalDestinationsTrafficVolumePerDay(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.ExternalDestinations, includeSubWebs);
        }
        public DataTable GetExternalDestinationsTrafficVolumePerDay(DateTime startDate, DateTime endDate)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.ExternalDestinations, null);
        }
        public DataTable GetExternalReferrersTrafficVolumePerDay(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.ExternalReferrers, includeSubWebs);
        }
        public DataTable GetExternalReferrersTrafficVolumePerDay(DateTime startDate, DateTime endDate)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.ExternalReferrers, null);
        }
        public DataTable GetSearchQueriesTrafficVolumePerDay(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.SearchQueries, includeSubWebs);
        }
        public DataTable GetSearchQueriesTrafficVolumePerDay(DateTime startDate, DateTime endDate)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.SearchQueries, null);
        }
        public DataTable GetUniqueVisitorsTrafficVolumePerDay(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTrafficVolumePerDay(startDate, endDate, TrafficVolumeMetricType.UniqueVisitors, includeSubWebs);
        }
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
        public DataTable GetPageViewsTotalTrafficVolume(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.PageViews, includeSubWebs);
        }
        public DataTable GetPageViewsTotalTrafficVolume(DateTime startDate, DateTime endDate)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.PageViews, null);
        }
        public DataTable GetExternalDestinationsTotalTrafficVolume(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.ExternalDestinations, includeSubWebs);
        }
        public DataTable GetExternalDestinationsTotalTrafficVolume(DateTime startDate, DateTime endDate)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.ExternalDestinations, null);
        }
        public DataTable GetExternalReferrersTotalTrafficVolume(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.ExternalReferrers, includeSubWebs);
        }
        public DataTable GetExternalReferrersTotalTrafficVolume(DateTime startDate, DateTime endDate)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.ExternalReferrers, null);
        }
        public DataTable GetSearchQueriesTotalTrafficVolume(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.SearchQueries, includeSubWebs);
        }
        public DataTable GetSearchQueriesTotalTrafficVolume(DateTime startDate, DateTime endDate)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.SearchQueries, null);
        }
        public DataTable GetUniqueVisitorsTotalTrafficVolume(DateTime startDate, DateTime endDate, bool includeSubWebs)
        {
            return GetTotalTrafficVolume(startDate, endDate, TrafficVolumeMetricType.UniqueVisitors, includeSubWebs);
        }
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
