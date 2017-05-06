using Google.Apis.Analytics.v3;
using Google.Apis.Analytics.v3.Data;
using Google.Apis.Services;
using System.Security.Cryptography.X509Certificates;
using Google.Apis.Auth.OAuth2;
using System.Collections.Generic;
using System.Linq;
using System;

namespace SSAddin
{
    // AnalyticDataPoint encapsulates a result set from Google Analytics
    public class AnalyticDataPoint {
        public AnalyticDataPoint( )
        {
            Rows = new List<IList<string>>();
        }
        public IList<GaData.ColumnHeadersData> ColumnHeaders { get; set; }
        public List<IList<string>> Rows { get; set; }
    }

    public class GoogleAnalyticsAPI
    {
        protected static AnalyticsService s_Service { get; set; }
        protected static IList<Profile> s_Profiles { get; set; }
        protected static GoogleAnalyticsAPI s_Instance;

        #region Worker thread methods

        // Yes, this singleton instance method is a bit clunky with it's parameters. Bear in mind
        // that keypath and email come from ConfigSheet, and all the ConfigSheet invocations happen
        // on the Excel thread. But all the GoogleAnalyticsAPI methods execute on the worker thread.
        // So to get rid of these parms we'd have to introduce locking like DataCache or SSWebClient.
        // From a threading perspective it's cleaner and simpler to pass keypath and email in from
        // every invocation of s2ganalytics on the Excel thread via the request dictionary and have
        // the worker thread fire this method after getting the incoming request.
        public static GoogleAnalyticsAPI Instance( string keypath, string email)
        {
            if (s_Instance == null) {
                Logr.Log(String.Format("GoogleAnalyticsAPI.Instance keypath({0}) email({1})", keypath, email));
                s_Instance = new GoogleAnalyticsAPI( keypath, email );
                var response = s_Service.Management.Profiles.List("~all", "~all").Execute();
                s_Profiles = response.Items;
            }
            return s_Instance;
        }

        protected GoogleAnalyticsAPI(string keyPath, string accountEmailAddress)
        {
            // "notasecret" is the default password Google supplies when you generate a P12 for a service account
            var certificate = new X509Certificate2(keyPath, "notasecret", X509KeyStorageFlags.Exportable);
            var credentials = new ServiceAccountCredential(
               new ServiceAccountCredential.Initializer(accountEmailAddress) {
                   Scopes = new[] { AnalyticsService.Scope.AnalyticsReadonly }
               }.FromCertificate(certificate));
            s_Service = new AnalyticsService(new BaseClientService.Initializer() {
                    HttpClientInitializer = credentials,
                    ApplicationName = "WorthlessVariable"
                });
        }

 
        public AnalyticDataPoint GetAnalyticsData(  string dimensions,  // comma separated list, no spaces
                                                    string metrics,     // comma separated list, no spaces
                                                    string startDate,   // yyyy-MM-dd, possibly supplied by s2today
                                                    string endDate)     // yyyy-MM-dd, possibly supplied by s2today
        {
            // TODO: add another SSAddin.xml.config key g2analytics.profile to select profile
            string profileId = s_Profiles[0].Id;
            if ( !profileId.Contains("ga:"))
                profileId = string.Format("ga:{0}", profileId);

            // Make initial call to service. Then check if a next link exists in the response,
            // if so parse and call again using start index param.
            GaData response = null;
            AnalyticDataPoint data = new AnalyticDataPoint();
            do {
                int startIndex = 1;
                if (response != null && !string.IsNullOrEmpty(response.NextLink)) {
                    Uri uri = new Uri(response.NextLink);
                    var parameters = uri.Query.Split('&');
                    string s = parameters.First(i => i.Contains("start-index")).Split('=')[1];
                    startIndex = int.Parse(s);
                }
                DataResource.GaResource.GetRequest request = s_Service.Data.Ga.Get( profileId, startDate, endDate, metrics);
                request.Dimensions = dimensions;
                response = request.Execute();
                data.ColumnHeaders = response.ColumnHeaders;
                data.Rows.AddRange(response.Rows);
            } while (!string.IsNullOrEmpty(response.NextLink));
            return data;
        }

        #endregion
    }
}