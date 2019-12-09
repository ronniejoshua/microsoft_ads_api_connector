import webbrowser
from bingads.service_client import ServiceClient
from bingads.authorization import AuthorizationData, OAuthDesktopMobileAuthCodeGrant
from bingads.v13.reporting import *
from suds import WebFault
import time
import datetime as dt

REFERENCE = """
    Bing Ads API Client Libraries
    https://docs.microsoft.com/en-us/advertising/guides/client-libraries?view=bingads-13

    Microsoft Advertising UI:
    https://ads.microsoft.com/
    
    Microsoft Advertising Getting a Developer Token
    https://developers.ads.microsoft.com/Account

    To Register an Application & Get Client ID
    https://portal.azure.com/

    Reference For this Code is Taken From the following:
    ----------------------------------------------------
    Walkthrough: Bing Ads API Desktop Application in Python
    https://docs.microsoft.com/en-us/advertising/guides/walkthrough-desktop-application-python?view=bingads-13

    Report Requests Code Example
    https://docs.microsoft.com/en-us/advertising/guides/code-example-report-requests?view=bingads-13

    Creating Other Report 
    https://docs.microsoft.com/en-us/advertising/guides/report-types?view=bingads-13

"""

class MicrosoftAdsAPI(object):
    def __init__(self, client_id, developer_token, environment, refresh_token, client_state):
        self.CLIENT_ID = client_id
        self.DEVELOPER_TOKEN = developer_token
        self.ENVIRONMENT = environment
        self.REFRESH_TOKEN = refresh_token
        self.CLIENT_STATE = client_state
        self.REPORT_FILE_FORMAT = 'Csv'
        if os.name == 'posix':
            self.FILE_DIRECTORY = r'./ms_ads/files'
        else:
            self.FILE_DIRECTORY = r''
        self.TIMEOUT_IN_MILLISECONDS = 3600000

    def authenticate(self, authorization_data):
        customer_service = ServiceClient(
            service='CustomerManagementService',
            version=13,
            authorization_data=authorization_data,
            environment=self.ENVIRONMENT,
        )

        # You should authenticate for Bing Ads API service operations with a Microsoft Account.
        self.authenticate_with_oauth(authorization_data)

        # Set to an empty user identifier to get the current authenticated Microsoft Advertising user,
        # and then search for all accounts the user can access.
        user = get_user_response = customer_service.GetUser(UserId=None).User
        accounts = self.search_accounts_by_user_id(customer_service, user.Id)

        # Custom Added Code
        account_ids = [{k.Id: k.Name} for k in accounts['AdvertiserAccount']]
        print(account_ids)

        # For this example we'll use the first account.
        # authorization_data.account_id = accounts['AdvertiserAccount'][0].Id
        # authorization_data.customer_id = accounts['AdvertiserAccount'][0].ParentCustomerId
        return [k.Id for k in accounts['AdvertiserAccount']]

    def authenticate_with_oauth(self, authorization_data):
        authentication = OAuthDesktopMobileAuthCodeGrant(
            client_id=self.CLIENT_ID,
            env=self.ENVIRONMENT
        )

        # It is recommended that you specify a non guessable 'state' request parameter to help prevent
        # cross site request forgery (CSRF).
        authentication.state = self.CLIENT_STATE

        # Assign this authentication instance to the authorization_data.
        authorization_data.authentication = authentication

        # Register the callback function to automatically save the refresh token anytime it is refreshed.
        # Uncomment this line if you want to store your refresh token. Be sure to save your refresh token securely.
        authorization_data.authentication.token_refreshed_callback = self.save_refresh_token

        refresh_token = self.get_refresh_token()

        try:
            # If we have a refresh token let's refresh it
            if refresh_token is not None:
                authorization_data.authentication.request_oauth_tokens_by_refresh_token(refresh_token)
            else:
                self.request_user_consent(authorization_data)
        except OAuthTokenRequestException:
            # The user could not be authenticated or the grant is expired.
            # The user must first sign in and if needed grant the client application access to the requested scope.
            self.request_user_consent(authorization_data)

    def request_user_consent(self, authorization_data):
        webbrowser.open(authorization_data.authentication.get_authorization_endpoint(), new=1)
        # For Python 3.x use 'input' instead of 'raw_input'
        if sys.version_info.major >= 3:
            response_uri = input(
                "You need to provide consent for the application to access your Microsoft Advertising accounts. "
                "After you have granted consent in the web browser for the application to access your Microsoft "
                "Advertising accounts, "
                "please enter the response URI that includes the authorization 'code' parameter: \n"
            )
        else:
            response_uri = raw_input(
                "You need to provide consent for the application to access your Microsoft Advertising accounts. "
                "After you have granted consent in the web browser for the application to access your Microsoft "
                "Advertising accounts, "
                "please enter the response URI that includes the authorization 'code' parameter: \n"
            )

        if authorization_data.authentication.state != self.CLIENT_STATE:
            raise Exception("The OAuth response state does not match the client request state.")

        # Request access and refresh tokens using the URI that you provided manually during program execution.
        authorization_data.authentication.request_oauth_tokens_by_response_uri(response_uri=response_uri)

    def get_refresh_token(self):
        """
        Returns a refresh token if found.
        """
        file = None
        try:
            file = open(self.REFRESH_TOKEN)
            line = file.readline()
            file.close()
            return line if line else None
        except IOError:
            if file:
                file.close()
            return None

    def save_refresh_token(self, oauth_tokens):
        """
        Stores a refresh token locally. Be sure to save your refresh token securely.
        """
        with open(self.REFRESH_TOKEN, "w+") as file:
            file.write(oauth_tokens.refresh_token)
            file.close()
        return None

    def search_accounts_by_user_id(self, customer_service, user_id):
        predicates = {
            'Predicate': [
                {
                    'Field': 'UserId',
                    'Operator': 'Equals',
                    'Value': user_id,
                },
            ]
        }

        accounts = []

        page_index = 0
        PAGE_SIZE = 100
        found_last_page = False

        while not found_last_page:
            paging = self.set_elements_to_none(customer_service.factory.create('ns5:Paging'))
            paging.Index = page_index
            paging.Size = PAGE_SIZE
            search_accounts_response = customer_service.SearchAccounts(
                PageInfo=paging,
                Predicates=predicates
            )

            if search_accounts_response is not None and hasattr(search_accounts_response, 'AdvertiserAccount'):
                accounts.extend(search_accounts_response['AdvertiserAccount'])
                found_last_page = PAGE_SIZE > len(search_accounts_response['AdvertiserAccount'])
                page_index += 1
            else:
                found_last_page = True

        return {
            'AdvertiserAccount': accounts
        }

    @staticmethod
    def set_elements_to_none(suds_object):
        for (element) in suds_object:
            suds_object.__setitem__(element[0], None)
        return suds_object

    @staticmethod
    def output_status_message(message):
        print(message)

    def output_bing_ads_webfault_error(self, error):
        if hasattr(error, 'ErrorCode'):
            self.output_status_message("ErrorCode: {0}".format(error.ErrorCode))
        if hasattr(error, 'Code'):
            self.output_status_message("Code: {0}".format(error.Code))
        if hasattr(error, 'Details'):
            self.output_status_message("Details: {0}".format(error.Details))
        if hasattr(error, 'FieldPath'):
            self.output_status_message("FieldPath: {0}".format(error.FieldPath))
        if hasattr(error, 'Message'):
            self.output_status_message("Message: {0}".format(error.Message))
        self.output_status_message('')

    def output_webfault_errors(self, ex):
        if not hasattr(ex.fault, "detail"):
            raise Exception("Unknown WebFault")

        error_attribute_sets = (
            ["ApiFault", "OperationErrors", "OperationError"],
            ["AdApiFaultDetail", "Errors", "AdApiError"],
            ["ApiFaultDetail", "BatchErrors", "BatchError"],
            ["ApiFaultDetail", "OperationErrors", "OperationError"],
            ["EditorialApiFaultDetail", "BatchErrors", "BatchError"],
            ["EditorialApiFaultDetail", "EditorialErrors", "EditorialError"],
            ["EditorialApiFaultDetail", "OperationErrors", "OperationError"],
        )

        for error_attribute_set in error_attribute_sets:
            if self.output_error_detail(ex.fault.detail, error_attribute_set):
                return

        # Handle serialization errors, for example: The formatter threw an exception while trying to deserialize the
        # message, etc.
        if hasattr(ex.fault, 'detail') \
                and hasattr(ex.fault.detail, 'ExceptionDetail'):
            api_errors = ex.fault.detail.ExceptionDetail
            if isinstance(api_errors, list):
                for api_error in api_errors:
                    self.output_status_message(api_error.Message)
            else:
                self.output_status_message(api_errors.Message)
            return

        raise Exception("Unknown WebFault")

    def output_error_detail(self, error_detail, error_attribute_set):
        api_errors = error_detail
        for field in error_attribute_set:
            api_errors = getattr(api_errors, field, None)
        if api_errors is None:
            return False
        if isinstance(api_errors, list):
            for api_error in api_errors:
                self.output_bing_ads_webfault_error(api_error)
        else:
            self.output_bing_ads_webfault_error(api_errors)
        return True

    @staticmethod
    def get_budget_summary_report_request(
            _reporting_service,
            account_id,
            exclude_column_headers,
            exclude_report_footer,
            exclude_report_header,
            report_file_format,
            return_only_complete_data,
            report_time):

        report_request = _reporting_service.factory.create('BudgetSummaryReportRequest')
        report_request.ExcludeColumnHeaders = exclude_column_headers
        report_request.ExcludeReportFooter = exclude_report_footer
        report_request.ExcludeReportHeader = exclude_report_header
        report_request.Format = report_file_format
        report_request.ReturnOnlyCompleteData = return_only_complete_data
        report_request.Time = report_time
        report_request.ReportName = "budget_summary_report"
        scope = _reporting_service.factory.create('AccountThroughCampaignReportScope')
        scope.AccountIds = {'long': [account_id]}
        scope.Campaigns = None
        report_request.Scope = scope

        report_columns = _reporting_service.factory.create('ArrayOfBudgetSummaryReportColumn')
        report_columns.BudgetSummaryReportColumn.append([
            'AccountName',
            'AccountNumber',
            'AccountId',
            'CampaignName',
            'CampaignId',
            'Date',
            'CurrencyCode',
            'MonthlyBudget',
            'DailySpend',
            'MonthToDateSpend'
        ])
        report_request.Columns = report_columns
        return report_request

    @staticmethod
    def get_campaign_performance_report_request(
            _reporting_service,
            account_id,
            aggregation,
            exclude_column_headers,
            exclude_report_footer,
            exclude_report_header,
            report_file_format,
            return_only_complete_data,
            report_time):

        report_request = _reporting_service.factory.create('CampaignPerformanceReportRequest')
        report_request.Aggregation = aggregation
        report_request.ExcludeColumnHeaders = exclude_column_headers
        report_request.ExcludeReportFooter = exclude_report_footer
        report_request.ExcludeReportHeader = exclude_report_header
        report_request.Format = report_file_format
        report_request.ReturnOnlyCompleteData = return_only_complete_data
        report_request.Time = report_time
        report_request.ReportName = "campaign_performance_report"
        scope = _reporting_service.factory.create('AccountThroughCampaignReportScope')
        scope.AccountIds = {'long': [account_id]}
        scope.Campaigns = None
        report_request.Scope = scope

        report_columns = _reporting_service.factory.create('ArrayOfCampaignPerformanceReportColumn')
        report_columns.CampaignPerformanceReportColumn.append([
            'TimePeriod',
            'CampaignId',
            'CampaignName',
            'DeviceType',
            'Network',
            'Impressions',
            'Clicks',
            'Spend'
        ])
        report_request.Columns = report_columns
        return report_request

    @staticmethod
    def get_search_query_performance_report_request(
            _reporting_service,
            account_id,
            aggregation,
            exclude_column_headers,
            exclude_report_footer,
            exclude_report_header,
            report_file_format,
            return_only_complete_data,
            report_time):
        """
        Reference:  https://docs.microsoft.com/en-us/advertising/reporting-service/searchqueryperformancereportrequest?view=bingads-13
                    https://docs.microsoft.com/en-us/advertising/reporting-service/searchqueryperformancereportcolumn?view=bingads-13
        """

        report_request = _reporting_service.factory.create('SearchQueryPerformanceReportRequest')
        report_request.Aggregation = aggregation
        report_request.ExcludeColumnHeaders = exclude_column_headers
        report_request.ExcludeReportFooter = exclude_report_footer
        report_request.ExcludeReportHeader = exclude_report_header
        report_request.Format = report_file_format
        report_request.ReturnOnlyCompleteData = return_only_complete_data
        report_request.Time = report_time
        report_request.ReportName = "search_query_performance_report"
        scope = _reporting_service.factory.create('AccountThroughAdGroupReportScope')
        scope.AccountIds = {'long': [account_id]}
        scope.Campaigns = None
        scope.AdGroups = None
        report_request.Scope = scope

        report_columns = _reporting_service.factory.create('ArrayOfSearchQueryPerformanceReportColumn')
        report_columns.SearchQueryPerformanceReportColumn.append([
            "AccountName",
            "AccountNumber",
            "AccountId",
            "TimePeriod",
            "CampaignName",
            "CampaignId",
            "AdGroupName",
            "AdGroupId",
            "AdId",
            "AdType",
            "DestinationUrl",
            "BidMatchType",
            "DeliveredMatchType",
            "CampaignStatus",
            "AdStatus",
            "Impressions",
            "Clicks",
            "AverageCpc",
            "Spend",
            "AveragePosition",
            "SearchQuery",
            "Keyword",
            "AdGroupCriterionId",
            "Conversions",
            "CostPerConversion",
            "Language",
            "KeywordId",
            "Network",
            "TopVsOther",
            "DeviceType",
            "DeviceOS",
            "Assists",
            "Revenue",
            "ReturnOnAdSpend",
            "CostPerAssist",
            "RevenuePerConversion",
            "RevenuePerAssist",
            "AccountStatus",
            "AdGroupStatus",
            "KeywordStatus",
            "CampaignType",
            "CustomerId",
            "CustomerName",
            "AllConversions",
            "AllRevenue",
            "AllCostPerConversion",
            "AllReturnOnAdSpend",
            "AllRevenuePerConversion"
        ])
        report_request.Columns = report_columns
        return report_request

    @staticmethod
    def get_keyword_performance_report_request(
            _reporting_service,
            account_id,
            aggregation,
            exclude_column_headers,
            exclude_report_footer,
            exclude_report_header,
            report_file_format,
            return_only_complete_data,
            report_time):
        """
        Reference:  https://docs.microsoft.com/en-us/advertising/reporting-service/keywordperformancereportrequest?view=bingads-13
                    https://docs.microsoft.com/en-us/advertising/reporting-service/keywordperformancereportcolumn?view=bingads-13
        """
        report_request = _reporting_service.factory.create('KeywordPerformanceReportRequest')
        report_request.Aggregation = aggregation
        report_request.ExcludeColumnHeaders = exclude_column_headers
        report_request.ExcludeReportFooter = exclude_report_footer
        report_request.ExcludeReportHeader = exclude_report_header
        report_request.Format = report_file_format
        report_request.ReturnOnlyCompleteData = return_only_complete_data
        report_request.Time = report_time
        report_request.ReportName = "keyword_performance_report"
        scope = _reporting_service.factory.create('AccountThroughAdGroupReportScope')
        scope.AccountIds = {'long': [account_id]}
        scope.Campaigns = None
        scope.AdGroups = None
        report_request.Scope = scope
        report_columns = _reporting_service.factory.create('ArrayOfKeywordPerformanceReportColumn')
        report_columns.KeywordPerformanceReportColumn.append([
            "AccountName",
            "AccountNumber",
            "AccountId",
            "TimePeriod",
            "CampaignName",
            "CampaignId",
            "AdGroupName",
            "AdGroupId",
            "Keyword",
            "KeywordId",
            "AdId",
            "AdType",
            "DestinationUrl",
            "CurrentMaxCpc",
            "CurrencyCode",
            "DeliveredMatchType",
            "AdDistribution",
            "Impressions",
            "Clicks",
            "AverageCpc",
            "Spend",
            "AveragePosition",
            "Conversions",
            "CostPerConversion",
            "BidMatchType",
            "DeviceType",
            "QualityScore",
            "ExpectedCtr",
            "AdRelevance",
            "LandingPageExperience",
            "Language",
            "HistoricalQualityScore",
            "HistoricalExpectedCtr",
            "HistoricalAdRelevance",
            "HistoricalLandingPageExperience",
            "QualityImpact",
            "CampaignStatus",
            "AccountStatus",
            "AdGroupStatus",
            "KeywordStatus",
            "Network",
            "TopVsOther",
            "DeviceOS",
            "Assists",
            "Revenue",
            "ReturnOnAdSpend",
            "CostPerAssist",
            "RevenuePerConversion",
            "RevenuePerAssist",
            "TrackingTemplate",
            "CustomParameters",
            "FinalUrl",
            "FinalMobileUrl",
            "FinalAppUrl",
            "BidStrategyType",
            "KeywordLabels",
            "Mainline1Bid",
            "MainlineBid",
            "FirstPageBid",
            "FinalUrlSuffix",
            "BaseCampaignId",
            "AllConversions",
            "AllRevenue",
            "AllCostPerConversion",
            "AllReturnOnAdSpend",
            "AllRevenuePerConversion"
        ])
        report_request.Columns = report_columns
        return report_request

    @staticmethod
    def get_user_location_performance_report_request(
            _reporting_service,
            account_id,
            aggregation,
            exclude_column_headers,
            exclude_report_footer,
            exclude_report_header,
            report_file_format,
            return_only_complete_data,
            report_time):
        """
        Reference:  https://docs.microsoft.com/en-us/advertising/reporting-service/userlocationperformancereportrequest?view=bingads-13
                    https://docs.microsoft.com/en-us/advertising/reporting-service/userlocationperformancereportcolumn?view=bingads-13
        """
        report_request = _reporting_service.factory.create('UserLocationPerformanceReportRequest')
        report_request.Aggregation = aggregation
        report_request.ExcludeColumnHeaders = exclude_column_headers
        report_request.ExcludeReportFooter = exclude_report_footer
        report_request.ExcludeReportHeader = exclude_report_header
        report_request.Format = report_file_format
        report_request.ReturnOnlyCompleteData = return_only_complete_data
        report_request.Time = report_time
        report_request.ReportName = "user_location_performance_report"
        scope = _reporting_service.factory.create('AccountThroughAdGroupReportScope')
        scope.AccountIds = {'long': [account_id]}
        scope.Campaigns = None
        scope.AdGroups = None
        report_request.Scope = scope

        report_columns = _reporting_service.factory.create('ArrayOfUserLocationPerformanceReportColumn')
        report_columns.UserLocationPerformanceReportColumn.append([
            "AccountName",
            "AccountNumber",
            "AccountId",
            "TimePeriod",
            "CampaignName",
            "CampaignId",
            "AdGroupName",
            "AdGroupId",
            "Country",
            "State",
            "MetroArea",
            "AdDistribution",
            "Impressions",
            "Clicks",
            "AverageCpc",
            "Spend",
            "AveragePosition",
            "ProximityTargetLocation",
            "Radius",
            "Language",
            "City",
            "QueryIntentCountry",
            "QueryIntentState",
            "QueryIntentCity",
            "QueryIntentDMA",
            "BidMatchType",
            "DeliveredMatchType",
            "Network",
            "TopVsOther",
            "DeviceType",
            "DeviceOS",
            "Assists",
            "Conversions",
            "Revenue",
            "ReturnOnAdSpend",
            "CostPerConversion",
            "CostPerAssist",
            "RevenuePerConversion",
            "RevenuePerAssist",
            "County",
            "PostalCode",
            "QueryIntentCounty",
            "QueryIntentPostalCode",
            "LocationId",
            "QueryIntentLocationId",
            "AllConversions",
            "AllRevenue",
            "AllCostPerConversion",
            "AllReturnOnAdSpend",
            "AllRevenuePerConversion"
        ])
        report_request.Columns = report_columns
        return report_request

    @staticmethod
    def get_goals_funnels_report_request(
            _reporting_service,
            account_id,
            aggregation,
            exclude_column_headers,
            exclude_report_footer,
            exclude_report_header,
            report_file_format,
            return_only_complete_data,
            report_time):
        """
        Reference:  https://docs.microsoft.com/en-us/advertising/reporting-service/goalsandfunnelsreportrequest?view=bingads-13
                    https://docs.microsoft.com/en-us/advertising/reporting-service/goalsandfunnelsreportcolumn?view=bingads-13
        """
        report_request = _reporting_service.factory.create('GoalsAndFunnelsReportRequest')
        report_request.Aggregation = aggregation
        report_request.ExcludeColumnHeaders = exclude_column_headers
        report_request.ExcludeReportFooter = exclude_report_footer
        report_request.ExcludeReportHeader = exclude_report_header
        report_request.Format = report_file_format
        report_request.ReturnOnlyCompleteData = return_only_complete_data
        report_request.ReportName = "goals_funnels_report"
        report_request.Time = report_time
        scope = _reporting_service.factory.create('AccountThroughAdGroupReportScope')
        scope.AccountIds = {'long': [account_id]}
        scope.Campaigns = None
        report_request.Scope = scope
        report_columns = _reporting_service.factory.create('ArrayOfGoalsAndFunnelsReportColumn')
        report_columns.GoalsAndFunnelsReportColumn.append([
            "AccountName",
            "AccountNumber",
            "AccountId",
            "TimePeriod",
            "CampaignName",
            "CampaignId",
            "AdGroupName",
            "AdGroupId",
            "Keyword",
            "KeywordId",
            "Goal",
            "AllConversions",
            "Assists",
            "AllRevenue",
            "GoalId",
            "DeviceType",
            "DeviceOS",
            "AccountStatus",
            "CampaignStatus",
            "AdGroupStatus",
            "KeywordStatus",
            "GoalType"
        ])
        report_request.Columns = report_columns
        return report_request

    @staticmethod
    def get_ad_performance_report_request(
            _reporting_service,
            account_id,
            aggregation,
            exclude_column_headers,
            exclude_report_footer,
            exclude_report_header,
            report_file_format,
            return_only_complete_data,
            report_time):
        """
        Reference:  https://docs.microsoft.com/en-us/advertising/reporting-service/adperformancereportrequest?view=bingads-13
                    https://docs.microsoft.com/en-us/advertising/reporting-service/adperformancereportcolumn?view=bingads-13
        """
        report_request = _reporting_service.factory.create('AdPerformanceReportRequest')
        report_request.Aggregation = aggregation
        report_request.ExcludeColumnHeaders = exclude_column_headers
        report_request.ExcludeReportFooter = exclude_report_footer
        report_request.ExcludeReportHeader = exclude_report_header
        report_request.Format = report_file_format
        report_request.ReturnOnlyCompleteData = return_only_complete_data
        report_request.ReportName = "ads_performance_report"
        report_request.Time = report_time
        scope = _reporting_service.factory.create('AccountThroughAdGroupReportScope')
        scope.AccountIds = {'long': [account_id]}
        scope.Campaigns = None
        report_request.Scope = scope
        report_columns = _reporting_service.factory.create('ArrayOfAdPerformanceReportColumn')
        report_columns.AdPerformanceReportColumn.append([
            "AccountName",
            "AccountNumber",
            "AccountId",
            "TimePeriod",
            "CampaignName",
            "CampaignId",
            "AdGroupName",
            "AdId",
            "AdGroupId",
            "AdTitle",
            "AdDescription",
            "AdDescription2",
            "AdType",
            "AdDistribution",
            "Impressions",
            "Clicks",
            "AverageCpc",
            "Spend",
            "AveragePosition",
            "Conversions",
            "CostPerConversion",
            "DestinationUrl",
            "DeviceType",
            "Language",
            "DisplayUrl",
            "AdStatus",
            "Network",
            "TopVsOther",
            "BidMatchType",
            "DeliveredMatchType",
            "DeviceOS",
            "Assists",
            "Revenue",
            "ReturnOnAdSpend",
            "CostPerAssist",
            "RevenuePerConversion",
            "RevenuePerAssist",
            "TrackingTemplate",
            "CustomParameters",
            "FinalUrl",
            "FinalMobileUrl",
            "FinalAppUrl",
            "AccountStatus",
            "CampaignStatus",
            "AdGroupStatus",
            "TitlePart1",
            "TitlePart2",
            "TitlePart3",
            "Headline",
            "LongHeadline",
            "BusinessName",
            "Path1",
            "Path2",
            "AdLabels",
            "CustomerId",
            "CustomerName",
            "CampaignType",
            "BaseCampaignId",
            "AllConversions",
            "AllRevenue",
            "AllCostPerConversion",
            "AllReturnOnAdSpend",
            "AllRevenuePerConversion",
            "FinalUrlSuffix"
        ])
        report_request.Columns = report_columns
        return report_request

    @staticmethod
    def get_ads_dictionary_report_request(
            _reporting_service,
            account_id,
            exclude_column_headers,
            exclude_report_footer,
            exclude_report_header,
            report_file_format,
            return_only_complete_data):
        """
        Reference:  https://docs.microsoft.com/en-us/advertising/reporting-service/adperformancereportrequest?view=bingads-13
                    https://docs.microsoft.com/en-us/advertising/reporting-service/adperformancereportcolumn?view=bingads-13
        """
        dates = MicrosoftAdsAPI.get_custom_dates(1000, 0)
        date_from = dates[0]
        date_to = dates[1]

        report_request = _reporting_service.factory.create('AdPerformanceReportRequest')
        report_request.Aggregation = 'Summary'

        report_request.ExcludeColumnHeaders = exclude_column_headers
        report_request.ExcludeReportFooter = exclude_report_footer
        report_request.ExcludeReportHeader = exclude_report_header

        report_request.Format = report_file_format
        report_request.ReturnOnlyCompleteData = return_only_complete_data
        report_request.ReportName = "ads_dictionary_report"

        report_time = _reporting_service.factory.create('ReportTime')
        custom_date_range_start = _reporting_service.factory.create('Date')
        custom_date_range_start.Day = date_from.day
        custom_date_range_start.Month = date_from.month
        custom_date_range_start.Year = date_from.year
        custom_date_range_end = _reporting_service.factory.create('Date')
        custom_date_range_end.Day = date_to.day
        custom_date_range_end.Month = date_to.month
        custom_date_range_end.Year = date_to.year
        report_time.CustomDateRangeStart = custom_date_range_start
        report_time.CustomDateRangeEnd = custom_date_range_end
        report_time.ReportTimeZone = 'PacificTimeUSCanadaTijuana'
        report_request.Time = report_time

        scope = _reporting_service.factory.create('AccountThroughAdGroupReportScope')
        scope.AccountIds = {'long': [account_id]}
        scope.Campaigns = None
        report_request.Scope = scope
        report_columns = _reporting_service.factory.create('ArrayOfAdPerformanceReportColumn')
        report_columns.AdPerformanceReportColumn.append([
            "AccountName",
            "AccountNumber",
            "AccountId",
            "CampaignName",
            "CampaignId",
            "AdGroupName",
            "AdId",
            "AdGroupId",
            "AdTitle",
            "AdDescription",
            "AdDescription2",
            "AdType",
            "AdDistribution",
            "Impressions",
            "DestinationUrl",
            "DisplayUrl",
            "AdStatus",
            "TrackingTemplate",
            "CustomParameters",
            "FinalUrl",
            "FinalMobileUrl",
            "FinalAppUrl",
            "AccountStatus",
            "CampaignStatus",
            "AdGroupStatus",
            "TitlePart1",
            "TitlePart2",
            "TitlePart3",
            "Headline",
            "LongHeadline",
            "BusinessName",
            "Path1",
            "Path2",
            "CustomerId",
            "CustomerName",
            "CampaignType",
            "BaseCampaignId",
            "FinalUrlSuffix"
        ])
        report_request.Columns = report_columns
        return report_request

    def get_report_request(self, account_id, _reporting_service, date_from, date_to):
        """
        Use a sample report request or build your own.
        """
        aggregation = 'Daily'
        exclude_column_headers = False
        exclude_report_footer = True
        exclude_report_header = True
        report_time = _reporting_service.factory.create('ReportTime')
        # You can either use a custom date range or predefined time.
        # report_time.PredefinedTime = 'Last30Days'
        report_time.PredefinedTime = None
        custom_date_range_start = _reporting_service.factory.create('Date')
        custom_date_range_start.Day = date_from.day
        custom_date_range_start.Month = date_from.month
        custom_date_range_start.Year = date_from.year
        custom_date_range_end = _reporting_service.factory.create('Date')
        custom_date_range_end.Day = date_to.day
        custom_date_range_end.Month = date_to.month
        custom_date_range_end.Year = date_to.year
        report_time.CustomDateRangeStart = custom_date_range_start
        report_time.CustomDateRangeEnd = custom_date_range_end
        report_time.ReportTimeZone = 'PacificTimeUSCanadaTijuana'
        return_only_complete_data = False

        """
        # BudgetSummaryReportRequest does not contain a definition for Aggregation.
        budget_summary_report_request = self.get_budget_summary_report_request(
            _reporting_service=_reporting_service,
            account_id=account_id,
            exclude_column_headers=exclude_column_headers,
            exclude_report_footer=exclude_report_footer,
            exclude_report_header=exclude_report_header,
            report_file_format=self.REPORT_FILE_FORMAT,
            return_only_complete_data=return_only_complete_data,
            report_time=report_time)

        campaign_performance_report_request = self.get_campaign_performance_report_request(
            _reporting_service=_reporting_service,
            account_id=account_id,
            aggregation=aggregation,
            exclude_column_headers=exclude_column_headers,
            exclude_report_footer=exclude_report_footer,
            exclude_report_header=exclude_report_header,
            report_file_format=self.REPORT_FILE_FORMAT,
            return_only_complete_data=return_only_complete_data,
            report_time=report_time)
        """

        search_query_performance_report_request = self.get_search_query_performance_report_request(
            _reporting_service=_reporting_service,
            account_id=account_id,
            aggregation=aggregation,
            exclude_column_headers=exclude_column_headers,
            exclude_report_footer=exclude_report_footer,
            exclude_report_header=exclude_report_header,
            report_file_format=self.REPORT_FILE_FORMAT,
            return_only_complete_data=return_only_complete_data,
            report_time=report_time)

        keyword_performance_report_request = self.get_keyword_performance_report_request(
            _reporting_service=_reporting_service,
            account_id=account_id,
            aggregation=aggregation,
            exclude_column_headers=exclude_column_headers,
            exclude_report_footer=exclude_report_footer,
            exclude_report_header=exclude_report_header,
            report_file_format=self.REPORT_FILE_FORMAT,
            return_only_complete_data=return_only_complete_data,
            report_time=report_time)

        user_location_performance_report_request = self.get_user_location_performance_report_request(
            _reporting_service=_reporting_service,
            account_id=account_id,
            aggregation=aggregation,
            exclude_column_headers=exclude_column_headers,
            exclude_report_footer=exclude_report_footer,
            exclude_report_header=exclude_report_header,
            report_file_format=self.REPORT_FILE_FORMAT,
            return_only_complete_data=return_only_complete_data,
            report_time=report_time)

        goals_funnels_report_request = self.get_goals_funnels_report_request(
            _reporting_service=_reporting_service,
            account_id=account_id,
            aggregation=aggregation,
            exclude_column_headers=exclude_column_headers,
            exclude_report_footer=exclude_report_footer,
            exclude_report_header=exclude_report_header,
            report_file_format=self.REPORT_FILE_FORMAT,
            return_only_complete_data=return_only_complete_data,
            report_time=report_time)

        ad_performance_report_request = self.get_ad_performance_report_request(
            _reporting_service=_reporting_service,
            account_id=account_id,
            aggregation=aggregation,
            exclude_column_headers=exclude_column_headers,
            exclude_report_footer=exclude_report_footer,
            exclude_report_header=exclude_report_header,
            report_file_format=self.REPORT_FILE_FORMAT,
            return_only_complete_data=return_only_complete_data,
            report_time=report_time)

        ads_dictionary_report_request = self.get_ads_dictionary_report_request(
            _reporting_service=_reporting_service,
            account_id=account_id,
            exclude_column_headers=exclude_column_headers,
            exclude_report_footer=exclude_report_footer,
            exclude_report_header=exclude_report_header,
            report_file_format=self.REPORT_FILE_FORMAT,
            return_only_complete_data=return_only_complete_data)

        return (
            ads_dictionary_report_request,
            ad_performance_report_request,
            keyword_performance_report_request,
            search_query_performance_report_request,
            goals_funnels_report_request,
            user_location_performance_report_request,
        )
        # return (test_dictionary_request)

    def submit_and_download(self, report_request, _result_file_name, _reporting_service_manager):
        """ Submit the download request and then use the ReportingDownloadOperation result to
        track status until the report is complete e.g. either using
        ReportingDownloadOperation.track() or ReportingDownloadOperation.get_status(). """

        # global reporting_service_manager
        # reporting_download_operation = _reporting_service_manager.submit_download(report_request)

        # You may optionally cancel the track() operation after a specified time interval.
        # reporting_operation_status = reporting_download_operation.track(
        #                                                 timeout_in_milliseconds=self.TIMEOUT_IN_MILLISECONDS)

        # You can use ReportingDownloadOperation.track() to poll until complete as shown above,
        # or use custom polling logic with get_status() as shown below.
        for _ in range(10):
            # time.sleep(_reporting_service_manager.poll_interval_in_milliseconds / 1000.0)
            reporting_download_operation = _reporting_service_manager.submit_download(report_request)
            time.sleep(20)
            download_status = reporting_download_operation.get_status()
            print(download_status.status)
            print(download_status.report_download_url)
            if download_status.report_download_url is not None:
                break

        result_file_path = reporting_download_operation.download_result_file(
            result_file_directory=self.FILE_DIRECTORY,
            result_file_name=_result_file_name,
            decompress=True,
            overwrite=True,  # Set this value true if you want to overwrite the same file.
            timeout_in_milliseconds=self.TIMEOUT_IN_MILLISECONDS  # You may optionally cancel the download after a
            # specified time interval.
        )

        self.output_status_message("Download result file: {0}".format(result_file_path))

    def get_requested_reports_submit_download(self, account_id, _reporting_service, _reporting_service_manager,
                                              date_from, date_to):
        try:
            report_request = self.get_report_request(account_id, _reporting_service, date_from, date_to)
            for report in report_request:
                _result_file_name = '{0}.'.format(report.ReportName) + self.REPORT_FILE_FORMAT.lower()
                print(_result_file_name)
                # Option B - Submit and Download with ReportingServiceManager
                # ------------------------------------------------------------
                # Submit the download request and then use the ReportingDownloadOperation result to
                # track status yourself using ReportingServiceManager.get_status().
                self.output_status_message("-----\nAwaiting Submit and Download...")
                self.submit_and_download(report, _result_file_name, _reporting_service_manager)
        except WebFault as ex:
            self.output_webfault_errors(ex)
        except Exception as ex:
            self.output_status_message(ex)

    def download_report(self, _reporting_download_parameters, _reporting_service_manager):
        """
        You can get a Report object by submitting a new download request via ReportingServiceManager.
        Although in this case you will not work directly with the file, under the covers a request is
        submitted to the Reporting service and the report file is downloaded to a local directory.
        """
        report_container = _reporting_service_manager.download_report(_reporting_download_parameters)
        if report_container is None:
            self.output_status_message("There is no report data for the submitted report request parameters.")
            # sys.exit(0)
            report_container.close()

    def get_requested_reports_download_report(self, account_id, _reporting_service, _reporting_service_manager,
                                              date_from, date_to):
        try:
            report_request = self.get_report_request(account_id, _reporting_service, date_from, date_to)
            for report in report_request:
                _result_file_name = '{0}_{1}_input.'.format(account_id,
                                                            report.ReportName) + self.REPORT_FILE_FORMAT.lower()
                print(_result_file_name)
                reporting_download_parameters = ReportingDownloadParameters(
                    report_request=report,
                    result_file_directory=self.FILE_DIRECTORY,
                    result_file_name=_result_file_name,
                    overwrite_result_file=True,  # Set this value true if you want to overwrite the same file.
                    timeout_in_milliseconds=self.TIMEOUT_IN_MILLISECONDS
                    # You may optionally cancel the download after a specified time interval.
                )

                # Option D - Download the report in memory with ReportingServiceManager.download_report
                # The download_report helper function downloads the report and summarizes results.
                self.output_status_message("-----\nAwaiting download_report...")
                self.download_report(reporting_download_parameters, _reporting_service_manager)

        except WebFault as ex:
            self.output_webfault_errors(ex)
        except Exception as ex:
            self.output_status_message(ex)

    @staticmethod
    def get_custom_dates(lb_window=29, days_skip=0):
        today = dt.datetime.utcnow()
        to_date = today - dt.timedelta(1 + days_skip)
        from_date = to_date - dt.timedelta(lb_window)
        return from_date, to_date
