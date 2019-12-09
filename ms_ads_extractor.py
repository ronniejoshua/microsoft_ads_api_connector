import ms_ads
import csv
import datetime as dt
import os
import ms_ads_uploads
import argparse

if __name__ == '__main__':

    # Passing Arguments through Command Line
    parser = argparse.ArgumentParser(
        description="Sending Microsoft Ads Data To Big Query"
    )

    parser.add_argument(
        "-d",
        "--days_back",
        type=int,
        metavar="",
        required=True,
        help="Look Back Window Start Date"
    )

    parser.add_argument(
        "-s",
        "--days_skip",
        type=int,
        metavar="",
        required=False,
        default=0,
        help="Look Back Window End Date, if 0 then end date = yesterday"
    )

    args = parser.parse_args()

    # Account Credentials
    CLIENT_ID = 'my_client_id'
    DEVELOPER_TOKEN = 'developer_id'
    ENVIRONMENT = 'production'
    
    if os.name == 'posix':
        REFRESH_TOKEN = r'./refresh_token/refresh.txt'
    else:
        REFRESH_TOKEN = r''
        
    CLIENT_STATE = 'my_client_state'
    
    WRITE_TO_BQ = True

    # Initializing an Extractor Instance
    extractor = ms_ads.MicrosoftAdsAPI(CLIENT_ID, DEVELOPER_TOKEN, ENVIRONMENT, REFRESH_TOKEN, CLIENT_STATE)

    # Input Dates
    custom_dates = extractor.get_custom_dates(args.days_back, args.days_skip)
    print(custom_dates)
    date_from = custom_dates[0]
    date_to = custom_dates[1]

    print("Loading the web service client proxies...")
    authorization_data = ms_ads.AuthorizationData(
        account_id=None,
        customer_id=None,
        developer_token=extractor.DEVELOPER_TOKEN,
        authentication=None,
    )

    customer_service = ms_ads.ServiceClient(
        service='CustomerManagementService',
        version=13,
        authorization_data=authorization_data,
        environment=extractor.ENVIRONMENT,
    )

    reporting_service_manager = ms_ads.ReportingServiceManager(
        authorization_data=authorization_data,
        poll_interval_in_milliseconds=5000,
        environment=extractor.ENVIRONMENT,
    )

    reporting_service = ms_ads.ServiceClient(
        service='ReportingService',
        version=13,
        authorization_data=authorization_data,
        environment=extractor.ENVIRONMENT,
    )

    account_ids = extractor.authenticate(authorization_data)
    print(account_ids)
    for account in account_ids:
        print(account)
        report_request = extractor.get_requested_reports_download_report(account,
                                                                         reporting_service,
                                                                         reporting_service_manager,
                                                                         date_from,
                                                                         date_to
                                                                         )

    _insert_time = dt.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    directory = extractor.FILE_DIRECTORY
    for file_name in os.listdir(directory):
        if file_name.endswith(".csv"):
            output_data = list()
            headers = list()
            with open(r'{0}/{1}'.format(directory, file_name), 'r', encoding='utf-8-sig')as read_file:
                data = csv.reader(read_file, delimiter=',')
                for enum, row in enumerate(data):
                    if enum == 0:
                        row.insert(0, '_insert_time')
                        headers.extend(row)
                        output_data.append(row)
                    if enum > 0:
                        row.insert(0, _insert_time)
                        output_data.append(row)

            with open(r'{0}/{1}'.format(directory, file_name.replace("input", "output")), 'w', newline='',
                      encoding='utf-8-sig') as resultFile:
                wr = csv.writer(resultFile)
                for item in output_data:
                    wr.writerow(item)
    if WRITE_TO_BQ:
        data_uploader = ms_ads_uploads.DataUploader()
        data_uploader.execute_uploader(directory)

    for file_name in os.listdir(directory):
         os.remove(r'{0}/{1}'.format(directory, file_name))

