import sys
import logging
import re
import requests
import time 
import random
import csv
import pytz
from openpyxl import Workbook
from datetime import datetime, timedelta
from lxml import html
from requests.packages.urllib3.exceptions import InsecureRequestWarning

requests.packages.urllib3.disable_warnings(InsecureRequestWarning);

#
# Get today (NZ time) as YYYY-mm-dd. This will be used in the
# output file name to indicate the day this script is run.
# Not used anywhere else.
#
def get_current_date_nztz():
    now = datetime.now(pytz.timezone('Pacific/Auckland'));
    print("NZ Time now: " + str(now));
    print("NZ Time now in a readable format: " + now.strftime("%Y-%m-%d %I:%M %p %Z"));
    fileNameTimeStamp = now.strftime("%Y_%m_%d_Time_%I_%p_%Z")
    return fileNameTimeStamp;

#
# URLs that we need to invoke at various steps
#
miqFlightCheckerURL = "https://allocation.miq.govt.nz/portal/flight-checker";

# Not used after the MIQ site update on 2021-10-01
miqPortalURL = "https://allocation.miq.govt.nz/portal/";
miqLobbyURL = "https://lobby.miq.govt.nz";

outputFileBaseName = "NZMIQ_Flights_Data_as_of_" + get_current_date_nztz();
csvFile = outputFileBaseName + '.csv';
excelFile = outputFileBaseName + '.xlsx';
csvLogger = None;
infoLogger = None;

#
# HTTP session
#
session = requests.Session();
session.verify = False;
session.strict_mode = True;

def setupLogger(name, logFile, format=None, consoleOutput=sys.stdout):
    fileHandler = logging.FileHandler(logFile, mode='w');
    consoleLogHandler = logging.StreamHandler(consoleOutput);
    if (format is not None):
        fileHandler.setFormatter(logging.Formatter(format));
        consoleLogHandler.setFormatter(logging.Formatter(format))
    logger = logging.getLogger(name);
    logger.setLevel(logging.INFO);

    # The logger will have a file to write the outputs
    logger.addHandler(fileHandler);

    # The logger also writes the output to console (stdout or stderr)
    logger.addHandler(consoleLogHandler);
    return logger

def configureLogging():
    global csvLogger, infoLogger;
    #
    # Create a CSV file that contains the details of
    # the flight numbers, origin, destination, etc.
    #
    csvLogger = setupLogger("csvLogger", csvFile);
    #
    # Create a log file to write any errors such as
    # Twitter API errors and date parsing errors.
    #
    infoLogger = setupLogger("infoLogger", "info.log");

def print_and_exit(errMsg):
    infoLogger.warning("Error: " + errMsg + ". Exiting...");
    sys.exit(1);

def print_call_start(i, url):
    infoLogger.info("Step " + str(i) + ": Start");
    infoLogger.info("URL to call: " + url);

def print_call_end(i):
    infoLogger.info("Step " + str(i) + ": End");
    infoLogger.info("");
    infoLogger.info("");

def get_standard_headers():
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:81.0) Gecko/20100101 Firefox/81.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',  
        'Accept-Encoding' : 'gzip, deflate, br',
        'Accept-Language' : 'en-US,en;q=0.5',
        'Connection' : 'keep-alive',
        'DNT' : '1',
        'Pragma' : 'no-cache',
        'Cache-Control' : 'no-cache',
        'Upgrade-Insecure-Requests' : '1',
    }
    return headers;

#
# Parse the HTML data containing the flight data for a given date
# Also, print the flight details thus found.
#
def parse_print_flight_data_html(dateOfInterest, flightDataHtml):
    tree = html.fromstring(flightDataHtml);
    flightDivXPath = '/html/body//div[@class="accordion__item"]';
    flightDivs = tree.xpath(flightDivXPath);
    dateOfInterestDateObj = datetime.strptime(dateOfInterest, '%Y-%m-%d');
    dayOfWeek = dateOfInterestDateObj.strftime("%A");
    for aFlightDiv in flightDivs:
        flightCarrierName = aFlightDiv.xpath(
            'div[contains(@class, "pt-4 pb-2 pb-sm-2")]/h3/button/text()')[0].strip();
        flightNumberRowXPath = 'div/div[@class="pb-10"]/table/tbody/tr[contains(@class, "d-block d-sm-table-row")]';
        flightNumberRows = aFlightDiv.xpath(flightNumberRowXPath);
        for aFlightNumRow in flightNumberRows:
            cells = aFlightNumRow.xpath('td/text()');
            cellCount = 0;
            nonEmptyCells = [];
            for aCell in cells:
                aCell = aCell.strip();
                if not aCell:
                    continue;
                nonEmptyCells.append(aCell);
            flightNum = nonEmptyCells[0].strip();
            origin = nonEmptyCells[1].strip();
            arrivalPort = nonEmptyCells[2].strip();
            estArrivalTime = nonEmptyCells[3].strip();
            csvLogger.info(dateOfInterest + "," 
                + dayOfWeek + ","
                + flightCarrierName + "," 
                + flightNum + "," 
                + origin + "," 
                + arrivalPort + "," 
                + estArrivalTime);

#
# Function to test the flight data extraction logic using
# a sample HTML file got from the MIQ site.
#
def parse_sample_html():
    with open('SampleFlightData.html', 'r') as f:
        flightDataHtml = f.read();
    parse_print_flight_data_html("2021-11-01", flightDataHtml);

#
# Convert the CSV file into an Excel sheet.
#
def convert_to_excel():
    wb = Workbook();
    ws = wb.active;
    with open(csvFile, 'r') as f:
        for row in csv.reader(f):
            ws.append(row);
    wb.save(excelFile);

def get_flight_date_data():
    #
    # Print the CSV heading row
    #
    csvLogger.info("Date of Arrival," 
        + "Day of the Week," 
        + "Airlines,"
        + "Flight Number,"
        + "Origin,"
        + "Arrival Port,"
        + "Est. Arrival Time (NZ)");

    #
    # There are a few HTTPS calls that we need to mae in order to get a
    # flight checker token and using that, to get the flight dates details.
    #
    
    #
    # Call 1: Visit the https://allocation.miq.govt.nz/portal/flight-checker
    # This call will return, in its HTML file, the flight checker token 
    # and the start/end dates of the flight data.
    #
    print_call_start(1, miqFlightCheckerURL);
    response = session.post(miqFlightCheckerURL, headers = get_standard_headers());
    htmlContent = response.text;
    if "flight_checker__token" not in htmlContent:
        print_and_exit("The HTML page does not contain the flight_checker__token snippet");
    tree = html.fromstring(htmlContent);
    flightCheckerTokenXPath = '/html/body/' + \
        '/form[@name="flight_checker"]/input[@id="flight_checker__token"]/@value';
    flightCheckerToken = tree.xpath(flightCheckerTokenXPath)[0];
    infoLogger.info("Got the flightCheckerToken as [" + flightCheckerToken + "]");

    chosenDateXPath = '/html/body//input[@id="flight_checker_chosenDate"]/'
    minDateStr = tree.xpath(chosenDateXPath + '@min')[0];
    maxDateStr = tree.xpath(chosenDateXPath + '@max')[0];
    infoLogger.info("Got these details: minDate [" 
        + minDateStr + "] and maxDate [" + maxDateStr + "]");
    print_call_end(1);

    #
    # Call 2:
    # Now that we have gotten flight Checker Token, min/max dates,
    # let us iterate through the dates and fetch the flight details
    # for each date. 
    #

    # Convert date strings of format YYYY-mm-DD (e.g.: 2021-12-31) into date objects 
    minAvailableDate = datetime.strptime(minDateStr, '%Y-%m-%d');
    maxAvailableDate = datetime.strptime(maxDateStr, '%Y-%m-%d');
    delta = timedelta(days=1)
    data = {
        'flight_checker[_token]': flightCheckerToken
    }
    cookies = response.cookies;
    count = 0;
    print_call_start(2, miqFlightCheckerURL);
    while minAvailableDate <= maxAvailableDate:
        dateOfInterest = minAvailableDate.strftime("%Y-%m-%d");
        infoLogger.info("");
        infoLogger.info("Fetching the flight details for date [" + dateOfInterest + "]");
        data['flight_checker[chosenDate]'] = dateOfInterest;
        response = session.post(miqFlightCheckerURL, headers = get_standard_headers(),
            data = data, cookies=cookies);
        parse_print_flight_data_html(dateOfInterest, response.text);
        minAvailableDate += delta
        #
        # Sleep for few seconds so that the server does not feel a DoS attack.
        #
        sleepTime = random.randint(1,4);
        infoLogger.info("Sleeping for " + str(sleepTime) + "s");
        time.sleep(sleepTime);
        count = count + 1;
        # Uncomment the below to do a test run of just a few dates
        #if (count > 2):
        #    break;

    print_call_end(2);

    #
    # Convert the generated CSV file into an Excel file
    #
    infoLogger.info("");
    infoLogger.info("Converting the CSV file into an Excel file");
    convert_to_excel();
    infoLogger.info("All done - See the output in the " + excelFile + " file");

def main():
    configureLogging();
    get_flight_date_data();
    #parse_sample_html();

if __name__== "__main__":
    main();
