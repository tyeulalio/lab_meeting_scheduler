# this script will automatically generate a schedule for the upcoming 
# school quarter given the appropriate input dates
from datetime import datetime, timedelta
import pandas as pd
import pytz
import dateutil.tz
import time
# connecting to google drive
import gspread
from oauth2client.service_account import ServiceAccountCredentials



class Spreadsheet():
    def __init__(self):
        # use creds to create a client to interact with the Google Drive API
        scope = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
                ]
        creds = ServiceAccountCredentials.from_json_keyfile_name('.ignore/client_secret.json', scope)
        client = gspread.authorize(creds)

        # Find a workbook by name and open the first sheet
        self.sheet = client.open("Montgomery Lab Meeting Data")

        # extract sheets
        self.response_sheet = self.sheet.get_worksheet(0)
        self.responses_sorted_sheet = self.sheet.get_worksheet(1)
        self.lab_presenters_sheet = self.sheet.get_worksheet(2)
        self.rotation_students_sheet = self.sheet.get_worksheet(3)
        self.breaks_sheet = self.sheet.get_worksheet(4)

        # example of pulling data from a sheet
        # test = self.sheet.get_worksheet(1)
        # list_of_hashes = test.get_all_records()
        # print(list_of_hashes)


    

class Calendar():
    def __init__(self, input_file):
        self.input_file = input_file
        self.data_dict = {}
        # create a diction to store days of the week
        self.days_dict = {
                'Monday': 0,
                'Tuesday': 1,
                'Wednesday': 2,
                'Thursday': 3,
                'Friday': 4,
                'Saturday': 5,
                'Sunday': 6}
        self.schedule = {}

        # read input file
        self.read_input()

        # decompose the meeting time a bit
        meet_hr,rest = self.data_dict['meeting_time'].split(':')
        self.meet_hr = int(meet_hr)
        self.meet_min = int(rest[:2])
        meet_m = rest[2:]

        if meet_m == 'PM':
            self.meet_hr += 12
        
        
        # process holidays
        self.data_dict['holiday_dates'] = self.get_holiday_dates()


    def get_holiday_dates(self):
        print("Processing holidays")
        # read in the holidays from the data dict
        # convert the dates to datetime objects
        # store back into data dict
        stanford_holidays = self.data_dict['stanford_holidays']
        date_format = "%m/%d/%Y"
        
        dates_dict = {}

        for holiday in stanford_holidays:
            name,date = holiday.split(';')

            # remove quotes from name
            name = name.strip('"')

            # some dates are ranges, so handle these
            start_date = 0
            end_date = 0
            drange = False
            d = []

            # range of dates
            if '-' in date:
                start_date, end_date = date.split('-')
                drange = True

                # convert to datetime objects
                startd = datetime.strptime(start_date, date_format)
                endd = datetime.strptime(end_date, date_format)

                d = pd.date_range(startd, endd, freq='d')

            else:
                # just a single date
                d.append(datetime.strptime(date, date_format))
          

            d_tz = []
            # add appropriate timezone utc offset for the holidays
            for dt in d:
                offset = self.get_offset(dt)
                new_dt = dt.replace(hour=self.meet_hr, minute=self.meet_min,
                        tzinfo=dateutil.tz.tzoffset(None, offset*60*60))
                d_tz.append(new_dt)
        
        
            # add each holiday date to the dictionary
            for holidate in d_tz:
                # *** TO DO: If two holidays fall on same day, last entry will overwrite previous **
                dates_dict[holidate] = "No lab meeting: {}".format(name) 

        return dates_dict


    def read_input(self):
        # read the input file and populate data fields
        with open(self.input_file, 'r') as f:
            label = '' # label for current data field
            data = '' # keeps track of data for current field

            within_list = False
            list_data = []

            for line in f:
                line = line.rstrip()

                # check if line is an assignment line (contains =)
                if '=' in line:
                    label,data = line.split('=', 1)
                else:
                    data = line

                # check if we have a list
                if data == '[':
                    # we need to keep reading until we find end bracket
                    within_list = True
                    list_data = []
                    continue

                # check for end of list
                if data == ']':
                    self.data_dict[label] = list_data
                    within_list = False
                    continue

                # if we're in list, then keep adding
                if within_list == True:
                    list_data.append(data)
                else:
                    # we're not in a list here, so just add data to dict
                    self.data_dict[label] = data

    def print_datadict(self):
        # used for testing - print the data dict
        for key in self.data_dict:
            print("{}: {}".format(key, self.data_dict[key]))

    
    def get_meeting_weekdays(self, start_date, end_date, meeting_day):
        # grab all meeting weekdays between the start and end date
        # For example, if meeting day is Monday, grab all Mondays that occur
        # between the start and end dates
        # start_date = datetime object
        # end_date = datetime object
        # meeting_day = integer for date
        d = start_date
        d += timedelta(days=meeting_day - d.weekday()) # get the first meeting_day

        meeting_dates = []

        while d < end_date:
            meeting_dates.append(d)
            d += timedelta(days=7)

        # add two more days for special meetings
        # for i in range(2):
            # meeting_dates.append(d)
            # d += timedelta(days=7)

        return meeting_dates

   
    def get_offset(self, dt):
        # -- we don't really need to do this because
        # the ics file takes timezone as a string anyways
        #
        # get the appropriate UTC offset for the current time
        # this dependnds on whether we are in daylight savings time or not
            
        # returns True if date is in daylight savings in LA timezone
        # returns False otherwise
        tz = pytz.timezone('America/Los_Angeles')
        tz_aware = tz.localize(dt, is_dst=None)
        isdst = tz_aware.tzinfo._dst.seconds != 0
       
        offset = -8
        # set the UTC offset
        if isdst:
            offset = -7
        
        return offset

        
    def add_special_meetings(self, schedule):
        # add the special meeting days that occur every quarter
        # these include:
        # - advocacy: last meeting of quarter
        # - housekeeping: first meeting of quarter
        # - rotation students: second to last meeting of quarter
        
        # get all of the dates
        all_dates = schedule.keys()
        # remove dates that already have holidays
        avail_dates = [x for x in all_dates if schedule[x] == ""]

        # get important dates
        first_date = avail_dates[0]
        last_date = avail_dates[-1]

        # set advocacy date to last day
        advocacy_date = last_date

        # assign rotation meeting if there are rotation students
        if (len(self.data_dict['rotation_students']) > 0) & (self.data_dict['rotation'] == "True"):
            # if no advocacy meeting, schedule on last day
            # else, schedule second to last day
            rotation_date = last_date

            schedule[rotation_date] = "Rotation Students Lab Meeting"

            # if there's a rotation meeting, set advocacy to second to last date
            advocacy_date = avail_dates[-2]


        # special_days = self.special_days
        # assign dates if they're set to true
        if self.data_dict['housekeeping'] == "True":
            schedule[first_date] = "Housekeeping (Stephen) Lab Meeting"
        if self.data_dict['advocacy'] == "True":
            schedule[advocacy_date] = "Advocacy Lab Meeting"




        return schedule

    
    def assign_presenters(self, schedule):
        # assign presenters to the open dates

        # get available dates
        all_dates = schedule.keys()
        avail_dates = [x for x in all_dates if schedule[x] == ""]

        # get last presenter
        last_presenter = self.data_dict['last_presenter']

        # get all presenters
        presenters = self.data_dict['lab_presenters']

        # find the next presenter
        last_presenter_idx = presenters.index(last_presenter)

        num_presenters = len(presenters)


        for i in range(len(avail_dates)):
            # get next presenter
            next_presenter_idx = (last_presenter_idx + i+1) % num_presenters
            next_presenter = presenters[next_presenter_idx]

            # get next available date
            next_date = avail_dates[i]

            # assign this person to this day
            schedule[next_date] = "{}: Lab Meeting".format(next_presenter)


        return schedule


    def create_schedule(self):
        # this is where most of the work takes place
        # create the schedule based on input in the datadict
        # Need to follow some predetermined rules:
        # - First meeting of quarter is housekeeping meeting
        # - First person to present is the person in the rotation after the 'last presenter' from last quarter
        # - Rotation students get scheduled on second to the last week of quarter
        # - No lab meeting on Stanford holidays
        # - Last meeting of quarter is advocacy meeting
        # Later:
        # - can incporporate requests for other recognized holidays, maybe birthdays too?
        
        data_dict = self.data_dict

        # Start by getting the start and end dates in datetime format
        date_format = "%m/%d/%Y"
        start_date = datetime.strptime(data_dict['start'], date_format)
        end_date = datetime.strptime(data_dict['quart_end'], date_format)

        # get the recurring meeting date and time
        # encode the day of the week as a digit
        
        # grab every meeting weekday between the start and end dates
        meeting_day = self.days_dict[data_dict['meeting_day']]
        meeting_days = self.get_meeting_weekdays(start_date, end_date, meeting_day)

       
        # add the meeting time with appropriate timezone utc offset 
        schedule = {}
        for dt in meeting_days:
            offset = self.get_offset(dt)
            new_dt = dt.replace(hour=self.meet_hr, minute=self.meet_min, 
                    tzinfo=dateutil.tz.tzoffset(None, offset*60*60))
            schedule[new_dt] = ""

        # populate any holidays if they exist
        for holiday in data_dict['holiday_dates']:
            print("{}: {}".format(holiday, data_dict['holiday_dates'][holiday]))
            if holiday in schedule.keys():
                schedule[holiday] = data_dict['holiday_dates'][holiday]


        # assign special meetings
        schedule = self.add_special_meetings(schedule)


        # assign students to meetings
        schedule = self.assign_presenters(schedule)

        # store schedule
        self.schedule = schedule


        print("---------------")
        for day in schedule:
            print("{}: {}".format(day, schedule[day]))


    def write_ics(self):
        # write the schedule to an ics file
        output_file = self.data_dict['output_file']

        today = datetime.date(datetime.now())

        
        if self.data_dict['cancel'] == "True":
            output_file = "{}_{}_cancel.ics".format(today, output_file) 
        else:
            output_file = "{}_{}.ics".format(today, output_file)


        # open the output file
        output = open(output_file, 'w+')

        # need to open vcalendar object
        output.write("BEGIN:VCALENDAR\n")

        # create VEVENT objects for each lab meeting date
        for date in self.schedule:
            output.write("\nBEGIN:VEVENT\n")

            # --- UID --- #
            # make an ID for the event in case we want to
            # delete it later
            uid = "{}_{}".format(date.strftime('%Y%m%d'), self.schedule[date])
            uid_outstr = "UID:{}\n".format(uid)
            # output to file
            output.write(uid_outstr)


            # --- SUMMARY --- #
            # get event summary
            summary = self.schedule[date]
            summary_outstr = "SUMMARY:{}\n".format(summary)

            # output to file
            output.write(summary_outstr)


            # --- DTSTART/DTEND --- #
            # get date start and end 
            str_format = '%Y%m%dT%H%M%S' 
            date_start_str = date.strftime(str_format)
            tz='America/Los_Angeles'

            # add an hour for the end time
            date_end_str = (date + timedelta(hours=1)).strftime(str_format)

            # create output string
            start_outstr = "DTSTART;TZID={}:{}\n".format(tz, date_start_str)
            end_outstr = "DTEND;TZID={}:{}\n".format(tz, date_end_str)
            
            # write to output
            output.write(start_outstr)
            output.write(end_outstr)


            # --- DESCRIPTION --- #
            # write the zoom link as a description
            zoom_link = self.data_dict['zoom_link']
            description_outstr = "DESCRIPTION:Zoom link: {}\n".format(zoom_link)
            
            # output to file
            output.write(description_outstr)


            # cancel event if the cancel flag is set
            if self.data_dict['cancel'] == "True":
                method_outstr = "METHOD:CANCEL\n"
                status_outstr = "STATUS:CANCELLED\n"
                
                output.write(method_outstr)
                output.write(status_outstr)


            output.write("END:VEVENT\n")


        # close the vcalendar object
        output.write("END:VCALENDAR\n")

        # close the output file
        output.close()



def main():
    # choose whether we should have the special meetings set up

    spreadsheet = Spreadsheet() 
    # exit()

    # read input file
    input_file = "lab_calendar_data.txt"

    # read in the calendar data
    cal_data = Calendar(input_file)

    # for testing
    cal_data.print_datadict()

    # create the schedule
    cal_data.create_schedule()

    # write schedule to ics file
    cal_data.write_ics()


    # write the cancel file also
    cal_data.data_dict['cancel'] = "True"
    cal_data.write_ics()



if __name__ == '__main__':
    main()
