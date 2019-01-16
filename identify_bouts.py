import csv
import sys
import xlsxwriter
from datetime import time
import xlrd
import numpy as np
from datetime import datetime, timedelta
import pytz
import matplotlib.pyplot as plt

'''

This script takes an XLSX of actigraphy data and sleep diary data and outputs a new spreadsheet, 
named output_[input-filename].xlsx with possible points sleep onset and offset and highlighted 
according to hierarchy.

Sample call: python identify_bouts.py SampleFakeData.xlsx

'''


def is_nan(num):
    if num == "NaN":
        return 1
    else:
        return 0


def get_epoch(date):
    return (date - datetime(1970, 1, 1)).total_seconds()


def epoch_to_str(ep):
    return datetime.fromtimestamp(ep, tz=pytz.utc).strftime('%m/%d/%Y       %I:%M %p')


def epoch_to_day(ep):
    return datetime.fromtimestamp(ep, tz=pytz.utc).strftime('%m/%d/%Y')


def epoch_to_time(ep):
    return datetime.fromtimestamp(ep, tz=pytz.utc).strftime('%I:%M:%S %p')


def add_one_day(day):
    return day + timedelta(days=1)


def find_closest(time_point, list):
    dist = 1e10
    closest = -1
    for list_point in list:
        dist_temp = abs(list_point - time_point)
        if dist_temp < dist:
            dist = dist_temp
            closest = list_point
    return closest


def find_concordance(i, j, dt):
    output = get_within_window(i, j, dt)
    if len(output) > 0:
        return output, i, j, dt
    else:
        if (j < 4):
            output, i, j, dt = find_concordance(i, j + 1, dt)
        else:
            if i < 3:
                output, i, j, dt = find_concordance(i + 1, i + 2, dt)
            else:
                if dt == 15:
                    output, i, j, dt = find_concordance(1, 2, 30)

    return output, i, j, dt


def get_within_window(i, j, dt):
    first_list = change_time_dictionary[i]
    second_list = change_time_dictionary[j]
    output = []
    for m in range(0, len(first_list)):
        for n in range(0, len(second_list)):
            tdelta = (first_list[m] - second_list[n])
            if abs(tdelta) / (60) < dt:
                output.append(first_list[m])

    return output


# Identify when sudden lux drop-offs occur
def lux_candidates(times, light, onset):
    width = 5
    output = []

    if onset == 1:
        for i in range(1, len(light) - width):
            valid = 1

            # Remove if previous entry was not in darkness.
            if light[i - 1] < 1:
                valid = 0

            # Remove if subsequent light
            for j in range(0, width):
                if light[i + j] >= 1:
                    valid = 0

            if valid == 1:
                output.append(times[i])
    else:
        for i in range(1, len(light) - width):
            valid = 1

            # Remove if previous entry was not in light.
            if light[i - 1] >= 1:
                valid = 0

            # Remove if subsequent darkness
            for j in range(0, width):
                if light[i + j] < 1:
                    valid = 0

            if valid == 1:
                output.append(times[i])

    return output


# Identify when sudden activity drop-offs occur
def sleep_wake_candidates(times, sw, onset):
    output = []
    width = 90

    thresh = 0.5
    if onset == 0:  # Sleep offset

        for i in range(1, len(sw) - width):
            valid = 1

            # Remove if first entry is not sleep
            if sw[i] > thresh:
                valid = 0

            # Remove if anything later in the window is not wake
            for j in range(1, width):
                if sw[i + j] <= thresh:
                    valid = 0

            if valid == 1:
                output.append(times[i])
    else:

        for i in range(1, len(sw) - width):
            valid = 1

            # Remove if first entry is not wake
            if sw[i] <= thresh:
                valid = 0

            for j in range(1, width):
                if sw[i + j] > thresh:
                    valid = 0
            if valid == 1:
                output.append(times[i])
    return output


# Identify when sudden activity drop-offs occur
def activity_candidates(times, act, onset):
    width = 5
    output = []

    if onset == 1:

        for i in range(1, len(act) - width):
            valid = 1

            # Remove if activity has already stopped
            if act[i - 1] == 0:
                valid = 0

            # Remove if activity occurs in the later times
            for j in range(0, width):
                if act[i + j] > 0:
                    valid = 0
            if valid == 1:
                output.append(times[i])
    else:

        for i in range(1, len(act) - width):
            valid = 1

            # Remove if activity has already stopped
            if (act[i - 1] > 0):
                valid = 0

            # Remove if activity occurs in the later times
            for j in range(0, width):
                if (act[i + j] == 0):
                    valid = 0
            if valid == 1:
                output.append(times[i])

    return output


# Identify when markers occur
def marker_candidates(times, mark):
    output = []
    for i in range(0, len(mark)):
        if mark[i] == 1:
            output.append(times[i])
    return output


# Initialize holders
dates = []
times = []
timestamps = []
activity = []
marker = []
white_light = []
sleep_wake = []
diary_onsets = []
diary_offsets = []
diary_dates = []

startIndex = 1e10
count = 0
collecting_diary = 0
diary_start = 1e10

# Load workbook for saving output spreadsheet
workbook = xlrd.open_workbook(sys.argv[1])
sheet = workbook.sheet_by_index(0)


if sys.argv[1][-4:] == "xlsx":
    for row_in_input in range(sheet.nrows):
        row = sheet.row_values(row_in_input)

        if len(row) > 0:
            if row[0] == "-------------------- Epoch-by-Epoch Data -------------------":
                startIndex = count + 14

            if row[2] == "SD REST INTERVAL END":
                diary_start = count
                collecting_diary = 1  # 1 if we are collecting diary data

            if count > diary_start and collecting_diary == 1:
                if len(row[1]) == 0:  # Set collecting_diary to 0 if we're done collecting the sleep diary data
                    collecting_diary = 0
                else:
                    diary_start_date = xlrd.xldate.xldate_as_datetime(row[0], workbook.datemode)
                    diary_dates.append(get_epoch(diary_start_date))

                    correctDate = True
                    try:
                        diary_onset_time = datetime.strptime(row[1], "%I:%M%p")
                        diary_offset_time = datetime.strptime(row[2], "%I:%M%p")
                    except ValueError:
                        correctDate = False

                    if correctDate:
                        midnight = diary_onset_time.replace(hour=0, minute=0, second=0, microsecond=0)
                        onset_seconds = (diary_onset_time - midnight).seconds

                        diary_onset_time = diary_onset_time.time()
                        diary_offset_time = diary_offset_time.time()

                        # Important: This is a check to add one day if the sleep diary date doesn't match the time
                        # Should be replaced; can cause issues with shift workers
                        # E.g. 1/1/2019 Sleep Onset at 1:00AM meaning 1:00AM on 1/2/2019
                        # TODO: Change formatting on inputs so dates and times always reflect true date/time

                        if onset_seconds > 16 * 60 * 60:  # If evening, use date
                            comb_onset_time = datetime.combine(diary_start_date, diary_onset_time)
                        else:                             # If morning, advance day one.
                            comb_onset_time = datetime.combine(add_one_day(diary_start_date), diary_onset_time)

                        comb_offset_time = datetime.combine(add_one_day(diary_start_date), diary_offset_time)

                        diary_onset_timestamp = get_epoch(comb_onset_time)
                        diary_offset_timestamp = get_epoch(comb_offset_time)
                        diary_onsets.append(diary_onset_timestamp)
                        diary_offsets.append(diary_offset_timestamp)

            if count > startIndex:  # When we have reached the raw actigraphy data, collect and store in holders
                date = xlrd.xldate.xldate_as_datetime(row[1], workbook.datemode)

                x = int(row[2] * 24 * 3600)  # Convert to number of seconds
                timePoint = time(x // 3600, (x % 3600) // 60, x % 60)  # hours, minutes, seconds

                dates.append(date.strftime("%m/%d/%Y"))
                times.append(timePoint.strftime("%I:%M %p"))

                timestamp = get_epoch(date) + x
                timestamps.append(timestamp)

                if is_nan(row[3]) == 0:
                    activity.append(int(row[3]))
                else:
                    activity.append(-1)

                if is_nan(row[4]) == 0:
                    marker.append(int(row[4]))
                else:
                    marker.append(-1)

                if is_nan(row[5]) == 0:
                    white_light.append(float(row[5]))
                else:
                    white_light.append(-1)

                if is_nan(row[6]) == 0:
                    sleep_wake.append(int(row[6]))

                else:
                    sleep_wake.append(-1)

        count = count + 1

mark = np.array(marker_candidates(timestamps, marker))
diary_onsets = np.array(diary_onsets)
diary_offsets = np.array(diary_offsets)

print('Marker times')
for val in mark:
    print(epoch_to_str(val))

print('\nDiary onsets')
for val in diary_onsets:
    print(epoch_to_str(val))

print('\nDiary offsets')
for val in diary_offsets:
    print(epoch_to_str(val))

names = ['Marker', 'Diary', 'Light', 'Activity']

for onset in [0, 1]:  # When onset == 0, get sleep offset points; otherwise, get sleep onset points

    # Candidate points for each criteria: light, step,
    lc = np.array(lux_candidates(timestamps, white_light, onset))
    act = np.array(activity_candidates(timestamps, activity, onset))
    sw = np.array(sleep_wake_candidates(timestamps, sleep_wake, onset))

    if onset == 0:
        diary = diary_offsets
        print('\nSleep offset')

    else:
        diary = diary_onsets
        print('\nSleep onset')

    start = int(timestamps[0])
    end = int(timestamps[-1])
    delta = int(60 * 60 * 1)

    verbose = 0

    all_outputs = []
    all_ranks = []
    for day in range(start, end, delta):
        window = 15

        change_time_dictionary = dict()
        change_time_dictionary[1] = mark[np.logical_and(mark >= day, mark < day + delta + window * 60)]
        change_time_dictionary[2] = diary[np.logical_and(diary >= day, diary < day + delta + window * 60)]
        change_time_dictionary[3] = lc[np.logical_and(lc >= day, lc < day + delta + window * 60)]
        change_time_dictionary[4] = act[np.logical_and(act >= day, act < day + delta + window * 60)]

        if verbose:
            print('\nMarker change times')
            for val in change_time_dictionary[1]:
                print(epoch_to_str(val))
            print('\nDiary change times')
            for val in change_time_dictionary[2]:
                print(epoch_to_str(val))
            print('\nLight change times')
            for val in change_time_dictionary[3]:
                print(epoch_to_str(val))
            print('\nActivity change times')
            for val in change_time_dictionary[4]:
                print(epoch_to_str(val))

        concordance_output, i, j, dt = find_concordance(1, 2, window)

        if len(concordance_output) > 0:  # Print, respecting hierarchy
            val = concordance_output[0]
            if i == 1:
                print(epoch_to_str(val) + ' ---- ' + names[i - 1] + ', ' + names[j - 1] + ', dt = ' + str(dt))
            if i == 2:
                print(epoch_to_str(val) + ' --------  ' + names[i - 1] + ', ' + names[j - 1] + ', dt = ' + str(dt))
            if i == 3:
                print(epoch_to_str(val) + ' ---------------- ' + names[i - 1] + ', ' + names[j - 1] + ', dt = ' + str(
                    dt))

            all_outputs.append(val)
            all_ranks.append([i, j])

    if onset == 1:
        all_onsets = np.array(all_outputs)
        all_onset_ranks = all_ranks
    else:
        all_offsets = np.array(all_outputs)
        all_offset_ranks = all_ranks

compiled_ranks = all_onset_ranks + all_offset_ranks
compiled_values = np.concatenate((all_onsets, all_offsets))
sorted_values = np.argsort(compiled_values)

# Create a workbook and add a worksheet.
worksheet_save_name = 'output_' + sys.argv[1][0:-4] + 'xlsx'
print('Worksheet save name: ' + worksheet_save_name)
workbook = xlsxwriter.Workbook(worksheet_save_name)

worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True})

worksheet.write('A1', 'Date', bold)
worksheet.write('B1', 'Candidate Rest Int Onset Time', bold)
worksheet.write('C1', 'Candidate Rest Int Offset Time', bold)
worksheet.write('D1', 'Marker is highest ranked', bold)
worksheet.write('E1', 'Diary is highest ranked', bold)
worksheet.write('F1', 'Light is highest ranked', bold)

onset_format = workbook.add_format()
onset_format.set_pattern(1)
onset_format.set_bg_color('#C1E1DC')
onset_format.set_bold(True)
onset_format.set_font_name('Times New Roman')

offset_format = workbook.add_format()
offset_format.set_pattern(1)
offset_format.set_bg_color('#FFEB94')
offset_format.set_bold(True)
offset_format.set_font_name('Times New Roman')

light_format = workbook.add_format()
light_format.set_pattern(1)
light_format.set_bg_color('#FDD475')
light_format.set_font_name('Times New Roman')

diary_format = workbook.add_format()
diary_format.set_pattern(1)
diary_format.set_bg_color('#FFCCAC')
diary_format.set_font_name('Times New Roman')

marker_format = workbook.add_format()
marker_format.set_pattern(1)
marker_format.set_bg_color('#F79B77')
marker_format.set_font_name('Times New Roman')

row = 1
col = 0

if len(diary_onsets) > 0:
    min_date = np.min([diary_onsets[0], diary_offsets[0]]) - 24 * 3600
    max_date = np.max([diary_onsets[-1], diary_offsets[-1]]) + 24 * 3600
else:
    min_date = np.min([diary_dates[0], diary_dates[0]]) - 24 * 3600
    max_date = np.max([diary_dates[-1], diary_dates[-1]]) + 24 * 3600

for item in range(0, len(sorted_values)):
    index = sorted_values[item]
    on_off_string = ''
    if index < len(all_onsets):
        is_onset = 1
        on_off_string = 'ONSET'
        on_off_format = onset_format

    else:
        is_onset = 0
        on_off_string = 'OFFSET'
        on_off_format = offset_format

    i = compiled_ranks[index][0]
    j = compiled_ranks[index][1]
    val = compiled_values[index]

    if val >= min_date and val < max_date:
        rank_string = names[i - 1] + ', ' + names[j - 1]

        worksheet.write(row, col, epoch_to_day(val))
        if is_onset == 1:
            worksheet.write(row, col + 1, epoch_to_time(val), on_off_format)
        else:
            worksheet.write(row, col + 2, epoch_to_time(val), on_off_format)

        worksheet.set_column(0, 2, 15)
        worksheet.set_column(3, 5, 25)

        if i == 1:
            worksheet.write(row, col + 3, rank_string, marker_format)
        if i == 2:
            worksheet.write(row, col + 4, rank_string, diary_format)
        if i == 3:
            worksheet.write(row, col + 5, rank_string, light_format)
        row += 1

workbook.close()
