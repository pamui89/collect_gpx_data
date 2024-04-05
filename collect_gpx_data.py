
import glob
import gpxpy
import gpxpy.gpx
import pandas as pd
import os
from tqdm import tqdm
from openpyxl import Workbook
from openpyxl.styles import NamedStyle
from datetime import datetime

def parse_gpx(file_path):
    with open(file_path, 'r') as gpx_file:
        gpx = gpxpy.parse(gpx_file)
        
        for track in gpx.tracks:
            for segment in track.segments:
                competitor_id = os.path.splitext(os.path.basename(file_path))[0]
                start_time = datetime.now()
                finish_time = datetime.now()
                # start_time = segment.points[0].time.replace(tzinfo=None)
                # finish_time = segment.points[-1].time.replace(tzinfo=None)
                distance2d = segment.length_2d()/1000  # Distance in kilometers
                distance3d = segment.length_3d()/1000  # Distance in kilometers
                elapsed_time = (finish_time - start_time).total_seconds()
                
                yield competitor_id, start_time, finish_time, elapsed_time, distance2d, distance3d

timestamp_str = datetime.now().strftime('%Y%m%d_%H%M%S')
output_file = f'/Users/pamui/Documents/Developer/PythonPrograms/collect-gpx-data/results/race_results_{timestamp_str}.xlsx'
def write_to_excel(data, output_file=output_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "Race Data"

    # Headers
    ws.append(['Competitor Id', 'Start Time', 'Finish Time', 'Total time', 'Total Distance 2d', 'Total Distance 3d'])

    # Data
    for row in data:
        ws.append(row)

    # Formatting duration as [h]:mm:ss
    duration_style = NamedStyle(name='duration', number_format='[h]:mm:ss')
    for cell in ws['D'][1:]:  # Column D contains durations in seconds
        # Convert seconds to Excel time format (fraction of 24 hours)
        cell.value = cell.value / 86400  # There are 86400 seconds in a day
        cell.style = duration_style

    wb.save(output_file)

# Path where your GPX files are stored
gpx_files = glob.glob('/Users/pamui/Documents/Developer/PythonPrograms/collect-gpx-data/assets/*.gpx')

# Collect data from each file
all_data = []
for file_path in tqdm(gpx_files, desc= 'Processing GPX Files'):
    all_data.extend(parse_gpx(file_path))

# Write the collected data to an Excel file
write_to_excel(all_data)
