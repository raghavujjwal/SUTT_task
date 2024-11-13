import os
import json
import pandas as pd


workbook_path = r'C:\Users\ujjwal\OneDrive\Desktop\Projects\.vscode\SUTT_Task\Timetable Workbook - SUTT Task 1.xlsx'
workbook = pd.ExcelFile(workbook_path)


timetable_data = {}


for sheet_name in workbook.sheet_names:
    
    df = workbook.parse(sheet_name)
    
    
    course_code = df.iloc[0, 0]
    course_title = df.iloc[0, 1]
    credit_structure = df.iloc[0, 2]
    
    sections = []
    for i in range(1, len(df)):
        time_slots_str = str(df.iloc[i, 7])
        time_slots = []
        if time_slots_str != 'nan':
            for time_slot in time_slots_str.split(', '):
                try:
                    time_slots.append(int(time_slot))
                except ValueError:
                    
                    continue
        section = {
            'section_type': df.iloc[i, 3],
            'section_number': df.iloc[i, 4],
            'instructor': str(df.iloc[i, 5]),
            'room_number': df.iloc[i, 6],
            'time_slots': time_slots
        }
        sections.append(section)
    
    
    timetable_data[course_code] = {
        'course_title': course_title,
        'credit_structure': credit_structure,
        'sections': sections
    }


def handle_nan_values(value):
    if value == 'NaN':
        return None
    else:
        return value


for course_code, course_data in timetable_data.items():
    for section in course_data['sections']:
        section['section_type'] = handle_nan_values(section['section_type'])
        section['section_number'] = handle_nan_values(section['section_number'])
        
       
        if section['instructor'] is None or section['instructor'] == 'NaN':
            section['instructor'] = 'N/A'
        else:
            section['instructor'] = str(section['instructor'])
        
        section['room_number'] = handle_nan_values(section['room_number'])


with open('timetable_data.json', 'w') as f:
    json.dump(timetable_data, f, indent=4)

print('Timetable data extracted and saved to timetable_data.json')