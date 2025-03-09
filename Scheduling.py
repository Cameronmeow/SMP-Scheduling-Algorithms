import pandas as pd
from datetime import datetime, timedelta
import logging

# Configure logging
logging.basicConfig(
    filename='schedule_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(message)s'
)

# Weekend time slots remain unchanged
weekend_time_slots_sat = [
    "Saturday 9:30AM-10:30AM", "Saturday 10:30AM-11:30AM", "Saturday 11:30AM-12:30PM",
    "Saturday 12:30PM-2PM", "Saturday 2PM-3:30PM", "Saturday 3:30PM-5PM",
    "Saturday 5:30PM-7:00PM", "Saturday 7PM-8:30PM", "Saturday 8:30PM-10:00PM",
    "Saturday 10:00PM-11:30PM", "Saturday 11:30PM-1:00AM"
]

weekend_time_slots_sun = [
    "Sunday 9:30AM-10:30AM", "Sunday 10:30AM-11:30AM", "Sunday 11:30AM-12:30PM",
    "Sunday 12:30PM-2PM", "Sunday 2PM-3:30PM", "Sunday 3:30PM-5PM",
    "Sunday 5:30PM-7:00PM", "Sunday 7PM-8:30PM", "Sunday 8:30PM-10:00PM",
    "Sunday 10:00PM-11:30PM", "Sunday 11:30PM-1:00AM"
]

# Dictionary mapping weekdays and time ranges to interviewer letters
weekday_interviewers = {
    "Monday": {
         "9:30AM-10:30AM": "H, A",
         "10:30AM-11:30AM": "",            # No interviewers
         "11:30AM-12:30PM": "H, A",
         "12:30PM-2PM": "H, S, A",
         "2PM-3:30PM": "S, A",
         "3:30PM-5PM": "S, A",
         "5:30PM-7PM": "",                # No interviewers
         "7PM-8:30PM": "A, S"
    },
    "Tuesday": {
         "9:30AM-10:30AM": "",
         "10:30AM-11:30AM": "",
         "11:30AM-12:30PM": "H, S",
         "12:30PM-2PM": "H, S",
         "2PM-3:30PM": "A",
         "3:30PM-5PM": "",
         "5:30PM-7PM": "H, A",
         "7PM-8:30PM": "H, A"
    },
    "Wednesday": {
         "9:30AM-10:30AM": "H, S",
         "10:30AM-11:30AM": "H, A",
         "11:30AM-12:30PM": "",
         "12:30PM-2PM": "H, S, A",
         "2PM-3:30PM": "H, S, A",
         "3:30PM-5PM": "H, S, A",
         "5:30PM-7PM": "H, S, A",
         "7PM-8:30PM": "H, S, A"
    },
    "Thursday": {
         "9:30AM-10:30AM": "H, A",
         "10:30AM-11:30AM": "H, S",
         "11:30AM-12:30PM": "A",
         "12:30PM-2PM": "H, S",
         "2PM-3:30PM": "H, S",
         "3:30PM-5PM": "H, S",
         "5:30PM-7PM": "",
         "7PM-8:30PM": "A"
    },
    "Friday": {
         "9:30AM-10:30AM": "H, S",
         "10:30AM-11:30AM": "H, S",
         "11:30AM-12:30PM": "S, A",
         "12:30PM-2PM": "H, S, A",
         "2PM-3:30PM": "A",
         "3:30PM-5PM": "",
         "5:30PM-7PM": "H, A",
         "7PM-8:30PM": "H, A"
    }
}

def load_data(file_path):
    """Load the dataset and filter relevant columns."""
    try:
        data = pd.read_excel(file_path)
        # Identify weekday time slots from the dataset columns
        time_slots = [col for col in data.columns if any(day in col for day in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'])]
        # Include 'Contact Number' column along with other required fields
        relevant_columns = [
            'First Name', 'Last Name', 'Email ID', 'Roll Number', 
            'Department', 'Contact Number', 'Interview Happened'
        ] + time_slots
        filtered_data = data[relevant_columns].fillna(0)
        
        # Ensure time slot values are binary (1 if available, else 0)
        for slot in time_slots:
            filtered_data[slot] = filtered_data[slot].apply(lambda x: 1 if x == 1 else 0)
        
        # Filter out candidates who have already been interviewed
        filtered_data = filtered_data[filtered_data['Interview Happened'] == 0]
        
        return filtered_data, time_slots
    except Exception as e:
        logging.error(f"Error loading data: {e}")
        return None, None

def assign_interview_slots(filtered_data, time_slots, start_date):
    """
    Assign interview slots for a 7-day window (including weekends).
    For weekdays, assign interviewers from the weekday_interviewers dictionary.
    Also, for weekdays, find up to 3 additional attendees (with contact numbers) available in the same slot.
    """
    schedule = []
    used_slots = {}
    end_date = start_date + timedelta(days=6)  # 7-day window
    
    # Shuffle candidates to avoid scheduling bias
    shuffled_candidates = filtered_data.sample(frac=1).reset_index(drop=True)
    
    for index, candidate in shuffled_candidates.iterrows():
        assigned_slot = None
        assigned_date = None
        temp_date = start_date
        
        # Find an available slot within the 7-day window
        while not assigned_slot and temp_date <= end_date:
            day_of_week = temp_date.strftime("%A")
            if day_of_week in ['Saturday', 'Sunday']:
                available_slots = weekend_time_slots_sat if day_of_week == 'Saturday' else weekend_time_slots_sun
            else:
                # Use only slots where the candidate is available
                available_slots = [slot for slot in time_slots if day_of_week in slot and candidate[slot] == 1]
            
            for slot in available_slots:
                if temp_date not in used_slots:
                    used_slots[temp_date] = set()
                if slot not in used_slots[temp_date]:
                    assigned_slot = slot
                    assigned_date = temp_date
                    break
            
            if not assigned_slot:
                temp_date += timedelta(days=1)
        
        if assigned_slot:
            # For weekdays, assign interviewers based on the provided schedule
            interviewers_list = []
            day_of_week = assigned_date.strftime("%A")
            if day_of_week in weekday_interviewers:
                # Extract the time portion from the slot string (e.g., "Monday 9:30AM-10:30AM" -> "9:30AM-10:30AM")
                time_range = assigned_slot.split(" ", 1)[1] if " " in assigned_slot else assigned_slot
                interviewers_str = weekday_interviewers.get(day_of_week, {}).get(time_range, "")
                interviewers_list = [x.strip() for x in interviewers_str.split(',')] if interviewers_str else []
            
            # For weekdays, find up to 3 additional attendees available in the same slot
            if day_of_week not in ['Saturday', 'Sunday']:
                additional_attendees = filtered_data[
                    (filtered_data[assigned_slot] == 1) &
                    (~filtered_data.index.isin([index] + [entry['Interviewee Index'] for entry in schedule]))
                ].iloc[:3]
            else:
                additional_attendees = pd.DataFrame()  # No additional attendees on weekends
            
            # Build the schedule entry including interviewer and additional attendee details
            entry = {
                "Interviewee Index": index,
                "Interviewee First Name": candidate["First Name"],
                "Interviewee Last Name": candidate["Last Name"],
                "Interviewee Email ID": candidate["Email ID"],
                "Interviewee Roll Number": candidate["Roll Number"],
                "Interviewee Department": candidate["Department"],
                "Interviewee Contact Number": candidate["Contact Number"],
                "Interview Date": assigned_date,
                "Interview Time": assigned_slot,
                "Interviewer 1": interviewers_list[0] if len(interviewers_list) > 0 else None,
                "Interviewer 2": interviewers_list[1] if len(interviewers_list) > 1 else None,
                "Interviewer 3": interviewers_list[2] if len(interviewers_list) > 2 else None,
                "Additional Person 1": None,
                "Additional Person 1 Email": None,
                "Additional Person 1 Contact": None,
                "Additional Person 2": None,
                "Additional Person 2 Email": None,
                "Additional Person 2 Contact": None,
                "Additional Person 3": None,
                "Additional Person 3 Email": None,
                "Additional Person 3 Contact": None,
            }
            
            # Add up to 3 additional attendees (only for weekdays)
            for i, (_, attendee) in enumerate(additional_attendees.iterrows(), start=1):
                entry[f"Additional Person {i}"] = f"{attendee['First Name']} {attendee['Last Name']}"
                entry[f"Additional Person {i} Email"] = attendee["Email ID"]
                entry[f"Additional Person {i} Contact"] = attendee["Contact Number"]
            
            # Mark the slot as used on the assigned date
            used_slots[assigned_date].add(assigned_slot)
            
            # Add the entry to the schedule list
            schedule.append(entry)
        else:
            logging.warning(f"Could not schedule interview for {candidate['First Name']} {candidate['Last Name']}")
    
    return pd.DataFrame(schedule)

def save_schedule(schedule_df, output_file):
    """Save the interview schedule to a CSV file."""
    try:
        schedule_df['Interview Date'] = pd.to_datetime(schedule_df['Interview Date'])
        schedule_df = schedule_df.sort_values(['Interview Date', 'Interview Time'])
        schedule_df = schedule_df.drop(columns=['Interviewee Index'])
        schedule_df.to_csv(output_file, index=False)
        logging.info(f"Schedule saved successfully to {output_file}")
    except Exception as e:
        logging.error(f"Error saving schedule: {e}")

# Main Execution
if __name__ == "__main__":
    file_path = './test_with_colors.xlsx'
    
    # Determine the start date: next Monday if today is not Monday
    current_date = datetime.now().date()
    if current_date.weekday() != 0:
        days_until_monday = (7 - current_date.weekday()) % 7
        start_date = current_date + timedelta(days=days_until_monday)
    else:
        start_date = current_date
    
    output_file = f'Schedule/weekly_interview_schedule_{start_date.strftime("%Y-%m-%d")}.csv'
    
    filtered_data, time_slots = load_data(file_path)
    if filtered_data is not None:
        schedule_df = assign_interview_slots(filtered_data, time_slots, start_date)
        save_schedule(schedule_df, output_file)
        print(f"Weekly interview scheduling complete for week starting {start_date}. Check the CSV file and log for details.")
