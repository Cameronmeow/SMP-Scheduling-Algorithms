import pandas as pd
from datetime import datetime, timedelta
import logging

# Configure logging
logging.basicConfig(filename='schedule_log.txt', level=logging.INFO, format='%(asctime)s - %(message)s')

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

def load_data(file_path):
    """Load the dataset and filter relevant columns."""
    try:
        data = pd.read_excel(file_path)
        time_slots = [col for col in data.columns if any(day in col for day in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'])]
        relevant_columns = ['First Name', 'Last Name', 'Email ID', 'Roll Number', 'Department', 'Interview Happened'] + time_slots
        filtered_data = data[relevant_columns].fillna(0)  # Fill NaN values with 0
        
        # Ensure time slot values are binary (0 or 1)
        for slot in time_slots:
            filtered_data[slot] = filtered_data[slot].apply(lambda x: 1 if x == 1 else 0)
        
        # Filter out candidates who have already had an interview
        filtered_data = filtered_data[filtered_data['Interview Happened'] == 0]
        
        return filtered_data, time_slots
    except Exception as e:
        logging.error(f"Error loading data: {e}")
        return None, None

def assign_interview_slots(filtered_data, time_slots, start_date):
    """Assign interview slots for the whole week including weekends based on availability."""
    schedule = []
    used_slots = {}
    current_date = start_date
    end_date = current_date + timedelta(days=6)  # Schedule for 7 days (including weekend)
    
    # Shuffle the candidates to avoid bias in scheduling
    shuffled_candidates = filtered_data.sample(frac=1).reset_index(drop=True)
    
    # Iterate through each candidate
    for index, candidate in shuffled_candidates.iterrows():
        assigned_slot = None
        assigned_date = None

        while not assigned_slot and current_date <= end_date:
            day_of_week = current_date.strftime("%A")
            if day_of_week in ['Saturday', 'Sunday']:
                available_slots = weekend_time_slots_sat if day_of_week == 'Saturday' else weekend_time_slots_sun
            else:
                available_slots = [slot for slot in time_slots if day_of_week in slot and candidate[slot] == 1]
            
            for slot in available_slots:
                if current_date not in used_slots:
                    used_slots[current_date] = set()
                
                if slot not in used_slots[current_date]:
                    assigned_slot = slot
                    assigned_date = current_date
                    break
            
            if not assigned_slot:
                current_date += timedelta(days=1)

        if assigned_slot:
            # Find up to 3 additional people available in the same slot (only for weekdays)
            if assigned_date.strftime("%A") not in ['Saturday', 'Sunday']:
                additional_attendees = filtered_data[
                    (filtered_data[assigned_slot] == 1) &
                    (~filtered_data.index.isin([index] + [entry['Interviewee Index'] for entry in schedule]))
                ].iloc[:3]
            else:
                additional_attendees = pd.DataFrame()  # No additional attendees on weekends

            entry = {
                "Interviewee Index": index,
                "Interviewee First Name": candidate["First Name"],
                "Interviewee Last Name": candidate["Last Name"],
                "Interviewee Email ID": candidate["Email ID"],
                "Interviewee Roll Number": candidate["Roll Number"],
                "Interviewee Department": candidate["Department"],
                "Interview Date": assigned_date,
                "Interview Time": assigned_slot,
                "Additional Person 1": None,
                "Additional Person 1 Email": None,
                "Additional Person 2": None,
                "Additional Person 2 Email": None,
                "Additional Person 3": None,
                "Additional Person 3 Email": None,
            }

            # Add up to 3 additional attendees to the entry (only for weekdays)
            for i, (_, attendee) in enumerate(additional_attendees.iterrows(), start=1):
                entry[f"Additional Person {i}"] = f"{attendee['First Name']} {attendee['Last Name']}"
                entry[f"Additional Person {i} Email"] = attendee["Email ID"]

            # Mark the slot as used for the assigned date
            used_slots[assigned_date].add(assigned_slot)

            # Add the entry to the schedule
            schedule.append(entry)
        else:
            logging.warning(f"Could not schedule interview for {candidate['First Name']} {candidate['Last Name']}")

    return pd.DataFrame(schedule)

def save_schedule(schedule_df, output_file):
    """Save the interview schedule to a CSV file."""
    try:
        # Sort the schedule by date and time
        schedule_df['Interview Date'] = pd.to_datetime(schedule_df['Interview Date'])
        schedule_df = schedule_df.sort_values(['Interview Date', 'Interview Time'])
        
        # Remove the temporary 'Interviewee Index' column
        schedule_df = schedule_df.drop(columns=['Interviewee Index'])
        
        schedule_df.to_csv(output_file, index=False)
        logging.info(f"Schedule saved successfully to {output_file}")
    except Exception as e:
        logging.error(f"Error saving schedule: {e}")

# Main Execution
if __name__ == "__main__":
    file_path = './test_with_colors.xlsx'
    
    # Get the current date
    current_date = datetime.now().date()
    
    # If today is not Monday, find the next Monday
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
        print(f"Weekly interview scheduling (including weekend) complete for week starting {start_date}. Check the CSV file and log for details.")
