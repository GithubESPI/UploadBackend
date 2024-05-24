def format_duration_to_minutes(duration_str):
    parts = duration_str.split('h')
    if len(parts) == 2:
        hours = int(parts[0])
        minutes = int(parts[1])
        return hours * 60 + minutes
    return int(duration_str.split()[0])

def format_minutes_to_duration(minutes):
    if minutes == 0:
        return "00h00"
    hours = minutes // 60
    remaining_minutes = minutes % 60
    if hours > 0:
        return f"{hours}h{remaining_minutes:02d}"
    return f"{remaining_minutes} minutes"

def sum_durations(duration_list):
    total_minutes = sum(duration_list)
    return format_minutes_to_duration(total_minutes)
