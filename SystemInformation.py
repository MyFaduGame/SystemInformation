from email.errors import MisplacedEnvelopeHeaderDefect
import os
import platform
import sqlite3
import datetime
import subprocess
import pytz
import win32file  # Ensure pywin32 is installed

# Output file
f = open("output.txt", "a")

def get_system_info():
    try:
        # Get the local hostname
        hostname = platform.node()

        # Get system information
        system_info = platform.uname()

        # Retrieve Original Install Date and System Manufacturer using wmic
        query_install_date = "wmic os get InstallDate /format:list"
        query_manufacturer = "wmic computersystem get Manufacturer /format:list"

        result_install_date = subprocess.check_output(query_install_date, shell=True).decode().strip().split("=")[1]
        result_manufacturer = subprocess.check_output(query_manufacturer, shell=True).decode().strip().split("=")[1]

        # Format Original Install Date as YYYY-MM-DD
        original_install_date = result_install_date[:8]
        
        original_install_date = datetime.datetime.strptime(original_install_date, "%Y%m%d").strftime("%Y-%m-%d")

        return {
            "Hostname": hostname,
            "System": system_info.system,
            "Node Name": system_info.node,
            "Release": system_info.release,
            "Version": system_info.version,
            "Machine": system_info.machine,
            "Processor": system_info.processor,
            "User Name": os.getlogin(),
            "Original Install Date": original_install_date,
            "System Manufacturer": result_manufacturer,
        }
    except Exception as e:
        return f"System Information Error: {str(e)}"

def convert_filetime_to_datetime(filetime):
    # Convert nanoseconds to seconds
    timestamp_in_seconds=13342260003944535
    # timestamp_in_nanoseconds = filetime /1000
    # timestamp_in_seconds = timestamp_in_nanoseconds/1e4
    # print(timestamp_in_seconds)

    # Create a datetime object
    # date_time = datetime.datetime.fromtimestamp(timestamp_in_seconds)

    # Format the datetime object as a string
    epoch = datetime.datetime(1601,1,1,tzinfo=pytz.UTC)
    formated_datetime = epoch + datetime.timedelta(microseconds=filetime)
    print(formated_datetime)

    # formatted_date_time = datetime.datetime.strptime(str(formatted_date_time),'%Y-%m-%d %H:%M:%S')
    # formatted_date_time = timestamp_in_seconds.strftime('%Y-%m-%d %H:%M:%S')
    # filetime=datetime.datetime.fromtimestamp()
    # filetime=datetime.datetime.strptime(filetime,"%d%m%y%H%M%S")
    # print(formatted_date_time,type(formatted_date_time))
    

    # microseconds = filetime / 10
    # 13340202316007148
    # # epoch = datetime.datetime(1601, 1, 1, tzinfo=pytz.utc)
    # epoch = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
    # # dt = epoch + datetime.timedelta(microseconds=microseconds)
    # # print(dt)
    # # ist = pytz.timezone('Asia/Kolkata')
    # # print(ist)
    # # dt_ist = dt.astimezone(ist)
    return formated_datetime

def get_chrome_history():
    try:
        # Find the user's Chrome history database
        history_db_path = os.path.expanduser("~") + r"\AppData\Local\Google\Chrome\User Data\Default\History"

        # Connect to the Chrome history database (SQLite)
        conn = sqlite3.connect(history_db_path)
        cursor = conn.cursor()

        # Query to retrieve browsing history
        cursor.execute("SELECT url, title, visit_time FROM urls ORDER BY visit_time DESC LIMIT 10")

        # Fetch and print the browsing history
        rows = cursor.fetchall()
        print("Recent Browsing History for Google Chrome:", file=f)
        for row in rows:
            url, title, visit_time = row
            visit_date = convert_filetime_to_datetime(visit_time)
            print(f"{visit_date}: {title} - {url}", file=f)

        # Close the database connection
        conn.close()

    except Exception as e:
        print(f"Error accessing Google Chrome browsing history: {str(e)}", file=f)

def get_firefox_history():
    try:
        # Find the user's Firefox history database
        app_data_path = os.path.expanduser("~") + r"\AppData\Roaming\Mozilla\Firefox\Profiles"

        # Find the Firefox profile directory
        profile_dir = None
        for folder_name in os.listdir(app_data_path):
            if folder_name.endswith(".default"):
                profile_dir = folder_name
                break

        if profile_dir is not None:
            history_db_path = os.path.join(app_data_path, profile_dir, "places.sqlite")

            # Connect to the Firefox history database (SQLite)
            conn = sqlite3.connect(history_db_path)
            cursor = conn.cursor()

            # Query to retrieve browsing history
            cursor.execute("SELECT url, title, visit_date FROM moz_places ORDER BY visit_date DESC LIMIT 10")

            # Fetch and print the browsing history
            rows = cursor.fetchall()
            print("Recent Browsing History for Mozilla Firefox:", file=f)
            for row in rows:
                url, title, visit_date = row
                print(visit_date,type(visit_date))
                visit_time = convert_filetime_to_datetime(visit_date)
                print(f"{visit_time}: {title} - {url}", file=f)

            # Close the database connection
            conn.close()
        else:
            print("Firefox profile directory not found.", file=f)

    except Exception as e:
        print(f"Error accessing Mozilla Firefox browsing history: {str(e)}", file=f)

def get_edge_history():
    try:
        # Find the user's Edge history database
        history_db_path = os.path.expanduser("~") + r"\AppData\Local\Microsoft\Edge\User Data\Default\History"

        # Connect to the Edge history database (SQLite)
        conn = sqlite3.connect(history_db_path)
        cursor = conn.cursor()

        # Query to retrieve browsing history
        cursor.execute("SELECT url, title, last_visit_time FROM urls ORDER BY last_visit_time DESC LIMIT 10")

        # Fetch and print the browsing history
        rows = cursor.fetchall()
        print("Recent Browsing History for Microsoft Edge:", file=f)
        for row in rows:
            url, title, last_visit_time = row
            print(last_visit_time,type(last_visit_time))
            visit_time = convert_filetime_to_datetime(last_visit_time)
            print(f"{visit_time}: {title} - {url}", file=f)

        # Close the database connection
        conn.close()

    except Exception as e:
        print(f"Error accessing Microsoft Edge browsing history: {str(e)}", file=f)

if __name__ == "__main__":
    # Gather system information
    system_info = get_system_info()
    if isinstance(system_info, dict):
        print("System Information:", file=f)
        for key, value in system_info.items():
            print(f"{key}: {value}", file=f)
    else:
        print("Failed to retrieve system information.", file=f)

    # Get browsing history from different browsers
    get_chrome_history()
    get_firefox_history()
    get_edge_history()

f.close()
