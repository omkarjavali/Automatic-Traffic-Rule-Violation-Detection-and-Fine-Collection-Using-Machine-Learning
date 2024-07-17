import os
import os.path
import glob
import csv
from datetime import datetime


def generate_excel():

    listFiles = glob.glob('output/*.txt')

    # Get current date and time in format yyyy-mm-dd_HH-MM-SS
    now = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')

    # Open new file with current date and time as filename
    filename = f'excels/{now}_output.csv'
    with open(filename, 'w', newline='') as file:
        writer = csv.writer(file)

        # Write headers for columns
        writer.writerow(['Helmet Status', 'LP Status', 'LP Number', 'No of passengers'])

        # Modify data and write rows to file
        for file in listFiles:
            try:
                data = open(file).read().split('\n')

                # Determine helmet status
                if data[1] == "None":
                    helmet_status = "Not Found"
                elif data[1] == "True":
                    helmet_status = "Wearing"
                elif data[1] == "False":
                    helmet_status = "Not Wearing"
                else:
                    helmet_status = ""

                # Determine LP status and number
                if data[2] == "None":
                    lp_status = "Not Found"
                    lp_number = "Not Found"
                elif data[2] == "True":
                    lp_status = "Found"
                    lp_number = data[3]
                elif data[2] == "False":
                    lp_status = "Not Found"
                    lp_number = "Not Found"
                else:
                    lp_status = ""
                    lp_number = "Not Found"

                # Determine number of passengers
                num_passengers = data[4] if data[4] != "None" else ""

                # Write row to file
                writer.writerow([helmet_status, lp_status, lp_number, num_passengers])

            except Exception as e:
                print(e)
