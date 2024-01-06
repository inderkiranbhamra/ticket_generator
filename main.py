import yagmail
from barcode import Code128
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os


# Function to generate barcode for a given data and add it to the boarding pass
def generate_boarding_pass(data, passenger_name, gate_number, flight_date, departure_location, arrival_time,
                           destination, output_folder):
    # Create a folder to store participant-specific files
    os.makedirs(output_folder, exist_ok=True)

    # Create the barcode using Code128
    code = Code128(data, writer=ImageWriter())
    # Save the barcode image
    barcode_path = os.path.join(output_folder, f'barcode_{data}.png')
    code_image = code.render()

    # Rotate the barcode 90 degrees to the right
    code_image = code_image.rotate(-90, expand=True)

    code_image.save(barcode_path)

    # Open the boarding pass template
    boarding_pass_template_path = 'boarding_pass_template.jpg'  # Replace with your boarding pass template path
    boarding_pass = Image.open(boarding_pass_template_path)

    # Open the saved barcode image and resize
    barcode_image = Image.open(barcode_path)
    barcode_image = barcode_image.resize(
        (int(barcode_image.width * 0.5), barcode_image.height*2))  # Adjust the width scaling factor as needed

    # Calculate the position to place the barcode on the right side vertically
    x_offset = boarding_pass.width - barcode_image.width
    y_offset = 0

    # Paste the barcode onto the boarding pass
    boarding_pass.paste(barcode_image, (950, 200))

    # Use Draw to add text to the boarding pass
    draw = ImageDraw.Draw(boarding_pass)

    # Load a different font and set the font size
    font_path = "Browood-Regular.ttf"  # Replace with the path to your desired font file
    font_size = 30
    font = ImageFont.truetype(font_path, font_size)

    # Define text positions and content
    text_positions = [
        (50, 260, f"{passenger_name}"),
        (545, 260, f"{gate_number}"),
        (50, 440, f"{flight_date}"),
        (545, 440, f"{departure_location}"),
        (50, 610, f"{arrival_time}"),
        (545, 610, f"{destination}")
    ]

    # Add text to the boarding pass
    for position in text_positions:
        draw.text(position, f"{position[2]}", font=font, fill=(0, 0, 0))

    # Save the final boarding pass with the rotated barcode and text
    boarding_pass_path = os.path.join(output_folder, f'boarding_pass_{data}.png')
    boarding_pass.save(boarding_pass_path)

    return boarding_pass_path


# Function to send email with boarding pass using yagmail
def send_email(passenger_name, passenger_email, boarding_pass_path, flight_date, departure_location, arrival_time, destination):
    # Email configuration
    sender_email = 'inderkiran20233@gmail.com'  # replace with your email
    app_password = 'krhu cexv lyue dmnz'  # replace with your generated app password
    subject = f'Your Flight Ticket for {destination}'

    # Create yagmail SMTP client
    yag = yagmail.SMTP(sender_email, app_password)

    # Participant details in the email body
    email_body = f"Dear {passenger_name},\n\nThank you for choosing XYZ Airlines!\n\nHere is your flight ticket with the attached barcode:\n\n"
    email_body += f"Flight Date: {flight_date}\nDeparture Location: {departure_location}\nArrival Time: {arrival_time}\nDestination: {destination}"

    # Attach the boarding pass with barcode
    attachment = boarding_pass_path

    # Send the email
    yag.send(to=participant_email, subject=subject, contents=[email_body, attachment])


# Read participant data from Excel sheet
excel_file = 'dummy_participant_data.xlsx'  # replace with your Excel file
df = pd.read_excel(excel_file)

# Iterate through each participant and send emails
for index, row in df.iterrows():
    participant_name = row['Name']
    participant_email = row['Email']
    unique_id = str(row['UniqueID'])  # Ensure unique_id is converted to string

    # Output folder for participant-specific files
    output_folder = f'participant_files/{unique_id}'

    # Generate and attach boarding pass with rotated barcode and text
    boarding_pass_path = generate_boarding_pass(unique_id, row['Passenger_name'], row['Gate_number'],
                                                row['Flight_date'], row['From'], row['Arrival_time'],
                                                row['Destination'], output_folder)

    # Send email with boarding pass
    send_email(participant_name, participant_email, boarding_pass_path, row['Flight_date'], row['From'], row['Arrival_time'], row['Destination'])

print("Emails sent successfully.")
