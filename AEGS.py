import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import os
import os.path
from PIL import Image, ImageDraw, ImageFont


def start_generate():
    print("\nPlease ensure the raw certificate image is in the same directory")
    image_name = input("Enter the file name of the image\n\n")
    try:
        Image.open(image_name)

        name_font_size = 118
        school_font_size = 53
        name_font = ImageFont.truetype("./fonts/PinyonScript-Regular.ttf", name_font_size)
        school_font = ImageFont.truetype("./fonts/march-rough.ttf", school_font_size)

        print("\nPlease enter the file name of the excel sheet")
        print("Ensure that it has 'name', 'school' and 'email' columns\n")
        file_name = input()
        df = pd.read_excel(file_name)
        if df.isnull().values.any():
            raise Exception("There are missing values")

        names = df["name"].tolist()
        schools = df["school"].tolist()

        generate(image_name, name_font, school_font, names, schools)

    except Exception as e:
        print(e)
        print("Please exit and try again~")
        while True:
            pass


def generate(image_name, name_font, school_font, names, schools):
    W = 2000
    print("Generating certificates...")

    if not os.path.exists("./generated_certificates"):
        os.mkdir("generated_certificates")

    for i in range(0, len(names)):
        name = names[i]
        school = schools[i]
        output = Image.open(image_name)
        image = ImageDraw.Draw(output)
        _, _, w1, h1 = image.textbbox((0, 0), name, font=name_font)
        _, _, w2, h2 = image.textbbox((0, 0), school.upper(), font=school_font)

        w_start_name = (W - w1) / 2
        w_start_school = (W - w2) / 2
        h_start_name = 619
        h_start_school = 809
        color = (30, 20, 20)

        image.text((w_start_name, h_start_name), name, font=name_font, fill=color)
        image.text((w_start_school, h_start_school), school.upper(), font=school_font, fill=color)

        output.save("./generated_certificates/{}_{}.jpg".format(name, school))
    print("{} Certificates generated in the generated_certificates folder. Bye~".format(len(names)))
    while True:
        pass


def check_certificates(names, schools):
    for i in range(0, len(names)):
        if not os.path.exists("./generated_certificates/{}_{}.jpg".format(names[i], schools[i])):
            return False
    return True


def start_send():
    print("\nBefore sending emails, ensure that the certs have been generated in the generated_certificates folder")
    print("Do not change the file names of the certificates\n")
    print("Please enter the gmail account that will send the emails\n")
    sender = input()
    print("""
Next, I need your app password. It's NOT the gmail account password.
To generate one, go to Google account -> Security -> 2-Step Verification -> App Passwords
-> Select App -> Other -> enter "Python" -> generate
    
Please enter the app password for the account.
    """)
    app_password = input()

    print("Logging in...")
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender, app_password)

        print("\nPlease enter the file name of the excel sheet")
        print("Ensure it has 'name', 'school' and 'email' columns\n")
        file_name = input()

        df = pd.read_excel(file_name)
        if df.isnull().values.any():
            raise Exception("There are missing values")

        names = df["name"].tolist()
        emails = df["email"].tolist()
        schools = df["school"].tolist()

        print("\nChecking whether the certificate images have been generated...")
        if not check_certificates(names, schools):
            print("Not all the certificates have been generated, please exit and check again~")
            while True:
                pass
        else:
            print("\nYou are good to go. Please check:\n")
            print("Email Subject:  Cosmic Cup Certificate")
            print("Email Message:  Thank you for participating in Cosmic Cup, we have attached your certificate below.")
            print("Attachment name:  Cosmic_Cup_Cert_{name}.jpg")
            print("\nType 'yes' to confirm sending the emails\n")
            if input() == "yes":
                send_email(server, sender, names, schools, emails)
            else:
                print("\nYour command is not confirmed. Come back when ready!")

    except Exception as e:
        print(e)
        server.quit()
        print("Please exit and try again~")
        while True:
            pass

    finally:
        server.quit()


def send_email(server, sender, names, schools, emails):
    subject = "Cosmic Cup Certificate"
    text = "Thank you for participating in Cosmic Cup, we have attached your certificate below."

    print("Sending emails...")
    counter = 0
    for i in range(0, len(names)):
        try:
            msg = MIMEMultipart()
            msg['Subject'] = subject
            msg['From'] = sender
            msg['To'] = emails[i]
            msg.attach(MIMEText(text, "plain"))

            filename = "./generated_certificates/{}_{}.jpg".format(names[i], schools[i])
            attachment_name = "Cosmic_Cup_Cert_{}.jpg".format(names[i])

            with open(filename, "rb") as attachment:
                mime = MIMEBase('image', 'jpg', filename=attachment_name)
                mime.add_header('Content-Disposition', 'attachment', filename=attachment_name)
                mime.add_header('X-Attachment-Id', '0')
                mime.add_header('Content-ID', '<0>')
                mime.set_payload(attachment.read())
                encoders.encode_base64(mime)
                msg.attach(mime)

            server.sendmail(sender, emails[i], msg.as_string())
            counter += 1
            print("Email sent to {}".format(names[i]))
        except Exception as e:
            print()
            print(e)
            print("******Error at row {}: {} from {}******\n".format(i + 2, names[i], schools[i]))

    server.quit()
    print("{} / {} Emails sent! Bye~".format(counter, len(emails)))
    while True:
        pass


def get_first_command():
    invalid_command_msg = "\nInvalid command, please try again.\n"

    first_command_msg = "Type 'generate' to start generating the certificates\n" + \
                        "Type 'send' to send the certificates by email\n" + \
                        "Type 'exit' to quit program\n"

    print(first_command_msg)
    first_command = input()
    if first_command == "generate":
        start_generate()
    elif first_command == "send":
        start_send()
    elif first_command == "exit":
        pass
    else:
        print(invalid_command_msg)
        get_first_command()


def main():
    welcome_msg = \
        """
Automatically generate and send certificates!
Author: Luo Jiale
        
- Use this program in a directory with the "fonts" folder.
- In the same directory, store recipient particulars in an excel sheet
     - The excel sheet should have three columns: "name", "school" and "email".
     - It should not have any missing values
- In the same directory, put the raw certificate image
    
        """
    print(welcome_msg)
    get_first_command()


if __name__ == "__main__":
    main()

