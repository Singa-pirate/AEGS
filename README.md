# Automatic Email Generator & Sender

Author: Luo Jiale

### Contents

- One executable file
- A "fonts" folder containing two fonts
- A raw certificate image
- An excel sheet containing "name", "school" and "email" columns

### Description

This program serves two functions:

1. Generate certificates by overlaying name and school name on the raw cert image
2. From a Gmail account, send the generated certificates to corresponding email addresses

### Usage

- Ensure all the required files and folders are in the same directory
  - The raw cert image should have the same dimensions as the example image
  - The excel sheet should have the required column names, without any missing value
  - Try your best to obtain valid email addresses (invalid ones will be informed later, but you'll have to manually deal with them)

- Run the .exe file and follow the instructions

  - DO NOT change file names / edit file content when the program is running

- At the end, you should see something like this:

  ```
  Type 'yes' to confirm sending the emails
  
  yes
  Sending emails...
  Email sent to 4F5DA2
  
  {'not_even_an_email_gmail_com': (553, b'5.1.3 The recipient address <not_even_an_email_gmail_com> is not a valid\n5.1.3 RFC-5321 address. Learn more at\n5.1.3  https://support.google.com/mail/answer/6596 s22-20020a63dc16000000b00502f4c62fd3sm6801573pgg.33 - gsmtp')}
  ******Error encountered at row 3: Person Whose Email Address is NOT an email from test1******
  
  Email sent to Person Whose Email Address is a wrong address
  Email sent to Very Very Very Very Looooooooooooooooong Name
  3 / 4 Emails sent! Bye~
  ```

  - 3 of the 4 emails were successfully sent.
  - If an email is **not in the format of an email address**, then the error will be displayed here
  - If an email is an email address but the **address is not found because it's wrong**, there will be no error here. **Check your mailbox** to see whether such errors happened.



### Note

##### **This program is specifically written for Cosmic Cup 2023 (https://www.cosmiccup.org/). For other events, the code needs to be changed due to different**

- Certificate dimensions
- Certificate content & font
- Email sender domain (Gmail is used in this program)
- Email content