import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def send_bulk_emails():
    # Get file path from user
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    
    if not file_path:
        return
    
    try:
        # Read Excel file
        df = pd.read_excel(file_path)
        
        # Print DataFrame to check its structure and contents
        print("DataFrame from Excel file:")
        print(df)

        if df.empty:
            messagebox.showwarning("Warning", "Excel file is empty. Please ensure it contains data.")
            return

        # Sender's credentials
        my_mail = entry_email.get()
        passcode = entry_password.get()

        # Email content
        subject = entry_subject.get()
        body = entry_body.get("1.0", tk.END)

        # Connect to Gmail SMTP server
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(my_mail, passcode)

        # Send email to each recipient
        for index, row in df.iterrows():
            recipient = row['mohitpardhi2003@gmail.com']
            print("Sending email to:", recipient)
            message = MIMEMultipart()
            message['From'] = my_mail
            message['To'] = recipient
            message['Subject'] = subject
            message.attach(MIMEText(body, 'plain'))
            server.sendmail(my_mail, recipient, message.as_string())

        # Close SMTP connection
        server.quit()

        messagebox.showinfo("Success", "Bulk emails sent successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Something went wrong: {e}")

# Create GUI window
window = tk.Tk()
window.title("Bulk Email Sender")

# Email Entry
label_email = tk.Label(window, text="Your Email:")
label_email.grid(row=0, column=0, padx=5, pady=5)
entry_email = tk.Entry(window)
entry_email.grid(row=0, column=1, padx=5, pady=5)

# Password Entry
label_password = tk.Label(window, text="Your Password:")
label_password.grid(row=1, column=0, padx=5, pady=5)
entry_password = tk.Entry(window, show="*")
entry_password.grid(row=1, column=1, padx=5, pady=5)

# Subject Entry
label_subject = tk.Label(window, text="Subject:")
label_subject.grid(row=2, column=0, padx=5, pady=5)
entry_subject = tk.Entry(window)
entry_subject.grid(row=2, column=1, padx=5, pady=5)

# Body Entry
label_body = tk.Label(window, text="Body:")
label_body.grid(row=3, column=0, padx=5, pady=5)
entry_body = tk.Text(window, height=5, width=30)
entry_body.grid(row=3, column=1, padx=5, pady=5)

# Upload Button
upload_button = tk.Button(window, text="Upload Excel", command=send_bulk_emails)
upload_button.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="WE")

# Run the GUI window
window.mainloop()



# my_mail = "mohitpardhi2003@gmail.com"
# passcode = "foob pdvi gzab dqyb"

# # List of recipients
# recipients = ["mohitpardhi2003@gmail.com", "mohitpardhi9284@gmail.com"]