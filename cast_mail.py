import win32com.client
import os, time
from dotenv import  load_dotenv
from pathlib import Path
import click

from utils import load_client_info, make_filename, move_to_client_folder, cleanup_folder

load_dotenv()
# Your Outlook email and password
email = os.getenv("OUTLOOK_USER")
password = os.getenv("OUTLOOK_PASSWORD")

onedrive_base = Path(r"C:\Users\Sullivan Boecker\OneDrive - Cast Accounting\Candice Hayes-McInnis's files - Clients")
staged_path = Path(r"C:\Users\Sullivan Boecker\scripts\report_staging")

client_name = "Bocoup"
client_data = load_client_info(client_name)
addresses = [e for e in client_data["info"]["email"]]

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)

mail.To = "; ".join(addresses)
mail.Subject = "Test Email"
mail.Body = """
Hello!
<br><br>
Attached is your P&L Detail for last month. Let me know if you have any questions!
"""

sig = """
<br><br>
<table cellpadding="0" cellspacing="0" border="0" style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
  <tr>
    <td>
      <strong>Sully Boecker</strong><br>
      Accounting &amp; Tax Associate<br>
      <a href="mailto:sully@castaccounting.co">sully@castaccounting.co</a><br><br>

      Need to send documents?? ~&gt;
      <a href="https://castaccounting.securefilepro.com">
        castaccounting.securefilepro.com
      </a><br>

      Need to schedule some time?? ~&gt;
      <a href="https://www.calendly.com/castaccounting">
        www.calendly.com/castaccounting
      </a>
    </td>
  </tr>
</table>
<img src="C:\\Users\\Sullivan Boecker\\scripts\\email_things\\Picture1.png">
"""
mail.HTMLBody = mail.Body + sig

new_file= ""
for file in staged_path.iterdir():
    new_file = make_filename(file, client_name)
    os.rename(file, new_file)

mail.Attachments.Add(str(new_file))
# mail.Categories = "To-Do"
mail.Display()

time.sleep(1)

cleanup_folder(Path(client_data["info"]["path"]))
move_to_client_folder(client_data["info"]["path"], new_file)
# mail.Send()
