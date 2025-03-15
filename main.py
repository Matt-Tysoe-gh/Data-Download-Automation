import win32com.client
 import re
 import webbrowser
 import os
 import sys
 import threading
 import pandas as pd
 from datetime import datetime
 from typing import Tuple, List, Dict
 import warnings
 import logging
 import constants
 
 logging.basicConfig(
     level=logging.DEBUG,
     format="%(asctime)s %(levelname)s: %(message)s",
     datefmt="%d-%m-%Y %H:%M:%S"
 )
 
 def initialise_win32com(outlook_folder_index: int) -> Tuple[object, object]:
     logging.info("initialising win32com Outlook application")
     try:
         outlook = win32com.client.Dispatch("Outlook.Application")
         namespace = outlook.GetNamespace("MAPI")
         inbox = namespace.GetDefaultFolder(outlook_folder_index)
         logging.info("Outlook initialised successfully")
         return inbox, outlook
     except Exception:
         logging.exception("Failed to initialise win32com Outlook application")
         raise
 
 
 def get_and_download_emails(inbox: object, filter_list: List[str], email_subjects: Dict[str, str], url_pattern: str, file_prefix: str, download_folder: str) -> Tuple[Dict[str, Dict[str, str]], str]:
 
     logging.info("Fetching and downloading emails")
     mapping = email_subjects
     dictionary = {}
     for subject in filter_list:
         logging.debug(f"Processing subject: {subject}")
         email_filter = f"[Subject] = '{subject}'"
         try:
             email_messages = inbox.Items.Restrict(email_filter)
             latest_email = email_messages.GetLast()
             logging.debug(f"Got latest email for subject: {subject}")
         except Exception:
             logging.exception(f"Error finding subject: {subject}")
             continue
 
         url_match = re.search(rf"{url_pattern}[^\s<>\"']+", latest_email.body)
         if not url_match:
             logging.error(f"Error extracting URL: {url_match}")
             continue
         url = url_match.group(0)
         logging.debug(f"Extracted URL: {url}")
 
         webbrowser.open(url)
         logging.info("Opening URL in browser")
 
         _, _, partition = url.partition(file_prefix)
         downloads = os.path.join(os.path.expanduser("~"), download_folder)
         file_name = file_prefix + partition
         file_path = os.path.join(downloads, file_name)
         logging.debug(f"Constructed file path: {file_path}")
 
         if subject in mapping:
             key = mapping[subject]
             dictionary[key] = {
                 "urls": url,
                 "file_paths": file_path
             }
             logging.info(f"Mapping dict for {subject} and {key}")
     return dictionary, downloads
 
 
 def wait_for_file(file_path: str, max_attempts: int = 1200) -> bool:
     logging.info(f"Waiting for file: {file_path}")
     event = threading.Event()
     for attempt in range(max_attempts):
         if os.path.exists(file_path):
             logging.info(f"File found after {attempt / 10:.1f}s")
             return True
         else:
             logging.debug(f"Waiting for file... {attempt / 10:.1f}s")
             sys.stdout.flush()
             event.wait(0.1)
     logging.error("File not found within maximum allowed time... Try again...")
     return False
 
 
 def reviews_transformations(weeknum: int, dictionary: Dict[str, Dict[str, str]], downloads: str) -> str:
     logging.info("Starting reviews file transformations")
     try:
         df_review = pd.read_excel(dictionary["review"]["file_paths"])
         logging.info("Reviews file read successfully")
     except(FileNotFoundError, PermissionError):
         logging.exception("Error reading reviews file")
         raise
 
     try:
         df_review.columns = df_review.iloc[6]
         df_review = df_review.iloc[8:]
         df_review['Review Submission Date'] = pd.to_datetime(df_review['Review Submission Date'])
         df_review.insert(2, 'Purchase Date', df_review['Review Submission Date'] - pd.Timedelta(days=1))
         df_review['Review Text'] = df_review['Review Text'].apply(lambda x: f'"{x}"')
         df_review.insert(2, 'email', 'email@email.com')
         df_review.insert(6, 'blank 1', None)
         df_review.insert(7, 'blank 2', None)
         logging.debug("Transformations applied to reviews file")
     except Exception:
         logging.exception("Could not apply transformations to reviews file")
         raise
 
     review_file_path = os.path.join(downloads, f'filename_reviews{weeknum}.xlsx')
 
     try:
         df_review.to_excel(review_file_path, index=False)
         logging.info(f"Reviews data written to {review_file_path}")
     except(FileNotFoundError, PermissionError):
         logging.exception("Error writing reviews Excel file")
         raise
     return review_file_path
 
 
 def catalog_transformations(weeknum: int, dictionary: Dict[str, Dict[str, str]], downloads: str) -> str:
     logging.info("Starting catalog file transformations")
     try:
         df_catalog = pd.read_excel(dictionary["catalog"]["file_paths"])
     except(FileNotFoundError, PermissionError):
         logging.exception("Error reading catalog file")
         raise
 
     try:
         df_catalog.columns = df_catalog.iloc[5]
         df_catalog = df_catalog.iloc[6:]
         df_catalog['EAN'] = df_catalog['EAN'].str.split(',').str[0]
     except Exception:
         logging.exception("Could not apply transformations to catalog file")
         raise
 
     catalog_file_path = os.path.join(downloads, f'filename_catalog{weeknum}.xlsx')
 
     try:
         df_catalog.to_excel(catalog_file_path, index=False)
         logging.info(f"Catalog data written to {catalog_file_path}")
     except(FileNotFoundError, PermissionError):
         logging.exception("Error writing catalog Excel file")
     return catalog_file_path
 
 
 def send_email(outlook: object, pd_file_paths: List[str], weeknum: int, year: int, recipient_email: str) -> bool:
     if len(pd_file_paths) != 2:
         logging.warning("Insufficient attachments available. Email will not be sent")
         return False
 
     logging.info("Preparing email to be sent")
     try:
         mail = outlook.CreateItem(0)
         mail.To = recipient_email
         mail.Subject = f'Review Data Export Week {weeknum} Year {year}'
         mail.Body = f'''
         Hi,
 
         Email body goes here! it is currenly {year} week {weeknum}
 
         Thanks,
         user        
         '''
 
         for attachment in pd_file_paths:
             logging.debug(f"Attaching file: {attachment}")
             mail.Attachments.Add(attachment)
 
         mail.Send()
         logging.info("Email sent successfully")
         return True
     except Exception:
         logging.exception("Error occurred while sendin email.")
         return False
 
 
 if __name__ == "__main__":
     """
     main guard for code/function organisation, also initialises empty lists and filter lists
     """
     try:
         warnings.filterwarnings("ignore", message="Workbook contains no default style, apply openpyxl's default")
         logging.info("Script started")
 
         filter_list = ["", ""]
         pd_file_paths = []
 
         weeknum = f"{datetime.now().isocalendar()[1]-1:02d}"
         year = f"{datetime.now().isocalendar()[0]}"
         logging.debug(f"Week number: {weeknum}, Year: {year}")
         
         inbox, outlook = initialise_win32com(constants.OUTLOOK_FOLDER_INDEX)
 
         dictionary, downloads = get_and_download_emails(inbox, filter_list, constants.EMAIL_SUBJECTS, constants.URL_PATTERN, constants.FILE_PREFIX, constants.DOWNLOAD_FOLDER)
 
         if wait_for_file(dictionary["review"]["file_paths"]):
             review_file_path = reviews_transformations(weeknum, dictionary, downloads)
             pd_file_paths.append(review_file_path)
         else:
             raise FileNotFoundError("File not found")
 
         if wait_for_file(dictionary["catalog"]["file_paths"]):
             catalog_file_path = catalog_transformations(weeknum, dictionary, downloads)
             pd_file_paths.append(catalog_file_path)
         else:
             raise FileNotFoundError("File not found")
 
         sent = send_email(outlook, pd_file_paths, weeknum, year, constants.RECIPIENT_EMAIL)
         if sent:
             logging.info("Script has completed successfully")
         else:
             logging.error("Script compeleted, but the email was not sent due to missing attachments or an error")
     except Exception:
         logging.exception("An unhandled error has occurred in the main script")
