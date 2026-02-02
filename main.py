import time
import shutil
import os
import logging
from datetime import datetime
import getpass
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from docx2pdf import convert

#init log file
logging.basicConfig(
    filename= "file_events.log",
    level=logging.INFO,
    format= "%(message)s"
)

def log_event(action, path):
    #write evnets on log file
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    user = getpass.getuser()
    filename = os.path.basename(path)

    logging.info(f"{timestamp} | {user} | {action} | {filename}")

def wait_until_stable(path, timeout=5):
    #Wait until file stops changing size.
    last_size = -1
    start = time.time()

    while time.time() - start < timeout:
        try:
            size = os.path.getsize(path)
        except FileNotFoundError:
            return False  # file disappeared

        if size == last_size:
            return True  # stable

        last_size = size
        time.sleep(0.5)

    return False

def convert_file(input_path, output_path):
    #conert word or excel file to pdf file
    convert(input_path, output_path)

class MyHandler(FileSystemEventHandler):
    #event handler when file changes 
    def on_created(self, event):

        # If a temporary file triggers the event, ignore it
        if os.path.basename(event.src_path).startswith("~$"):
            return
        path = event.src_path
        destination_folder = "loactionA"
        outputfile = os.path.splitext(os.path.basename(event.src_path))[0] + ".pdf"        
        ext = os.path.splitext(path)[1].lower()

        if event.is_directory:
            return #ignore folders events
        
        if not wait_until_stable(path):
            #PDF file is not generate yet
            print("File never stabilized: please try again", path)
            return

        #only converted  to pdf file if file ext is .docx
        if ext == ".docx":
            convert_file(event.src_path,outputfile)
            print("file converted to PDF file", event.src_path)
        else: 
            print("Not a Docx file:", event.src_path)
            return
        
        #if destination_folder does not exit it will create new destination_folder
        os.makedirs(destination_folder, exist_ok=True)
        destionation = os.path.join(destination_folder, outputfile)
        shutil.move(outputfile, destionation)

        print(f"Copied to {os.path.dirname(destionation)}", destionation)
        log_event(f"Copied to {os.path.dirname(destionation)}", destionation)

    def on_deleted(self, event):
        if os.path.basename(event.src_path).startswith("~$"):
            return
        print(f"Deleted: {event.src_path}")
        log_event("Deleted", event.src_path)

#Main class
class App:
    def __init__(self):
        #need folder to send docx file eg. Test
        path_to_march = "./Test"
        event_handler = MyHandler()
        self.observer = Observer()
        self.observer.schedule(event_handler, path_to_march, recursive=True)

        self.observer.start()
        print(f"Watching for change in: {path_to_march}")

    def run(self):
        #runing program
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            #ctrl + c to stop the program
            self.observer.stop()
        
        #cleaning up observer
        self.observer.join()

if __name__ == "__main__":
    app = App()

    app.run()
