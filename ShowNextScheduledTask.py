import win32com.client
import datetime
import tkinter as tk




def convert_com_time(com_time):
    """Convert COM time to Python datetime"""
    if com_time:
        return datetime.datetime(com_time.year, com_time.month, com_time.day, 
                                 com_time.hour, com_time.minute, com_time.second)
    return None

def get_today_tasks():
    # Connect to the Task Scheduler service
    scheduler = win32com.client.Dispatch('Schedule.Service')
    scheduler.Connect()
    
    # Get the root folder
    root_folder = scheduler.GetFolder('\\RPA')

    # Get the current date
    today = datetime.datetime.now().date()

    # Get all tasks in the root folder
    tasks = root_folder.GetTasks(0)
    today_tasks = []

    for task in tasks:
        next_run_time = convert_com_time(task.NextRunTime)
        task_name = task.Name
        task_state = task.State

        if next_run_time and today == next_run_time.date() and task_state==3:  # Scheduled to run today
            today_tasks.append((task_name, next_run_time, "Scheduled"))

    return today_tasks

class TextPrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Task Scheduler Timer")

        # Create a Label widget to display the current text
        self.label = tk.Label(root, text="", font=("Helvetica", 16))
        self.label.pack(padx=10, pady=10)

        # Start the updating loop
        self.update_text()

    def update_text(self):
        tasks = get_today_tasks()

        now = datetime.datetime.now()
        
        #find earlier schedule
        schedules = []
        for task in tasks:
            schedules.append(task[1])
        earliest_datetime = min((dt, idx) for idx, dt in enumerate(schedules))
        left = (earliest_datetime[0] - now).seconds
        left_hours = left // 3600
        left_minutes = left % 3600 // 60
        left_seconds = left % 3600 % 60

        # Get the current time and update the label
        announce = tasks[earliest_datetime[1]][0] + "：残り " + str(left_hours) + "時" + str(left_minutes) + "分" + str(left_seconds) + "秒"
        self.label.config(text=announce)

        # Schedule the next update after 1000 milliseconds (1 second)
        self.root.after(1000, self.update_text)


if __name__ == "__main__":
    root = tk.Tk()
    app = TextPrinterApp(root)
    root.mainloop()

