import asyncio
from datetime import datetime
import time

class Reminder:
    def __init__(self, excel_handler, bot, USER_ID):
        self.scheduled_tasks = set()
        self.excel_handler = excel_handler
        self.bot = bot
        self.USER_ID = USER_ID
        self.running_frequencies = set()  # Keeps track of active reminder loops
        self.last_sent = {}

    async def refresh_tasks(self):
        while True:
            await asyncio.sleep(300)  # Wait for 5 minutes

            tasks = self.db.get_tasks()  # Fetch tasks from your database (or Excel)
            for task in tasks:
                # Check if this task already has a running reminder
                if task.index not in self.running_frequencies:
                    print(f"Scheduling reminder for task: {task.description}")
                    self.schedule_reminder(task)
                else:
                    print(f"Reminder already running for task: {task.description}")

    # Deletes up overdue tasks after 3 days
    def clean_overdue_tasks(self):
        print("Cleaning overdue tasks...")
        current_tasks = self.fetch_pending_tasks()
        today = datetime.now().date()
        removed_tasks = []

        print(f"Total pending tasks: {len(current_tasks)}")

        for task in current_tasks:
            try:
                due_date = datetime.strptime(task.due_date, "%d-%B-%Y").date()
            except ValueError:
                print(
                    f"Skipping task '{task.description}': Invalid due date format -> {task.due_date}"
                )
                continue

            days_overdue = (today - due_date).days
            print(
                f"Checking task: {task.description}, Due Date: {due_date}, Overdue by: {days_overdue} days"
            )

            if days_overdue > 3:
                removed_tasks.append(task)
                self.excel_handler.delete_task(
                    str(task.index)
                )  # Ensure task ID is passed as string

        print(f"{len(removed_tasks)} overdue tasks removed.")

    # Loads the Pending tasks from the sheet
    def fetch_pending_tasks(self):
        # Directly call the method from the ExcelHandler class
        pending_tasks = self.excel_handler.get_tasks()
        print(f"Pending Tasks Fetched: {len(pending_tasks)} tasks found.")
        return pending_tasks

    def days_until_due(self, due_date):
        today = datetime.now().date()

        if isinstance(due_date, str):
            due_date = datetime.strptime(due_date, "%d-%B-%Y").date()

        days_left = (due_date - today).days
        return days_left

    # Sets Reminder Frequency
    def reminder_frequency(self, days_left):
        if days_left < 0:
            return None  # Overdue tasks
        elif days_left == 0:
            return "Hourly Reminder"  # Tasks due today
        elif days_left <= 3:
            return "Every 4 Hours"  # Tasks due within 3 days
        elif days_left <= 7:
            return "Daily Reminder"  # Tasks due later in the week
        else:
            return "Weekly Reminder"  # Far future tasks

    # Schedules the reminders
    async def schedule_reminder(self, task):
        task_description = task.description
        due_date = task.due_date
        days_left = self.days_until_due(due_date)
        frequency = self.reminder_frequency(days_left)

        # Intervals for different frequencies (in seconds)
        intervals = {
            "Hourly Reminder": 3600,  # 3600
            "Every 4 Hours": 14400,  # 14400
            "Daily Reminder": 86400,  # 86400
            "Weekly Reminder": 604800,  # 604800
        }
        interval = intervals.get(frequency, None)

        if not interval:
            print(f"Unknown frequency '{frequency}' for task: {task_description}")
            return

        # Ensure grouped task tracking
        if frequency not in self.last_sent:
            self.last_sent[frequency] = 0  # Initialize last sent time

        if frequency not in self.running_frequencies:
            self.running_frequencies.add(frequency)
        else:
            print(
                f"Reminder loop for '{frequency}' already running. Skipping duplicate."
            )
            return

        while True:
            current_tasks = self.fetch_pending_tasks()

            grouped_tasks = [
                t
                for t in current_tasks
                if self.days_until_due(t.due_date) >= 0
                and self.reminder_frequency(self.days_until_due(t.due_date))
                == frequency
            ]

            if not grouped_tasks:
                print(
                    f"No active tasks to send for frequency '{frequency}'. Ending loop."
                )
                self.running_frequencies.discard(frequency)
                break

            current_time = time.time()
            last_sent_time = self.last_sent[frequency]

            if current_time - last_sent_time >= interval:
                # Only send unique reminders once per loop tick
                unique_task_set = set()
                message = f"\n📌 **Task Reminders:**\n"
                for t in grouped_tasks:
                    if t.description not in unique_task_set:
                        unique_task_set.add(t.description)
                        message += (
                            f"• **{t.description}**\n   Due Date: {t.due_date}\n\n"
                        )

                try:
                    user = await self.bot.fetch_user(self.USER_ID)
                    await user.send(message)
                    print(
                        f"Grouped reminder sent for frequency '{frequency}' at {time.ctime(current_time)}."
                    )
                except Exception as e:
                    print(f"Failed to send reminder: {e}")

                self.last_sent[frequency] = current_time
            else:
                remaining_time = interval - (current_time - last_sent_time)
                print(
                    f"Tasks for frequency '{frequency}' waiting {remaining_time:.1f} seconds before next reminder."
                )

            await asyncio.sleep(interval)

    # Looks for newly added tasks within the sheet
    async def check_and_update_tasks(self):
        while True:
            new_tasks = self.fetch_pending_tasks()

            for i, task in enumerate(new_tasks):  # Generate task_id dynamically
                task_id = i + 1  # Task ID based on enumerate()

                if task_id not in self.scheduled_tasks:
                    self.scheduled_tasks.add(task_id)
                    # Introduce staggered delay before scheduling
                    await asyncio.sleep(i * 3)
                    asyncio.create_task(self.schedule_reminder(task))

            await asyncio.sleep(300)  # Check for new tasks every 5 minutes

    # Runs the clean_over_due() function
    async def daily_cleanup(self):
        while True:
            self.clean_overdue_tasks()
            await asyncio.sleep(86400)  # 24 hours

    # Runs Everything:
    async def schedule_all_reminders(self):
        asyncio.create_task(self.daily_cleanup())
        await self.check_and_update_tasks()
