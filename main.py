from database import ExcelHandler, Task
from reminder import Reminder
from dotenv import load_dotenv
from functools import wraps
import discord
from discord.ext import commands
from discord import app_commands
import os

# FILE NAME FOR DATABASE and TOKEN Setup
load_dotenv()
TOKEN = str(os.getenv("TOKEN"))
CHANNEL_ID_general = int(os.getenv("CHANNEL_ID_general"))
USER_ID = int(os.getenv("USER_ID"))
GUILD_ID = int(os.getenv("GUILD_ID"))
file_name = "database.xlsx"


# Decorator: Restricting commands to the relevant channel
def only_general_channel():
    def decorator(func):
        @wraps(func)
        async def wrapper(*args, **kwargs):
            first_arg = args[0]

            # For slash commands
            if isinstance(first_arg, discord.Interaction):
                interaction = first_arg
                if interaction.channel.name != "task-manager":
                    await interaction.response.send_message(
                        "This command can only be used in the `task-manager` channel.",
                        ephemeral=True,
                    )
                    return
                return await func(*args, **kwargs)

            # For message commands
            elif hasattr(first_arg, "channel"):
                ctx = first_arg
                if ctx.channel.name != "task-manager":
                    await ctx.send(
                        "This command can only be used in the `task-manager` channel."
                    )
                    return
                return await func(*args, **kwargs)

            return await func(*args, **kwargs)

        return wrapper

    return decorator


# Initialize bot and database
bot = commands.Bot(command_prefix="!", intents=discord.Intents.all())
tree = bot.tree
tasks = ExcelHandler(file_name)
reminder = Reminder(tasks, bot, USER_ID)
tasks.workbook_setup()
guild = discord.Object(id=GUILD_ID)


# Bot event handlers
@bot.event
async def on_ready():
    await tree.sync(guild=guild)
    print(f"Logged in as {bot.user}")

    general_channel = bot.get_channel(CHANNEL_ID_general)
    if general_channel:
        await general_channel.send("✅ BOT is now active!")

    user_dm = await bot.fetch_user(USER_ID)
    if user_dm:
        await user_dm.send("✅ Reminder system is now active!")

    bot.loop.create_task(reminder.refresh_tasks())  # Refresh every 5 minutes
    bot.loop.create_task(
        reminder.schedule_all_reminders()
    )  # Continuously check for new reminders
    print("Reminder system has started!")


# Handles Relevant Commands:
def bot_commands():
    # Adds Task to the Sheet!
    @tree.command(name="create", description="Add a new task", guild=guild)
    @app_commands.describe(
        description="Task description", due_date="Due date", status="Task status"
    )
    async def create_task(
        interaction: discord.Interaction,
        description: str,
        due_date: str,
        status: str = "P",
    ):
        try:
            # Validate description
            if not description.strip():
                await interaction.response.send_message(
                    "❌ Task description cannot be empty.", ephemeral=True
                )
                return

            # Validate status
            status = status.upper()
            if status not in ["P", "C"]:
                await interaction.response.send_message(
                    "❌ Status must be 'P' (Pending) or 'C' (Completed).",
                    ephemeral=True,
                )
                return

            # Use existing format_date method to validate and format the date
            formatted_date = Task.format_date(due_date)
            if formatted_date.startswith("Error formatting date"):
                await interaction.response.send_message(
                    f"❌ {formatted_date}", ephemeral=True
                )
                return

            # Save task
            new_task = Task(description, due_date, status)
            tasks.add_tasks(new_task)
            await interaction.response.send_message(
                f"✅ {description} added successfully!"
            )

        except Exception as e:
            await interaction.response.send_message(
                f"❌ Failed to add task: {str(e)}", ephemeral=True
            )

    # View Tasks In the Sheet!
    @tree.command(name="show", description="View all tasks", guild=guild)
    async def view_tasks(interaction: discord.Interaction):
        all_tasks = tasks.get_tasks()
        if all_tasks:
            formatted_tasks = "\n".join(
                [
                    f"**{i+1}.** **Task:** '{task.description}'\n   **Due Date:** '{task.due_date}'\n   **Status:** {task.status}\n"
                    for i, task in enumerate(all_tasks)
                ]
            )
            # Send the formatted output to Discord
            await interaction.response.send_message(
                f"\n**ACTIVE TASKS**\n\n{formatted_tasks}"
            )
        else:
            await interaction.response.send_message(
                "No tasks found in the Excel sheet."
            )

    # Marking the Tasks Complete
    @tree.command(name="complete", description="Mark a task as complete", guild=guild)
    @app_commands.describe(task_index="The task index to mark as complete")
    async def mark_complete(interaction: discord.Interaction, task_index: int):
        try:
            tasks.complete_task(task_index, "C")
            tasks.delete_completed_task()
            await interaction.response.send_message(
                f"Task {task_index} marked as 'Completed'."
            )
        except Exception as e:
            await interaction.response.send_message(
                f"Error marking task {task_index} as complete: {str(e)}"
            )

    # Deleting the Tasks:
    @tree.command(name="remove", description="Remove a task", guild=guild)
    @app_commands.describe(task_id="The ID of the task to remove")
    async def delete_task(interaction: discord.Interaction, task_id: str):
        if not task_id.isdigit():
            await interaction.response.send_message(
                f"Invalid Task ID: '{task_id}'. Please use numeric values only.",
                ephemeral=True,
            )
            return

        # Call the delete_task method
        try:
            task_deleted = tasks.delete_task(task_id)
            if task_deleted:
                await interaction.response.send_message(
                    f"Task ID '{task_id}' successfully deleted"
                )
            else:
                await interaction.response.send_message(
                    f"Task ID '{task_id}' not found in the sheet."
                )
        except Exception as e:
            await interaction.response.send_message(
                f"An error occurred while deleting the task: {str(e)}"
            )

    # Clears out all the messages -- Tidys the server
    @tree.command(name="clear", description="Clear messages", guild=guild)
    @app_commands.describe(amount="The number of messages to clear")
    @commands.has_permissions(manage_messages=True)
    async def clear(interaction: discord.Interaction, amount: int = 100):
        try:
            # Acknowledge the interaction immediately to avoid the timeout
            await interaction.response.defer()

            # Purge the messages
            deleted = await interaction.channel.purge(limit=amount)

            # Send the confirmation message after the purge is complete
            await interaction.followup.send(
                f"✅ Cleared {len(deleted)} messages!", delete_after=5
            )
        except Exception as e:
            # Log only if there's an actual error
            print(f"Error clearing messages: {e}")
            await interaction.followup.send(
                "❌ Failed to clear messages. Make sure I have the right permissions.",
                ephemeral=True,
            )


def main():
    bot_commands()
    bot.run(TOKEN)


if __name__ == "__main__":
    main()
