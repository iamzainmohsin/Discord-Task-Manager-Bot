# Discord Task Manager Bot

A simple Discord bot built with Python to manage personal or group tasks using an Excel file for storage. Users can create, view, complete, and delete tasks through slash commands. The bot also sends reminder messages based on due dates.

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Commands](#commands)
- [Hosting Options](#hosting-options)
- [Project Structure](#project-structure)
- [License](#license)

---

## Overview

This bot helps manage daily tasks on Discord using slash commands. Tasks are saved in an Excel sheet and reminders are sent based on their due dates. The code is organized into files for commands, task storage, and reminders.

---

## Features

- **Add, list, complete, and delete tasks**
- **Save tasks in an Excel sheet using `openpyxl`**
- **Automatic Excel file creation and setup**
- **Send reminders based on how close the due date is**
- **Use of slash commands with `discord.py`**

---

## Installation

1. **Clone the repository:**

```bash
git clone https://github.com/<your-username>/<your-repo-name>.git
cd <your-repo-name>
```

2. **Create a virtual environment:**

```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. **Install dependencies:**

```bash
pip install -r requirements.txt
```

If you don’t have a `requirements.txt`, install the main libraries manually:

```bash
pip install discord.py python-dotenv openpyxl
```

---

## Configuration

1. **Create a `.env` file in the root directory:**

```env
TOKEN=your_discord_bot_token
CHANNEL_ID_general=your_channel_id_for_messages
USER_ID=your_user_id_for_direct_messages
GUILD_ID=your_guild_id
```

2. **Excel Database:**

The bot uses a file called `database.xlsx` to store tasks. It will be created automatically if it doesn’t exist.

---

## Usage

Run the bot using:

```bash
python main.py
```

Use the following slash commands in Discord:

### ➕ Add a Task

```text
/create description:"Finish assignment" due_date:"05-April-2025" status:"P"
```

### 📋 View Tasks

```text
/show
```

### ✅ Complete a Task

```text
/complete task_index:2
```

### ❌ Delete a Task

```text
/remove task_id:"3"
```

### 🧹 Clear Messages

```text
/clear amount:50
```

---

## Commands

| Command     | Description                         | Example                                        |
|-------------|-------------------------------------|------------------------------------------------|
| /create     | Add a new task                      | /create description:"Task", due_date:"...", status:"P" |
| /show       | View all tasks                      | /show                                          |
| /complete   | Mark a task as complete             | /complete task_index:2                         |
| /remove     | Delete a task                       | /remove task_id:"3"                            |
| /clear      | Delete recent messages from channel | /clear amount:50                               |

---

## Hosting Options

You can host this bot on different platforms:

### Option 1: Local Machine (for development)

- Run `main.py` directly from your computer.
- Keep the terminal open for the bot to stay active.

### Option 2: Oracle Cloud (free virtual machine)

- Create a free Ubuntu VM using Oracle Cloud Free Tier.
- Install Python and clone your repository onto the VM.
- Use tools like `screen` or `tmux` to keep the bot running in the background.

### Option 3: Other VPS Providers

- You can also host the bot on platforms like DigitalOcean, Linode, or AWS EC2.
- Steps are similar to Oracle Cloud: set up Python, clone your repo, and run `main.py`.

---

## Project Structure

```
.
├── main.py          # Handles bot setup and commands
├── database.py      # Task storage and Excel logic
├── reminder.py      # Reminder scheduling
├── database.xlsx    # Excel file for tasks (auto-generated)
├── .env             # Environment config (not shared)
└── README.md        # Project guide
```

---

## License

This project is open-source. You may use and modify it as needed.

---

Have fun using your Discord Task Manager Bot!

