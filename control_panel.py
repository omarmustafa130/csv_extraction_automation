import os
import sys
import json
import traceback
import subprocess
from flask import Flask, jsonify, request, render_template

app = Flask(__name__)

CONFIG_FILE = "config.json"

# Track processes
processes = {
    "daily_service": None,
    "pickup_manifest": None,
    "weekly_service": None
}

# Default script configurations
default_config = {
    "daily_service": {
        "start_hour": 9,
        "end_hour": 22,
        "frequency": 60,
        "folder_id": "FOLDER_ID",
        "username": "USERNAME",
        "password": "PASSWORD"
    },
    "pickup_manifest": {
        "start_hour": 8,
        "end_hour": 23,
        "frequency": 120,
        "folder_id": "FOLDER_ID",
        "username": "USERNAME",
        "password": "PASSWORD"
    },
    "weekly_service": {
        "schedule_run": 1,  # 1 = Scheduled, 0 = Force Run
        "folder_id": "FOLDER_ID",
        "username": "USERNAME",
        "password": "PASSWORD"
    }
}

# Load settings from JSON file
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return default_config

# Save settings to JSON file
def save_config():
    with open(CONFIG_FILE, "w") as f:
        json.dump(scripts_config, f, indent=4)

# Load initial settings
scripts_config = load_config()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/status", methods=["GET"])
def get_status():
    """Return running/stopped status + config for each script."""
    script_name = request.args.get("script")

    if script_name:
        if script_name not in processes:
            return jsonify({"error": "Invalid script name"}), 400

        proc = processes[script_name]
        cfg = scripts_config[script_name]
        running = proc is not None and proc.poll() is None
        status_txt = "Running" if running else "Stopped"

        return jsonify({
            "running": running,
            "status": status_txt,
            "schedule_run": cfg.get("schedule_run"),
            "hours": [cfg.get("start_hour"), cfg.get("end_hour")],
            "frequency": cfg.get("frequency"),
            "folder_id": cfg["folder_id"],
            "username": cfg["username"],
            "password": "******"
        })

    # Return status for all scripts
    all_data = {
        sname: {
            "running": proc is not None and proc.poll() is None,
            "status": "Running" if proc is not None and proc.poll() is None else "Stopped",
            "schedule_run": cfg.get("schedule_run"),
            "hours": [cfg.get("start_hour"), cfg.get("end_hour")],
            "frequency": cfg.get("frequency"),
            "folder_id": cfg["folder_id"],
            "username": cfg["username"],
            "password": "******"
        }
        for sname, proc in processes.items()
        for cfg in [scripts_config[sname]]
    }

    return jsonify(all_data)


@app.route("/control/<script_name>", methods=["POST"])
def control_script(script_name):
    """Start or stop a script."""
    if script_name not in processes:
        return jsonify({"error": "Invalid script name"}), 400

    data = request.json or {}
    action = data.get("action")

    if action == "start":
        if processes[script_name] and processes[script_name].poll() is None:
            return jsonify({"status": f"{script_name} is already running"})

        # Apply updated settings before starting the script
        env = os.environ.copy()
        cfg = scripts_config[script_name]
        env["FOLDER_ID"] = cfg["folder_id"]
        env["SCRIPT_USERNAME"] = cfg["username"]
        env["SCRIPT_PASSWORD"] = cfg["password"]

        if script_name == "weekly_service":
            env["SCHEDULE_RUN"] = str(cfg["schedule_run"])
        else:
            env["START_HOUR"] = str(cfg["start_hour"])
            env["END_HOUR"] = str(cfg["end_hour"])
            env["FREQUENCY"] = str(cfg["frequency"])

        try:
            proc = subprocess.Popen([sys.executable, f"{script_name}.py"], env=env)
            processes[script_name] = proc
            return jsonify({"status": f"Started {script_name}, PID={proc.pid}"})
        except Exception as e:
            traceback.print_exc()
            return jsonify({"error": str(e)}), 500

    elif action == "stop":
        proc = processes[script_name]
        if not proc or proc.poll() is not None:
            return jsonify({"status": f"{script_name} not running"})
        try:
            proc.terminate()
            proc.wait(timeout=10)
            processes[script_name] = None
            return jsonify({"status": f"Stopped {script_name}"})
        except Exception as e:
            traceback.print_exc()
            return jsonify({"error": str(e)}), 500

    return jsonify({"status": "no change"})


@app.route("/update_settings/<script_name>", methods=["POST"])
def update_settings(script_name):
    """Update run hours, frequency, folder_id, etc. in memory and save them."""
    if script_name not in scripts_config:
        return jsonify({"error": "Invalid script name"}), 400

    cfg = scripts_config[script_name]
    data = request.json or {}

    # Optional fields
    start = data.get("start_hour")
    end = data.get("end_hour")
    freq = data.get("frequency")
    fold_id = data.get("folder_id")

    if start is not None:
        cfg["start_hour"] = int(start)
    if end is not None:
        cfg["end_hour"] = int(end)
    if freq is not None:
        cfg["frequency"] = int(freq)
    if fold_id is not None:
        cfg["folder_id"] = fold_id

    # Save changes to JSON file
    save_config()

    return jsonify({
        "status": "settings updated",
        "config": cfg
    })


@app.route("/update_credentials/<script_name>", methods=["POST"])
def update_credentials(script_name):
    """Update username/password for a script in memory and save them."""
    if script_name not in scripts_config:
        return jsonify({"error": "Invalid script name"}), 400

    cfg = scripts_config[script_name]
    data = request.json or {}

    new_username = data.get("username")
    new_password = data.get("password")

    if new_username is not None:
        cfg["username"] = new_username
    if new_password is not None:
        cfg["password"] = new_password

    # Save changes to JSON file
    save_config()

    return jsonify({
        "status": "credentials updated",
        "username": cfg["username"],
        "password": "******"
    })


@app.route("/update_schedule/<script_name>", methods=["POST"])
def update_schedule(script_name):
    """Update the schedule mode for the weekly service."""
    if script_name != "weekly_service":
        return jsonify({"error": "Only applicable to weekly_service"}), 400

    data = request.json or {}
    schedule_run = data.get("schedule_run")
    if schedule_run is not None:
        scripts_config[script_name]["schedule_run"] = int(schedule_run)

    # Save changes to JSON file
    save_config()

    return jsonify({
        "status": "schedule updated",
        "schedule_run": scripts_config[script_name]["schedule_run"]
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
