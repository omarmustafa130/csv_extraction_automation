# control_panel.py
import os
import sys
import traceback
import subprocess
from flask import Flask, jsonify, request, render_template

app = Flask(__name__)

# Track processes
processes = {
    "daily_service": None,
    "pickup_manifest": None
}

# Store config for each script (NO facility_name)
scripts_config = {
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
    }
}

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/status", methods=["GET"])
def get_status():
    """
    Return running/stopped status + config for each script.
    If ?script=foo, return just that script's status.
    Otherwise return them all.
    """
    script_name = request.args.get("script")
    if script_name:
        if script_name not in processes:
            return jsonify({"error": "Invalid script name"}), 400
        proc = processes[script_name]
        cfg = scripts_config[script_name]
        if proc is None:
            running = False
            status_txt = "Stopped"
        else:
            code = proc.poll()
            if code is None:
                running = True
                status_txt = "Running"
            else:
                running = False
                status_txt = f"Stopped (exit code {code})"

        return jsonify({
            "running": running,
            "status": status_txt,
            "hours": [cfg["start_hour"], cfg["end_hour"]],
            "frequency": cfg["frequency"],
            "folder_id": cfg["folder_id"],
            "username": cfg["username"],
            "password": cfg["password"]  # optional to show
        })
    else:
        # Return status for all scripts
        all_data = {}
        for sname, proc in processes.items():
            cfg = scripts_config[sname]
            if proc is None:
                running = False
                status_txt = "Stopped"
            else:
                code = proc.poll()
                if code is None:
                    running = True
                    status_txt = "Running"
                else:
                    running = False
                    status_txt = f"Stopped (exit code {code})"

            all_data[sname] = {
                "running": running,
                "status": status_txt,
                "hours": [cfg["start_hour"], cfg["end_hour"]],
                "frequency": cfg["frequency"],
                "folder_id": cfg["folder_id"],
                "username": cfg["username"],
                "password": cfg["password"]
            }
        return jsonify(all_data)

@app.route("/control/<script_name>", methods=["POST"])
def control_script(script_name):
    if script_name not in processes:
        return jsonify({"error": "Invalid script name"}), 400

    data = request.json or {}
    action = data.get("action")

    if action == "start":
        if processes[script_name] and processes[script_name].poll() is None:
            return jsonify({"status": f"{script_name} is already running"})

        # Put updated config into environment variables
        env = os.environ.copy()
        cfg = scripts_config[script_name]
        env["START_HOUR"] = str(cfg["start_hour"])
        env["END_HOUR"] = str(cfg["end_hour"])
        env["FREQUENCY"] = str(cfg["frequency"])
        env["FOLDER_ID"] = cfg["folder_id"]
        env["SCRIPT_USERNAME"] = cfg["username"]
        env["SCRIPT_PASSWORD"] = cfg["password"]

        try:
            if script_name == "daily_service":
                proc = subprocess.Popen([sys.executable, "daily_service.py"], env=env)
            elif script_name == "pickup_manifest":
                proc = subprocess.Popen([sys.executable, "pickup_manifest.py"], env=env)
            else:
                return jsonify({"error": "Unknown script"}), 400

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
    """Update run hours, frequency, folder_id, etc. in memory (no facility_name)."""
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

    return jsonify({
        "status": "settings updated",
        "config": cfg
    })

@app.route("/update_credentials/<script_name>", methods=["POST"])
def update_credentials(script_name):
    """Update username/password for a script in memory."""
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

    return jsonify({
        "status": "credentials updated",
        "username": cfg["username"],
        "password": "******"
    })

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
