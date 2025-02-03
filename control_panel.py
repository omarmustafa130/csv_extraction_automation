# control_panel.py
import os
import sys
import traceback
import subprocess
from flask import Flask, jsonify, request, render_template

app = Flask(__name__)

# Track the Popen objects (the worker scripts)
processes = {
    "daily_service": None,
    "pickup_manifest": None
}

# Store each script's configuration so you can display & update it
scripts_config = {
    "daily_service": {
        "start_hour": 9,
        "end_hour": 22,
        "frequency": 60,
        "facility_name": "DailyService",
        "folder_id": "FOLDER_ID"
    },
    "pickup_manifest": {
        "start_hour": 8,
        "end_hour": 23,
        "frequency": 120,
        "facility_name": "PickUpManifest",
        "folder_id": "FOLDER_ID"
    }
}


@app.route("/")
def index():
    # Renders the templates/index.html file
    return render_template("index.html")


@app.route("/status", methods=["GET"])
def get_status():
    """
    Return the running/stopped status + config for each script.
    If ?script=foo is in the query, return just that one script's data.
    Otherwise return all.
    """
    script_name = request.args.get("script")
    if script_name:
        if script_name not in processes:
            return jsonify({"error": "Invalid script name"}), 400
        proc = processes[script_name]
        config = scripts_config[script_name]
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
            "hours": [config["start_hour"], config["end_hour"]],
            "frequency": config["frequency"],
            "facility_name": config["facility_name"],
            "folder_id": config["folder_id"]
        })

    else:
        # Return status for all scripts
        all_data = {}
        for sname, proc in processes.items():
            config = scripts_config[sname]
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
                "hours": [config["start_hour"], config["end_hour"]],
                "frequency": config["frequency"],
                "facility_name": config["facility_name"],
                "folder_id": config["folder_id"]
            }
        return jsonify(all_data)


@app.route("/control/<script_name>", methods=["POST"])
def control_script(script_name):
    """Start or stop the given script as a separate process."""
    if script_name not in processes:
        return jsonify({"error": "Invalid script name"}), 400

    data = request.json or {}
    action = data.get("action")

    if action == "start":
        # If it's already running, do nothing
        if processes[script_name] and processes[script_name].poll() is None:
            return jsonify({"status": f"{script_name} already running"})

        # Build environment variables so the worker script can read the config
        env = os.environ.copy()
        cfg = scripts_config[script_name]
        env["START_HOUR"] = str(cfg["start_hour"])
        env["END_HOUR"] = str(cfg["end_hour"])
        env["FREQUENCY"] = str(cfg["frequency"])
        env["FACILITY_NAME"] = cfg["facility_name"]
        env["FOLDER_ID"] = cfg["folder_id"]

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
        # Stop the script if running
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
    """Update the config for a script in memory (doesn't restart automatically)."""
    if script_name not in scripts_config:
        return jsonify({"error": "Invalid script name"}), 400

    cfg = scripts_config[script_name]
    data = request.json or {}

    # Optional fields
    start = data.get("start_hour")
    end = data.get("end_hour")
    freq = data.get("frequency")
    fname = data.get("facility_name")
    fold_id = data.get("folder_id")

    if start is not None:
        cfg["start_hour"] = int(start)
    if end is not None:
        cfg["end_hour"] = int(end)
    if freq is not None:
        cfg["frequency"] = int(freq)
    if fname is not None:
        cfg["facility_name"] = fname
    if fold_id is not None:
        cfg["folder_id"] = fold_id

    return jsonify({
        "status": "settings updated",
        "config": cfg
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
