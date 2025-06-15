import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent))
from flask import Flask, render_template, request, jsonify, Response
import os
import subprocess
import platform
import glob
import shutil
import re
from datetime import datetime, timedelta
import importlib.util
import threading
import logging
from typing import Dict, List, Optional

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Add the module path to system path
module_path = "C:/VoicePilot/Version2/simple-whisper-transcription/src"
if module_path not in sys.path:
    sys.path.append(module_path)

# Import LiveTranscriber
spec = importlib.util.spec_from_file_location(
    "LiveTranscriber", 
    "C:/VoicePilot/Version2/simple-whisper-transcription/src/LiveTranscriber.py"
)
live_transcriber_module = importlib.util.module_from_spec(spec)
sys.modules["LiveTranscriber"] = live_transcriber_module
spec.loader.exec_module(live_transcriber_module)
from LiveTranscriber import LiveTranscriber

app = Flask(__name__)

# Configuration
app.config['SECRET_KEY'] = 'voicepilot-secure-key-2024'
app.config['UPLOAD_FOLDER'] = 'uploads'

# Initialize the transcriber
transcriber = LiveTranscriber()
transcription_thread = None
transcription_active = False

# Create necessary directories
os.makedirs('static/css', exist_ok=True)
os.makedirs('static/js', exist_ok=True)
os.makedirs('templates', exist_ok=True)
os.makedirs('uploads', exist_ok=True)

# Platform detection
SYSTEM = platform.system().lower()

def open_file_or_folder(path: str) -> str:
    """Open a file or folder using platform-specific commands."""
    try:
        path = os.path.abspath(os.path.expanduser(path))
        if not os.path.exists(path):
            return f"Path does not exist: {path}"
        
        if SYSTEM == "windows":
            os.startfile(path)
        elif SYSTEM == "darwin":  # macOS
            subprocess.run(["open", path], check=True)
        elif SYSTEM == "linux":
            subprocess.run(["xdg-open", path], check=True)
        else:
            return "Unsupported operating system"
        return f"Opened: {path}"
    except Exception as e:
        logger.error(f"Error opening path {path}: {e}")
        return f"Error opening path: {e}"

def open_application(app_name: str) -> str:
    """Open an application based on the platform."""
    try:
        app_name = app_name.strip().lower()
        # Map common application names to platform-specific commands
        app_mapping = {
            "file explorer": "explorer" if SYSTEM == "windows" else "finder" if SYSTEM == "darwin" else "nautilus",
            "explorer": "explorer" if SYSTEM == "windows" else "finder" if SYSTEM == "darwin" else "nautilus"
        }
        app_command = app_mapping.get(app_name, app_name)
        
        if SYSTEM == "windows":
            subprocess.run(["start", "", app_command], shell=True, check=True)
        elif SYSTEM == "darwin":
            subprocess.run(["open", "-a", app_command], check=True)
        elif SYSTEM == "linux":
            subprocess.run([app_command], check=True)
        return f"Opened application: {app_name}"
    except Exception as e:
        logger.error(f"Error opening application {app_name}: {e}")
        return f"Error opening application: {e}"

def find_files(pattern: str, directory: str = None, time_filter: str = None) -> List[str]:
    """Find files matching a pattern in the specified directory, optionally filtered by time."""
    try:
        directory = os.path.expanduser(directory or "~/")
        files = glob.glob(os.path.join(directory, pattern), recursive=True)
        
        if time_filter == "last week":
            one_week_ago = datetime.now() - timedelta(days=7)
            files = [
                f for f in files
                if os.path.getmtime(f) >= one_week_ago.timestamp()
            ]
        
        return files[:10]  # Limit to 10 results for brevity
    except Exception as e:
        logger.error(f"Error finding files: {e}")
        return []

def organize_files(file_type: str, destination: str) -> str:
    """Move files of a specific type to a destination folder."""
    try:
        destination = os.path.expanduser(destination)
        if not os.path.exists(destination):
            os.makedirs(destination)
        
        patterns = {
            "images": ["*.jpg", "*.jpeg", "*.png", "*.gif"],
            "pdfs": ["*.pdf"],
            "documents": ["*.doc", "*.docx", "*.txt"]
        }
        
        files_moved = 0
        for pattern in patterns.get(file_type.lower(), [f"*.{file_type}"]):
            for file in find_files(pattern):
                dest_path = os.path.join(destination, os.path.basename(file))
                if not os.path.exists(dest_path):
                    shutil.move(file, dest_path)
                    files_moved += 1
        
        return f"Moved {files_moved} {file_type} to {destination}"
    except Exception as e:
        logger.error(f"Error organizing files: {e}")
        return f"Error organizing files: {e}"

def sort_files(directory: str, by: str = "size") -> str:
    """Sort files in a directory by size or other criteria."""
    try:
        directory = os.path.expanduser(directory)
        files = [(f, os.path.getsize(f)) for f in glob.glob(os.path.join(directory, "*"))]
        files.sort(key=lambda x: x[1], reverse=True if by.lower() == "size" else False)
        
        # Create a sorted directory
        sorted_dir = os.path.join(directory, "sorted")
        if not os.path.exists(sorted_dir):
            os.makedirs(sorted_dir)
        
        for i, (file_path, _) in enumerate(files):
            dest_path = os.path.join(sorted_dir, f"{i+1}_{os.path.basename(file_path)}")
            if not os.path.exists(dest_path):
                shutil.move(file_path, dest_path)
        
        return f"Sorted {len(files)} files in {directory} by {by} to {sorted_dir}"
    except Exception as e:
        logger.error(f"Error sorting files: {e}")
        return f"Error sorting files: {e}"

def parse_command(command: str) -> Dict[str, any]:
    """Parse the voice command using regex."""
    command = command.lower().strip()
    
    patterns = {
        "open_file": r"open\s+(?:file\s+)?(.+\.(?:txt|pdf|docx?|jpg|png|gif))$",
        "open_folder": r"open\s+(?:folder\s+)?([^\.]+)$",
        "open_app": r"open\s+(?:application\s+|app\s+|file\s+)?(?:explorer|notepad|firefox|chrome|safari|finder|nautilus)(?:\s+.*)?$",
        "find_files": r"(?:find|search)\s+(?:all\s+)?(.+?)(?:\s+from\s+(.+))?(?:\s+from\s+last\s+week)?$",
        "organize": r"move\s+(?:all\s+)?(.+?)\s+to\s+(.+)",
        "sort": r"sort\s+(?:files\s+|documents\s+)?(?:in\s+)?(.+?)\s+by\s+(.+)",
        "recent_downloads": r"(?:open\s+)?recent\s+downloads(?:\s+.*)?$"
    }
    
    for action, pattern in patterns.items():
        match = re.match(pattern, command)
        if match:
            groups = match.groups()
            if action == "find_files":
                return {
                    "action": action,
                    "params": {
                        "file_type": groups[0],
                        "directory": groups[1] if groups[1] else None,
                        "time_filter": "last week" if "last week" in command else None
                    }
                }
            return {"action": action, "params": groups}
    
    # Handle ambiguous or incomplete search commands
    if "search" in command or "find" in command:
        return {
            "action": "find_files",
            "params": {
                "file_type": "files",  # Default to generic file search
                "directory": None,
                "time_filter": None
            }
        }
    
    return {"action": "unknown", "params": ()}

def process_voice_command(command: str) -> Dict:
    """Process the parsed command and execute the appropriate action."""
    parsed = parse_command(command)
    action = parsed["action"]
    params = parsed["params"]
    
    logger.info(f"Processing command: {command}, Action: {action}, Params: {params}")
    
    response = {
        'status': 'success',
        'message': '',
        'action': action,
        'timestamp': datetime.now().isoformat()
    }
    
    if action == "open_file" or action == "open_folder":
        response['message'] = open_file_or_folder(params[0])
    elif action == "open_app":
        response['message'] = open_application(params[0])
    elif action == "find_files":
        file_type = params["file_type"]
        directory = params["directory"]
        time_filter = params["time_filter"]
        # Handle generic "files" search
        pattern = "*.pdf" if file_type == "files" else f"*.{file_type}" if not file_type.startswith("*") else file_type
        files = find_files(pattern, directory, time_filter)
        response['message'] = f"Found {len(files)} files: {', '.join(files)}" if files else f"No {file_type} found"
    elif action == "organize":
        file_type, destination = params
        response['message'] = organize_files(file_type, destination)
    elif action == "sort":
        directory, criteria = params
        response['message'] = sort_files(directory, criteria)
    elif action == "recent_downloads":
        downloads = os.path.expanduser("~/Downloads")
        response['message'] = open_file_or_folder(downloads)
    else:
        response['status'] = 'error'
        response['message'] = f"Unknown command: {command}. Try 'open file explorer', 'find PDFs', or 'open recent downloads'."
    
    return response

@app.route('/')
def home():
    """Main dashboard for VoicePilot"""
    return render_template('index.html')

@app.route('/api/transcribe/start', methods=['POST'])
def start_transcription():
    """Start live transcription"""
    global transcription_thread, transcription_active
    
    if not transcription_active:
        transcription_active = True
        transcriber.start()
        transcription_thread = threading.Thread(target=transcriber.run)
        transcription_thread.start()
        return jsonify({'status': 'success', 'message': 'Transcription started'})
    else:
        return jsonify({'status': 'error', 'message': 'Transcription already running'}), 400

@app.route('/api/transcribe/stop', methods=['POST'])
def stop_transcription():
    """Stop live transcription"""
    global transcription_active
    
    if transcription_active:
        transcription_active = False
        transcriber.stop()
        return jsonify({'status': 'success', 'message': 'Transcription stopped'})
    else:
        return jsonify({'status': 'error', 'message': 'No active transcription'}), 400

@app.route('/api/transcribe/stream')
def stream_transcription():
    """Stream transcription results"""
    def generate():
        while transcription_active:
            transcript = transcriber.get_latest_transcript()
            if transcript:
                yield f"data: {json.dumps({'transcript': transcript, 'is_final': True})}\n\n"
            threading.Event().wait(0.1)
    
    return Response(generate(), mimetype='text/event-stream')

@app.route('/api/voice-command', methods=['POST'])
def handle_voice_command():
    """Process voice commands"""
    try:
        data = request.get_json()
        command = data.get('command', '').lower()
        if not command:
            return jsonify({'status': 'error', 'message': 'No command provided'}), 400
        
        response = process_voice_command(command)
        return jsonify(response)
    
    except Exception as e:
        logger.error(f"Error processing voice command: {e}")
        return jsonify({'status': 'error', 'message': f"Error processing command: {e}"}), 500

@app.route('/api/status')
def get_status():
    """Get system status"""
    return jsonify({
        'status': 'online',
        'mode': 'offline',
        'ai_ready': True,
        'microphone': True,
        'transcription_active': transcription_active
    })

@app.errorhandler(404)
def not_found(error):
    return render_template('error.html', error="Page not found"), 404

@app.errorhandler(500)
def internal_error(error):
    return render_template('error.html', error="Internal server error"), 500

if __name__ == '__main__':
    print("🚀 Starting VoicePilot...")
    print("📁 Offline Voice Assistant for File Management")
    print("🔒 Privacy-first, Edge AI powered")
    print("-" * 50)
    app.run(debug=True, host='0.0.0.0', port=5000)