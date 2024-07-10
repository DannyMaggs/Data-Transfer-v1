from flask import Flask, request, jsonify
import subprocess
from dotenv import load_dotenv
import os

load_dotenv()

app = Flask(__name__)

@app.route('/update_ppt', methods=['POST'])
def update_ppt():
    data = request.get_json()
    source_file_id = data['sourceFileId']
    destination_file_id = data['destinationFileId']

    # Execute the update_ppt.py script with the provided file IDs
    result = subprocess.run(['python3', 'update_ppt.py', source_file_id, destination_file_id], capture_output=True, text=True)

    if result.returncode == 0:
        return jsonify({"message": "PowerPoint updated successfully", "output": result.stdout})
    else:
        return jsonify({"message": "Failed to update PowerPoint", "output": result.stderr}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
