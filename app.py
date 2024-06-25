from flask import Flask, request, jsonify
import subprocess
import os

app = Flask(__name__)

@app.route('/update_ppt', methods=['POST'])
def update_ppt():
    source_file_id = request.json.get('sourceFileId')
    destination_file_id = request.json.get('destinationFileId')

    if not source_file_id or not destination_file_id:
        return jsonify({"error": "sourceFileId and destinationFileId are required"}), 400

    # Run the Python script
    result = subprocess.run(['python3', 'update_ppt.py', source_file_id, destination_file_id], capture_output=True, text=True)

    if result.returncode != 0:
        return jsonify({"error": result.stderr}), 500

    return jsonify({"message": "PowerPoint updated successfully", "output": result.stdout}), 200

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
