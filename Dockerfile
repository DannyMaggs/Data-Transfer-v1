FROM python:3.9

# Install required packages
RUN pip install --no-cache-dir Flask==2.0.3 Werkzeug==2.0.3 msal openpyxl python-pptx requests python-dotenv

# Copy application files
COPY app.py /app/app.py
COPY update_ppt.py /app/update_ppt.py

# Set the working directory
WORKDIR /app

# Expose the port the app runs on
EXPOSE 5000

# Run the application
CMD ["flask", "run", "--host=0.0.0.0"]
