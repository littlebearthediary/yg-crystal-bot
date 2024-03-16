import requests

def send_line_notification(message):
    url = "https://notify-api.line.me/api/notify"
    token = "kFjE4JEoUxrUyiidpdrbokSmgPStzaigsr0auHx2k3D"  # Replace with your access token
    headers = {"Authorization": f"Bearer {token}"}
    data = {"message": message}
    response = requests.post(url, headers=headers, data=data)
    if response.status_code == 200:
        print("Notification sent successfully!")
    else:
        print("Failed to send notification")