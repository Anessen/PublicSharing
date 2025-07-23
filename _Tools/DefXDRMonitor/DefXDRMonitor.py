import tkinter as tk
from tkinter import ttk, messagebox
import requests
import json
from datetime import datetime
import msal
import atexit
import os

# --- Configuration ---
# Enter your Microsoft Entra ID application details here
TENANT_ID = "YOUR_TENANT_ID"
CLIENT_ID = "YOUR_CLIENT_ID"

# API scope and endpoint
API_SCOPE = ["https://api.security.microsoft.com/Incident.Read.All"]
API_ENDPOINT = "https://api.security.microsoft.com/api/incidents"
CACHE_FILE = "token_cache.bin"

class IncidentViewer(tk.Tk):
    def __init__(self, msal_app, token_cache):
        super().__init__()
        self.title("Microsoft Defender XDR Incident Viewer")
        self.geometry("1200x600")

        self.msal_app = msal_app
        self.token_cache = token_cache
        self.access_token = None
        self.refresh_rate = tk.IntVar(value=300)  # Default refresh rate: 5 minutes

        self.create_widgets()
        self.handle_authentication()

    def create_widgets(self):
        # Frame for controls
        control_frame = ttk.Frame(self)
        control_frame.pack(side="top", fill="x", padx=10, pady=5)

        ttk.Label(control_frame, text="Refresh Rate:").pack(side="left")
        ttk.Radiobutton(control_frame, text="60s", variable=self.refresh_rate, value=60).pack(side="left", padx=5)
        ttk.Radiobutton(control_frame, text="5 min", variable=self.refresh_rate, value=300).pack(side="left", padx=5)
        ttk.Radiobutton(control_frame, text="10 min", variable=self.refresh_rate, value=600).pack(side="left", padx=5)
        
        self.refresh_button = ttk.Button(control_frame, text="Refresh Now", command=self.fetch_incidents)
        self.refresh_button.pack(side="left", padx=20)
        
        self.status_label = ttk.Label(control_frame, text="Status: Initializing...")
        self.status_label.pack(side="right")

        # Treeview to display incidents
        self.tree = ttk.Treeview(self, columns=("Severity", "Title", "Status", "Created Time", "ID"), show="headings")
        self.tree.pack(side="left", fill="both", expand=True)

        # Define column headings and sorting
        self.tree.heading("Severity", text="Severity", command=lambda: self.sort_by_column("Severity", False))
        self.tree.heading("Title", text="Title")
        self.tree.heading("Status", text="Status")
        self.tree.heading("Created Time", text="Created Time")
        self.tree.heading("ID", text="Incident ID")

        # Define column widths
        self.tree.column("Severity", width=100)
        self.tree.column("Title", width=500)
        self.tree.column("Status", width=100)
        self.tree.column("Created Time", width=150)
        self.tree.column("ID", width=250)

        # Scrollbar
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Define color tags for severity
        self.tree.tag_configure('Critical', background='#ff4d4d')
        self.tree.tag_configure('High', background='#ff9966')
        self.tree.tag_configure('Medium', background='#ffcc66')
        self.tree.tag_configure('Low', background='#cce5ff')
        self.tree.tag_configure('Informational', background='#e0e0e0')

    def handle_authentication(self):
        if TENANT_ID == "YOUR_TENANT_ID" or CLIENT_ID == "YOUR_CLIENT_ID":
            messagebox.showerror("Configuration Error", "Please fill in your TENANT_ID and CLIENT_ID in the script.")
            self.destroy()
            return

        accounts = self.msal_app.get_accounts()
        result = None

        if accounts:
            result = self.msal_app.acquire_token_silent(API_SCOPE, account=accounts[0])

        if not result:
            self.status_label.config(text="Status: Awaiting user login in browser...")
            try:
                result = self.msal_app.acquire_token_interactive(scopes=API_SCOPE)
            except Exception as e:
                messagebox.showerror("Authentication Error", f"Could not acquire token interactively: {e}")
                self.destroy()
                return

        if "access_token" in result:
            self.access_token = result["access_token"]
            self.status_label.config(text="Authentication Successful. Fetching incidents...")
            self.fetch_and_schedule()
        else:
            messagebox.showerror("Authentication Failed", f"Could not acquire access token. Error: {result.get('error_description')}")
            self.destroy()

    def fetch_incidents(self):
        if not self.access_token:
            messagebox.showwarning("Warning", "Not authenticated.")
            return

        headers = {"Authorization": f"Bearer {self.access_token}"}
        
        try:
            response = requests.get(API_ENDPOINT, headers=headers)
            response.raise_for_status()
            incidents = response.json().get("value", [])
            self.update_incident_list(incidents)
            self.status_label.config(text=f"Last updated: {datetime.now().strftime('%H:%M:%S')}")

        except requests.exceptions.RequestException as e:
            if e.response.status_code == 401:
                self.status_label.config(text="Token expired. Re-authenticating...")
                self.handle_authentication() # Re-authenticate if token expired
            else:
                self.status_label.config(text=f"Error fetching incidents: {e}")

    def update_incident_list(self, incidents):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        severity_map = {"Critical": 1, "High": 2, "Medium": 3, "Low": 4, "Informational": 5, "Unknown": 6}
        sorted_incidents = sorted(incidents, key=lambda i: (severity_map.get(i.get('severity'), 6), i.get('createdDateTime', '')), reverse=False)

        for incident in sorted_incidents:
            severity = incident.get('severity', 'Unknown')
            title = incident.get('title', 'N/A')
            status = incident.get('status', 'N/A')
            created_time_str = incident.get('createdDateTime', '')
            created_time = datetime.strptime(created_time_str, '%Y-%m-%dT%H:%M:%S.%fZ').strftime('%Y-%m-%d %H:%M:%S') if created_time_str else 'N/A'
            incident_id = incident.get('incidentId', 'N/A')

            tag = severity if severity in ['Critical', 'High', 'Medium', 'Low', 'Informational'] else ''
            self.tree.insert("", "end", values=(severity, title, status, created_time, incident_id), tags=(tag,))
            
    def fetch_and_schedule(self):
        self.fetch_incidents()
        self.after(self.refresh_rate.get() * 1000, self.fetch_and_schedule)

    def sort_by_column(self, col, reverse):
        # This re-fetches and applies the default sort which is by severity
        self.fetch_incidents()

def main():
    # Set up token cache
    token_cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        token_cache.deserialize(open(CACHE_FILE, "r").read())
    atexit.register(lambda:
        open(CACHE_FILE, "w").write(token_cache.serialize()) if token_cache.has_state_changed else None
    )

    # Initialize MSAL application
    msal_app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=token_cache
    )

    # Start the GUI
    app = IncidentViewer(msal_app, token_cache)
    app.mainloop()

if __name__ == "__main__":
    main()