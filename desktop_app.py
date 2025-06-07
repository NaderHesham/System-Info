import customtkinter as ctk
import platform
import psutil
import os
import subprocess
import json
import threading
from PIL import Image
import time
from customtkinter import filedialog
from tkinter import messagebox

# ===================================================================
# DATA-GATHERING FUNCTIONS (Unchanged)
# ===================================================================
def get_powershell_output(command):
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        result = subprocess.run(
            ["powershell", "-Command", command], capture_output=True, text=True, startupinfo=startupinfo, check=False)
        return result.stdout.strip() if result.stdout.strip() else "N/A"
    except Exception: return "N/A"
def get_serial_number(): return get_powershell_output("(Get-CimInstance Win32_BIOS).SerialNumber")
def get_pc_model(): return get_powershell_output("(Get-CimInstance Win32_ComputerSystem).Model")
def get_gpu(): return get_powershell_output("(Get-CimInstance Win32_VideoController | Select-Object -First 1 -ExpandProperty Name)")
def get_cpu(): return get_powershell_output("(Get-CimInstance Win32_Processor | Select-Object -First 1 -ExpandProperty Name)")
def get_monitors():
    script = """
    $mon = Get-CimInstance -Namespace root\\wmi -ClassName WmiMonitorID | Select-Object -First 1
    if ($mon) {
      $name = [System.Text.Encoding]::ASCII.GetString($mon.UserFriendlyName).Trim([char]0)
      if ($name -and $name.Length -gt 2) { $name } else { $mon.InstanceName }
    } else { 'N/A' }
    """
    return get_powershell_output(script)
def get_ram_model():
    script = """
    Get-CimInstance Win32_PhysicalMemory |
    Select-Object -First 1 Manufacturer, PartNumber |
    ForEach-Object { "$($_.Manufacturer) $($_.PartNumber)" }
    """
    return get_powershell_output(script)
def get_ipv4_and_mac():
    script = """
    $adapter = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' -and $_.HardwareInterface -eq $true } | Sort-Object -Property InterfaceMetric | Select-Object -First 1
    if ($adapter) {
        $ip = Get-NetIPAddress -InterfaceIndex $adapter.InterfaceIndex -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike '169.*' } | Select-Object -First 1
        [PSCustomObject]@{ IPv4 = $ip.IPAddress; MAC  = $adapter.MacAddress } | ConvertTo-Json
    }
    """
    try:
        output = get_powershell_output(script)
        if output and output != "N/A":
            data = json.loads(output)
            return data.get("IPv4", "N/A"), data.get("MAC", "N/A")
        return "N/A", "N/A"
    except (json.JSONDecodeError, AttributeError): return "N/A", "N/A"
def get_disks():
    script = "Get-PhysicalDisk | Select-Object MediaType, SerialNumber, Size, FriendlyName | ConvertTo-Json"
    try:
        output = get_powershell_output(script)
        if not output or output == "N/A": return []
        data = json.loads(output)
        if not isinstance(data, list): data = [data]
        disks = []
        for disk_info in data:
            size_gb = round(int(disk_info['Size']) / (1024 ** 3), 2)
            disks.append({"name": disk_info['FriendlyName'],"type": disk_info['MediaType'],"serial": disk_info['SerialNumber'],"size": size_gb})
        return disks
    except (json.JSONDecodeError, KeyError): return []


# ==================================
# STYLING CONFIGURATION
# ==================================
class Style:
    BACKGROUND_COLOR = "#242434"
    CARD_BG_COLOR = "#2D2D42"
    TITLE_COLOR = "#A4A5B3"
    VALUE_COLOR = "#FFFFFF"
    SSD_TAG_COLOR = "#2FA87C"
    BORDER_COLOR = "#2D2D42"

    FONT_FAMILY = "Segoe UI"
    TITLE_FONT = (FONT_FAMILY, 12)
    VALUE_FONT = (FONT_FAMILY, 14, "bold")
    SECTION_TITLE_FONT = (FONT_FAMILY, 16, "bold")
    LOADING_FONT = (FONT_FAMILY, 20, "bold")
    FOOTER_FONT = (FONT_FAMILY, 10)

# ==================================
# CUSTOM WIDGETS
# ==================================
class InfoCard(ctk.CTkFrame):
    def __init__(self, master, title, image_object):
        super().__init__(master, fg_color=Style.CARD_BG_COLOR, corner_radius=10, border_width=2, border_color=Style.BORDER_COLOR)
        self.grid_columnconfigure(1, weight=1)
        icon_label = ctk.CTkLabel(self, image=image_object, text="")
        icon_label.grid(row=0, column=0, rowspan=2, padx=15, pady=15)
        title_label = ctk.CTkLabel(self, text=title.upper(), font=Style.TITLE_FONT, text_color=Style.TITLE_COLOR, anchor="w")
        title_label.grid(row=0, column=1, padx=(0, 10), pady=(10, 0), sticky="ew")
        self.value_label = ctk.CTkLabel(self, text=" ", font=Style.VALUE_FONT, text_color=Style.VALUE_COLOR, anchor="w")
        self.value_label.grid(row=1, column=1, padx=(0, 10), pady=(0, 10), sticky="ew")
    def update_value(self, new_value):
        self.value_label.configure(text=new_value)

class StorageCard(ctk.CTkFrame):
    def __init__(self, master, name, disk_type, size_gb, image_object):
        super().__init__(master, fg_color=Style.CARD_BG_COLOR, corner_radius=10, border_width=2, border_color=Style.BORDER_COLOR)
        self.grid_columnconfigure(1, weight=1)
        icon_label = ctk.CTkLabel(self, image=image_object, text="")
        icon_label.grid(row=0, column=0, rowspan=2, padx=15, pady=15)
        name_label = ctk.CTkLabel(self, text=name, font=Style.VALUE_FONT, text_color=Style.VALUE_COLOR, anchor="w")
        name_label.grid(row=0, column=1, padx=(0, 10), pady=(10, 0), sticky="ew")
        bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_frame.grid(row=1, column=1, padx=10, pady=(0, 10), sticky="w")
        tag_label = ctk.CTkLabel(bottom_frame, text=disk_type, font=(Style.FONT_FAMILY, 10, "bold"), fg_color=Style.SSD_TAG_COLOR if disk_type == "SSD" else "grey", text_color="white", corner_radius=5, padx=5)
        tag_label.pack(side="left", anchor="w")
        size_str = f"{size_gb} TB" if size_gb >= 1000 else f"{size_gb} GB"
        size_label = ctk.CTkLabel(bottom_frame, text=f"Size: {size_str}", font=Style.TITLE_FONT, text_color=Style.TITLE_COLOR)
        size_label.pack(side="left", anchor="w", padx=(10, 0))

# ==================================
# MAIN APPLICATION
# ==================================
class App(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color=Style.BACKGROUND_COLOR)
        self.title("System Information")
        self.geometry("850x700")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.icons = {}
        self.info_cards = {}
        self.storage_frame = None
        self.system_data = {}

        self.main_content_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_content_frame.grid(row=0, column=0, sticky="nsew")
        self.main_content_frame.grid_columnconfigure(0, weight=1)

        self.loading_frame = ctk.CTkFrame(self, fg_color=Style.BACKGROUND_COLOR)
        self.loading_frame.grid(row=0, column=0, sticky="nsew")


        # <--- MODIFIED SECTION: Configure Loading Screen for perfect centering ---
        # This footer will be at the very bottom of the loading screen
        loading_footer = ctk.CTkLabel(self.loading_frame, text="Developed By IT Team \u00ae", font=Style.FOOTER_FONT, text_color=Style.TITLE_COLOR)
        loading_footer.pack(side="bottom", pady=10)

        # This frame will hold the centered content (label and progress bar)
        loading_center_frame = ctk.CTkFrame(self.loading_frame, fg_color="transparent")
        loading_center_frame.pack(expand=True) # expand=True pushes this frame to the middle

        loading_label = ctk.CTkLabel(loading_center_frame, text="Collecting System Data...", font=Style.LOADING_FONT, text_color=Style.VALUE_COLOR)
        loading_label.pack(pady=(0, 20)) # Pack label on top

        self.progress_bar = ctk.CTkProgressBar(loading_center_frame, orientation="horizontal", mode="indeterminate")
        self.progress_bar.pack(padx=50, fill="x") # Pack progress bar below
        # <--- END OF MODIFIED SECTION ---


        self._load_icons()
        self._create_widgets()
        self.load_data_in_thread()

    def show_loading_screen(self):
        self.loading_frame.tkraise()
        self.progress_bar.start()

    def hide_loading_screen(self):
        self.progress_bar.stop()
        self.main_content_frame.tkraise()

    def _load_icons(self):
        script_directory = os.path.dirname(os.path.abspath(__file__))
        icon_folder = os.path.join(script_directory, "icons")
        icon_names = [ "pc_name", "username", "ipv4", "serial", "pc_model", "mac", "cpu", "ram", "gpu", "screen", "storage", "info_section", "hardware_section", "storage_section" ]
        for name in icon_names:
            try:
                path = os.path.join(icon_folder, f"{name}.png")
                self.icons[name] = ctk.CTkImage(Image.open(path), size=(24, 24))
            except FileNotFoundError:
                print(f"Warning: Icon '{name}.png' not found in '{icon_folder}'")
                self.icons[name] = None

    def _create_widgets(self):
        gen_info_frame = self._create_section(self.main_content_frame, "General Information", "info_section", 0)
        gen_info_frame.grid_columnconfigure((0, 1, 2), weight=1)
        self.info_cards['pc_name'] = InfoCard(gen_info_frame, "PC Name", self.icons.get('pc_name'))
        self.info_cards['pc_name'].grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.info_cards['username'] = InfoCard(gen_info_frame, "Username", self.icons.get('username'))
        self.info_cards['username'].grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.info_cards['ipv4'] = InfoCard(gen_info_frame, "IPv4 Address", self.icons.get('ipv4'))
        self.info_cards['ipv4'].grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        self.info_cards['serial'] = InfoCard(gen_info_frame, "Serial Number", self.icons.get('serial'))
        self.info_cards['serial'].grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        self.info_cards['pc_model'] = InfoCard(gen_info_frame, "PC Model", self.icons.get('pc_model'))
        self.info_cards['pc_model'].grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.info_cards['mac'] = InfoCard(gen_info_frame, "MAC Address", self.icons.get('mac'))
        self.info_cards['mac'].grid(row=1, column=2, padx=5, pady=5, sticky="ew")

        hw_frame = self._create_section(self.main_content_frame, "Hardware Components", "hardware_section", 1)
        hw_frame.grid_columnconfigure((0, 1), weight=1)
        self.info_cards['cpu'] = InfoCard(hw_frame, "CPU", self.icons.get('cpu'))
        self.info_cards['cpu'].grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.info_cards['ram'] = InfoCard(hw_frame, "RAM", self.icons.get('ram'))
        self.info_cards['ram'].grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.info_cards['gpu'] = InfoCard(hw_frame, "GPU", self.icons.get('gpu'))
        self.info_cards['gpu'].grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        self.info_cards['screen'] = InfoCard(hw_frame, "Screen Model", self.icons.get('screen'))
        self.info_cards['screen'].grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        self.storage_frame = self._create_section(self.main_content_frame, "Storage Devices", "storage_section", 2)
        self.storage_frame.grid_columnconfigure(0, weight=1)
        
        button_frame = ctk.CTkFrame(self.main_content_frame, fg_color="transparent")
        button_frame.grid(row=3, column=0, padx=20, pady=(20, 10), sticky="ew")
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)

        refresh_button = ctk.CTkButton(button_frame, text="Refresh Data", command=self.load_data_in_thread, height=40)
        refresh_button.grid(row=0, column=0, padx=(0, 5), sticky="ew")

        export_button = ctk.CTkButton(button_frame, text="Export to TXT", command=self.export_data, height=40)
        export_button.grid(row=0, column=1, padx=(5, 0), sticky="ew")

        footer_label = ctk.CTkLabel(self.main_content_frame, text="Developed By IT Team \u00ae", font=Style.FOOTER_FONT, text_color=Style.TITLE_COLOR)
        footer_label.grid(row=4, column=0, padx=20, pady=(0, 10), sticky="s")

    def _create_section(self, parent, title, icon_key, row):
        section_frame = ctk.CTkFrame(parent, fg_color="transparent")
        section_frame.grid(row=row, column=0, padx=20, pady=10, sticky="ew")
        section_frame.grid_columnconfigure(1, weight=1)
        title_icon = ctk.CTkLabel(section_frame, image=self.icons.get(icon_key), text="")
        title_icon.grid(row=0, column=0, padx=(0, 10), pady=10)
        title_label = ctk.CTkLabel(section_frame, text=title, font=Style.SECTION_TITLE_FONT, text_color=Style.TITLE_COLOR)
        title_label.grid(row=0, column=1, pady=10, sticky="w")
        content_frame = ctk.CTkFrame(section_frame, fg_color="transparent")
        content_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        return content_frame

    def fetch_and_update_data(self):
        ipv4, mac = get_ipv4_and_mac()
        
        self.system_data = {
            "PC Name": platform.node(),
            "Username": os.getenv("USERNAME") or "N/A",
            "IPv4 Address": ipv4,
            "MAC Address": mac,
            "Serial Number": get_serial_number(),
            "PC Model": get_pc_model(),
            "CPU": get_cpu(),
            "RAM": f"{round(psutil.virtual_memory().total / (1024 ** 3), 1)} GB ({get_ram_model()})",
            "GPU": get_gpu(),
            "Screen Model": get_monitors(),
            "Disks": get_disks()
        }

        self.info_cards['pc_name'].update_value(self.system_data["PC Name"])
        self.info_cards['username'].update_value(self.system_data["Username"])
        self.info_cards['ipv4'].update_value(self.system_data["IPv4 Address"])
        self.info_cards['mac'].update_value(self.system_data["MAC Address"])
        self.info_cards['serial'].update_value(self.system_data["Serial Number"])
        self.info_cards['pc_model'].update_value(self.system_data["PC Model"])
        self.info_cards['cpu'].update_value(self.system_data["CPU"])
        self.info_cards['ram'].update_value(self.system_data["RAM"])
        self.info_cards['gpu'].update_value(self.system_data["GPU"])
        self.info_cards['screen'].update_value(self.system_data["Screen Model"])

        for widget in self.storage_frame.winfo_children(): widget.destroy()
        
        if self.system_data["Disks"]:
            for i, disk in enumerate(self.system_data["Disks"]):
                storage_card = StorageCard(self.storage_frame, disk['name'], disk['type'], disk['size'], self.icons.get('storage'))
                storage_card.grid(row=i, column=0, padx=5, pady=5, sticky="ew")
        else:
            no_disk_label = ctk.CTkLabel(self.storage_frame, text="No storage devices found.", text_color=Style.TITLE_COLOR)
            no_disk_label.grid(row=0, column=0, padx=5, pady=5)
        
        self.after(500, self.hide_loading_screen)

    def load_data_in_thread(self):
        self.show_loading_screen()
        thread = threading.Thread(target=self.fetch_and_update_data)
        thread.daemon = True
        thread.start()

    def export_data(self):
        if not self.system_data:
            messagebox.showwarning("No Data", "Please refresh data before exporting.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
            title="Save System Information"
        )

        if not file_path:
            return

        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write("SYSTEM INFORMATION REPORT\n")
                f.write("="*30 + "\n\n")

                for key, value in self.system_data.items():
                    if key != "Disks":
                        f.write(f"{key:<20}: {value}\n")

                f.write("\n" + "="*30 + "\n")
                f.write("STORAGE DEVICES\n")
                f.write("="*30 + "\n\n")

                if self.system_data["Disks"]:
                    for disk in self.system_data["Disks"]:
                        f.write(f"  - Name: {disk['name']}\n")
                        f.write(f"    Type: {disk['type']}\n")
                        f.write(f"    Size: {disk['size']} GB\n")
                        f.write(f"    Serial: {disk['serial']}\n\n")
                else:
                    f.write("No storage devices found.\n")

            messagebox.showinfo("Success", f"Data successfully exported to:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data: {e}")

# ==================================
# MAIN EXECUTION BLOCK
# ==================================
if __name__ == "__main__":
    app = App()
    app.mainloop()