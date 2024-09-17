#####################################################################
########################### vATISLoad.py ############################
#####################################################################
import subprocess, sys, os, time, json, re, uuid, ctypes


# pip uninstall -y pyautogui pyperclip pygetwindow pywin32 pywinutils psutil
import importlib.util as il
if None in [il.find_spec('pyautogui'), il.find_spec('pyperclip'), \
            il.find_spec('pygetwindow'), il.find_spec('win32api'), \
            il.find_spec('psutil'), il.find_spec('requests'), \
            il.find_spec('pyscreeze')]:
    subprocess.check_call([sys.executable, '-m', 'pip', 
                       'install', 'pyautogui']);
    subprocess.check_call([sys.executable, '-m', 'pip', 
                           'install', 'pyperclip']);
    subprocess.check_call([sys.executable, '-m', 'pip', 
                           'install', 'pygetwindow']);
    subprocess.check_call([sys.executable, '-m', 'pip', 
                           'install', 'pywinutils']);
    subprocess.check_call([sys.executable, '-m', 'pip', 
                           'install', 'psutil']);
    subprocess.check_call([sys.executable, '-m', 'pip', 
                           'install', 'tkinter']);
    subprocess.check_call([sys.executable, '-m', 'pip', 
                           'install', 'requests']);
    subprocess.check_call([sys.executable, '-m', 'pip', 
                           'install', 'pyscreeze']);
    subprocess.check_call([sys.executable, '-m', 'pip', 
                           'install', '--upgrade', 'Pillow']);
    os.system('cls')
else:
    os.system('cls')
    
import requests, pyautogui, psutil, pyperclip, pygetwindow as gw, tkinter as tk
from win32 import win32api, win32gui, win32gui, win32process
from tkinter import messagebox, Listbox
from win32.lib import win32con

scale_factor = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100

def read_config():
    profiles = {}
    timeout = 2
    config = os.getenv('LOCALAPPDATA') + '\\vATIS-4.0\\vATISLoadConfig.json'
    if not os.path.isfile(config):
        config = os.getenv('LOCALAPPDATA') + '\\vATIS\\vATISLoadConfig.json'
        if not os.path.isfile(config):
            return profiles, timeout

    with open(config, 'r') as f:
        data = json.loads(f.read())

    return data['facilities'], data['timeout']

def add_profile(facility, airports):
    facility = facility.upper()
    airports = re.sub('[^0-9A-z,/]', '', airports).upper().split(',')
    if len(facility) == 0 or len(airports) == 0:
        return

    config_folder = os.getenv('LOCALAPPDATA') + '\\vATIS-4.0'
    if not os.path.exists(config_folder):
        config_folder = os.getenv('LOCALAPPDATA') + '\\vATIS'
    if not os.path.exists(config_folder):
        messagebox.showerror("Error", "No vATIS folder found.")
        return

    config = os.path.join(config_folder, 'vATISLoadConfig.json')
    data = {}
    if not os.path.isfile(config):
        data['facilities'] = {}
        data['timeout'] = 2
    else:
        with open(config, 'r') as f:
            data = json.loads(f.read())

    data['facilities'][facility] = airports

    with open(config, 'w') as f_out:
        f_out.write(json.dumps(data, indent=4))

    messagebox.showinfo("Success", f"Profile '{facility}' with airports {', '.join(airports)} added.")
    return facility
def delete_profile(profile):
    config = os.getenv('LOCALAPPDATA') + '\\vATIS-4.0\\vATISLoadConfig.json'
    if not os.path.isfile(config):
        config = os.getenv('LOCALAPPDATA') + '\\vATIS\\vATISLoadConfig.json'

    if not os.path.isfile(config):
        messagebox.showerror("Error", "Configuration file not found.")
        return

    with open(config, 'r') as f:
        data = json.loads(f.read())

    if profile in data['facilities']:
        del data['facilities'][profile]
        messagebox.showinfo("Deleted", f"Profile '{profile}' deleted successfully.")
    else:
        messagebox.showwarning("Not Found", f"Profile '{profile}' not found.")

    with open(config, 'w') as f_out:
        f_out.write(json.dumps(data, indent=4))

    return
def check_datis_profile(profile):
    # Load vATIS configuration
    config = os.getenv('LOCALAPPDATA') + '\\vATIS-4.0\\AppConfig.json'
    if not os.path.isfile(config):
        config = os.getenv('LOCALAPPDATA') + '\\vATIS\\AppConfig.json'

    with open(config, 'r') as f:
        data = json.loads(f.read())

    added_datis = 0
    for i in range(len(data['profiles'])):
        if profile not in data['profiles'][i]['name']:
            continue

        for j in range(len(data['profiles'][i]['composites'])):
            comp = data['profiles'][i]['composites'][j]
            comp_ident = comp['identifier'][1:]

            if len(comp['presets']) == 0 or 'D-ATIS' not in comp['presets'][0]['name']:
                if added_datis == 0:
                    messagebox.showinfo("Information", "Creating D-ATIS presets...")

                datis_preset = {}
                if len(comp['presets']) != 0:
                    comp['presets'].insert(0, comp['presets'][0].copy())
                    datis_preset = comp['presets'][0]
                else:
                    comp['presets'].insert(0, datis_preset)

                # Set preset attributes
                datis_preset['id'] = str(uuid.uuid4())
                datis_preset['name'] = 'D-ATIS'
                datis_preset['airportConditions'] = ''
                datis_preset['notams'] = ''
                datis_preset['externalGenerator'] = {'enabled': False}
                added_datis += 1

    if added_datis > 0:
        save_output = messagebox.askyesno("Save Presets", "Would you like to save the new D-ATIS presets?")
        if save_output:
            with open(config, 'w') as f_out:
                f_out.write(json.dumps(data, indent=4))
            messagebox.showinfo("Saved", "New D-ATIS presets saved successfully.")
        else:
            messagebox.showinfo("Cancelled", "Preset additions were cancelled.")
    else:
        messagebox.showinfo("No Changes", "No new D-ATIS presets were added.")

def open_vATIS():
    os.system('taskkill /f /im vATIS.exe 2>nul 1>nul')
    exe = os.getenv('LOCALAPPDATA') + '\\vATIS-4.0\\Application\\vATIS.exe'
    if not os.path.isfile(exe):
        exe = os.getenv('LOCALAPPDATA') + '\\vATIS\\Application\\vATIS.exe'
    subprocess.Popen(exe)

    # Wait for the vATIS window to appear
    for _ in range(10):
        vatis_open = any('vATIS Profiles' in window.title for window in gw.getAllWindows())
        if vatis_open:
            time.sleep(1)
            return
        else:
            time.sleep(0.5)

def center_win(exe_name, window_title):
    win = None

    # Select the window
    for window in gw.getAllWindows():
        hwnd = window._hWnd
        thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(process_id)
        process_name = process.name()
        process_path = process.exe()
        if exe_name in process_path and window.title == window_title:
            win = window

    # Move the window to the center
    screen_dim = [win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1)]
    win.moveTo(int((screen_dim[0] - win.size[0]) / 2), int((screen_dim[1] - win.size[1]) / 2))

    # Bring the window to the foreground
    hwnd = win._hWnd
    pyautogui.FAILSAFE = False
    pyautogui.press('alt')
    win32gui.SetForegroundWindow(hwnd)
    return win

def click_xy(xy, win, d=0):
    x, y = xy
    x *= scale_factor
    y *= scale_factor
    x += win.left
    y += win.top
    time.sleep(d)
    pyautogui.moveTo(x, y)
    pyautogui.click()

def run_profile(profile):
    global root
    profiles, TIMEOUT = read_config()

    # Open vATIS and center the window
    open_vATIS()
    win = center_win('vATIS.exe', 'vATIS Profiles')

    # Select profile
    n_profile = get_profile_pos(profile, sort=True)
    if n_profile == -1:
        messagebox.showerror("Error", "Selected profile not found!")
        return
    elif n_profile > 18:
        messagebox.showerror("Error", "vATISLoad does not support more than 19 profiles.")
        return
    loc_profile = [90, 40 + 14 * n_profile]

    click_xy(loc_profile, win)
    pyautogui.press('enter')

    time.sleep(1)

    # Center the vATIS window and interact with it
    win = center_win('vATIS.exe', 'vATIS')

    AIRPORTS = profiles[profile]
    for ident in AIRPORTS:
        tab = get_tab(ident, profile)
        if tab == -1:
            messagebox.showinfo("Not Found", f'{ident} NOT FOUND.')
            continue

        loc_tab = [38.6 + 53.6 * tab, 64]
        click_xy(loc_tab, win)
        click_xy([400, 330], win)
        pyautogui.press(['up', 'enter'])

        atis, code = get_atis(ident)

        if len(atis) > 0:
            click_xy([200, 250], win)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyperclip.copy(atis[0])
            pyautogui.hotkey('ctrl', 'v')
            click_xy([40, 295], win)

        if len(atis) > 1:
            click_xy([600, 250], win)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyperclip.copy(atis[1])
            pyautogui.hotkey('ctrl', 'v')
            click_xy([415, 295], win)

        if len(code) == 0 or len(atis[0]) == 0:
            messagebox.showinfo("No Code", f'{ident.upper()} - UN')
            continue

        click_xy([720, 330], win)
    root.destroy()
def get_profiles():
    config = os.getenv('LOCALAPPDATA') + '\\vATIS-4.0\\AppConfig.json'
    if not os.path.isfile(config):
        config = os.getenv('LOCALAPPDATA') + '\\vATIS\\AppConfig.json'

    f = open(config, 'r')
    data = json.loads(f.read())
    f.close()
    
    profiles = []
    for profile in data['profiles']:
        profiles.append(profile['name'])
    
    return profiles
        
def get_profile_pos(name, sort, exact=False):
    profiles = get_profiles()
    if sort:
        profiles.sort()
        for i in range(0, len(profiles)):
            profiles[i] = re.sub(r'[^A-z0-9 ]', '', profiles[i])
    for i in range(0, len(profiles)):
        prof = profiles[i]
        if re.sub(r'[^A-z0-9 ]', '', name) in prof and not exact:
            return i
        elif name == prof:
            return i

    return -1

def get_idents(n_profile):
    config = os.getenv('LOCALAPPDATA') + '\\vATIS-4.0\\AppConfig.json'

    if not os.path.isfile(config):
        config = os.getenv('LOCALAPPDATA') + '\\vATIS\\AppConfig.json'

    f = open(config, 'r')
    data = json.loads(f.read())
    f.close()

    prof_name = data['profiles'][n_profile]['name']
    idents = []

    for comp in data['profiles'][n_profile]['composites']:
        idents.append([comp['identifier'], 
                       comp['atisType'][0].replace('C', 'Z')])
    idents.sort()
    
    return idents

def get_tab(airport, PROFILE):
    idents = get_idents(get_profile_pos(PROFILE, sort=False))
    for i in range(0, len(idents)):
        ident = idents[i]
        if '/' in airport:
            apt, atis_type = airport.split('/')
            if apt in ident[0]:
                if atis_type[0] != ident[1][0]:
                    return i
        elif airport in ident[0] and ident[1] == 'Z':
            return i
    return -1

def get_atis(ident):
    atis_type = 'C'
    if '/' in ident:
        ident, atis_type = ident.split('/')
    if len(ident) == 3:
        ident = 'K' + ident
    url = 'https://datis.clowd.io/api/' + ident

    atis_info, code = [], ''
    atis_data = json.loads(requests.get(url).text)
    if 'error' in atis_data:
        return [], ''
    
    for n in range(0, len(atis_data)):
        datis = atis_data[n]['datis']
        if atis_type == 'C' and n == 0:
            code = atis_data[n]['code']
        elif atis_type == 'D' and atis_data[n]['type'] == 'dep':
            code = atis_data[n]['code']
        elif atis_type == 'A' and atis_data[n]['type'] == 'arr':
            code = atis_data[n]['code']

        datis = re.sub('.*INFO [A-Z] [0-9][0-9][0-9][0-9]Z. ', '', datis)
        datis = '. '.join(datis.split('. ')[1:])
        datis = re.sub(' ...ADVS YOU HAVE.*', '', datis)
        datis = datis.replace('...', '/./').replace('..', '.') \
            .replace('/./', '...').replace('  ', ' ').replace(' . ', '. ') \
            .replace(', ,', ',').replace(' ; ', '; ').replace(' .,', ' ,') \
            .replace(' , ', ', ').replace('., ', ', ').replace('&amp;', '&')

        info = []
        if 'NOTAMS' in datis:
            info = datis.split('NOTAMS... ')
        elif 'NOTICE TO AIR MISSIONS. ' in datis:
            info = datis.split('NOTICE TO AIR MISSIONS. ')
        else:
            info = [datis, '']
        
        if n == 0:
            atis_info = info[:]
        else:
            if atis_type == 'A':
                atis_info[0] = re.sub(r'\s+', ' ', atis_info[0])
            elif atis_type == 'D':
                atis_info[0] = re.sub(r'\s+', ' ', info[0])
            else:
                atis_info[0] = re.sub(r'\s+', ' ', 
                                      info[0] + ' ' + atis_info[0])

    return atis_info, code

def char_position(letter):
    if len(letter) == 0:
        return -1
    return ord(letter.lower()) - 97

def create_ui():
    global root
    root = tk.Tk()
    root.title("vATIS Profile Manager")

    # Load existing profiles
    profiles, TIMEOUT = read_config()

    # Listbox to display profiles
    tk.Label(root, text="Existing Profiles").grid(row=0, column=0)
    profile_listbox = Listbox(root)
    for profile in profiles:
        profile_listbox.insert(tk.END, profile)
    profile_listbox.grid(row=1, column=0)

    # Input for new facility
    tk.Label(root, text="Facility").grid(row=2, column=0)
    facility_entry = tk.Entry(root)
    facility_entry.grid(row=2, column=1)

    # Input for airports
    tk.Label(root, text="Airports (comma separated)").grid(row=3, column=0)
    airports_entry = tk.Entry(root)
    airports_entry.grid(row=3, column=1)

    # Add Profile button
    def add_profile_action():
        facility = facility_entry.get()
        airports = airports_entry.get()
        if facility and airports:
            added_facility = add_profile(facility, airports)
            profile_listbox.insert(tk.END, added_facility)
        else:
            messagebox.showwarning("Input Error", "Please fill both fields.")

    add_button = tk.Button(root, text="Add Profile", command=add_profile_action)
    add_button.grid(row=4, column=0, columnspan=2)
    
    def delete_profile_action():
        selected_profile = profile_listbox.get(tk.ACTIVE)
        if selected_profile:
            delete_profile(selected_profile)
            profile_listbox.delete(tk.ACTIVE)
        else:
            messagebox.showwarning("Selection Error", "Please select a profile.")

    delete_button = tk.Button(root, text="Delete Profile", command=delete_profile_action)
    delete_button.grid(row=4, column=1, columnspan=2)

    # Run automation button
    def run_automation():
        selected_profile = profile_listbox.get(tk.ACTIVE)
        if selected_profile:
            run_profile(selected_profile)
        else:
            messagebox.showwarning("Selection Error", "Please select a profile.")

    run_button = tk.Button(root, text="Run profile", command=run_automation)
    run_button.grid(row=5, column=0, columnspan=2)

    root.mainloop()

# Launch the UI instead of printing profiles in console
if __name__ == '__main__':
    create_ui()
