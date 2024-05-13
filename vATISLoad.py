#####################################################################
########################### vATISLoad.py ############################
#####################################################################
import subprocess, sys, os, time, json, re, requests, uuid, ctypes

# pip uninstall -y pyautogui pyperclip pygetwindow pywin32 pywinutils psutil
import importlib.util as il
if None in [il.find_spec('pyautogui'), il.find_spec('pyperclip'), \
            il.find_spec('pygetwindow'), il.find_spec('win32api'), \
            il.find_spec('psutil')]:
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
    os.system('cls')
    os.execv(sys.executable, ['python'] + sys.argv)
else:
    os.system('cls')
    
import pyautogui, psutil, pyperclip, pygetwindow as gw
from win32 import win32api, win32gui, win32gui, win32process
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
        
    f = open(config, 'r')
    data = json.loads(f.read())
    f.close
    
    return data['facilities'], data['timeout']

def check_datis_profile(profile):
    config = os.getenv('LOCALAPPDATA') + '\\vATIS-4.0\\AppConfig.json'
    if not os.path.isfile(config):
        config = os.getenv('LOCALAPPDATA') + '\\vATIS\\AppConfig.json'

    f = open(config, 'r')
    data = json.loads(f.read())
    f.close
    
    added_datis = 0
    for i in range(0, len(data['profiles'])):
        if not profile in data['profiles'][i]['name']:
            continue
        
        for j in range(0, len(data['profiles'][i]['composites'])):
            comp = data['profiles'][i]['composites'][j]
            comp_ident = comp['identifier'][1:]
            
            if len(comp['presets']) == 0 or \
                'D-ATIS' not in comp['presets'][0]['name']:
                if added_datis == 0:
                    print('=========== vATISLoad ===========')
                
                datis_preset = {}
                if len(comp['presets']) != 0:
                    comp['presets'].insert(0, comp['presets'][0].copy())
                    datis_preset = comp['presets'][0]
                else:
                    comp['presets'].insert(0, datis_preset)
                datis_preset['id'] = str(uuid.uuid4())
                datis_preset['name'] = 'D-ATIS'
                datis_preset['airportConditions'] = ''
                datis_preset['notams'] = ''
                datis_preset['externalGenerator'] = {'enabled': False}
                print(f'Created D-ATIS preset for {comp_ident}')
                added_datis += 1
                
        if added_datis > 0:
            print('\nPress ENTER to save D-ATIS preset additions')
            save_output = input('Input any other text to cancel')
        
            if len(save_output) == 0:
                with open(config, 'w+') as f_out:
                    f_out.write(json.dumps(data, indent=2))
                    print('Saved new D-ATIS presets')
                    time.sleep(0.5)
                    os.system('cls')

def open_vATIS():
    os.system('taskkill /f /im vATIS.exe 2>nul 1>nul')
    exe = os.getenv('LOCALAPPDATA') + '\\vATIS-4.0\\Application\\vATIS.exe'
    if not os.path.isfile(exe):
        exe = os.getenv('LOCALAPPDATA') + '\\vATIS\\Application\\vATIS.exe'
    subprocess.Popen(exe);

def center_win(exe_name, window_title):
    win = None

    # Select window
    for window in gw.getAllWindows():
        hwnd = window._hWnd
        thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(process_id)
        process_name = process.name()
        process_path = process.exe()
        if exe_name in process_path:
            if window.title == window_title:
                win = window
    
    # Move window to center
    screen_dim = [win32api.GetSystemMetrics(0), \
                  win32api.GetSystemMetrics(1)]
    win.moveTo(int((screen_dim[0] - win.size[0]) / 2), \
                  int((screen_dim[1] - win.size[1]) / 2))
    
    # Move window to foreground
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
    
def get_profiles():
    config = os.getenv('LOCALAPPDATA') + '\\vATIS-4.0\\AppConfig.json'
    if not os.path.isfile(config):
        config = os.getenv('LOCALAPPDATA') + '\\vATIS\\AppConfig.json'

    f = open(config, 'r')
    data = json.loads(f.read())
    f.close
    
    profiles = []
    for profile in data['profiles']:
        profiles.append(profile['name'])
    
    return profiles

def get_profile_pos(name, sort, exact=False):
    profiles = get_profiles()
    if sort:
        for i in range(0, len(profiles)):
            profiles[i] = re.sub(r'[^A-z0-9]', '', profiles[i])
        profiles.sort()
    for i in range(0, len(profiles)):
        prof = profiles[i]
        if re.sub(r'[^A-z0-9]', '', name) in prof and not exact:
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
    f.close

    prof_name = data['profiles'][n_profile]['name']
    idents = []

    for comp in data['profiles'][n_profile]['composites']:
        idents.append([comp['identifier'], 
                       comp['atisType'][0].replace('C', 'Z')])
    idents.sort()
    
    return idents

def get_tab(airport, PROFILE):
    idents = get_idents(get_profile_pos(PROFILE, False))
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

def add_profile(facility, airports):
    facility = facility.upper()
    airports = re.sub('[^0-9A-z,/]', '', airports).upper().split(',')
    if len(facility) == 0 or len(airports) == 0:
        os.execv(sys.executable, ['python'] + sys.argv)
        return
    
    config_folder = os.getenv('LOCALAPPDATA') + '\\vATIS-4.0'
    if not os.path.exists(config_folder):
        config_folder = os.getenv('LOCALAPPDATA') + '\\vATIS'
    if not os.path.exists(config_folder):
        print('No vATIS folder found.')
        return
    
    config = os.path.join(config_folder, 'vATISLoadConfig.json')
    data = {}
    if not os.path.isfile(config):
        data['facilities'] = {}
        data['timeout'] = 2
    else:
        f = open(config, 'r')
        data = json.loads(f.read())
        f.close
        
    data['facilities'][facility] = airports
        
    with open(config, 'w+') as f_out:
        f_out.write(json.dumps(data, indent=2))
    
    os.system('cls')
    os.execv(sys.executable, ['python'] + sys.argv)

# Center command prompt
for win in gw.getAllWindows():
    if 'py.exe' in win.title or win.title == 'vATIS': 
        screen_dim = [win32api.GetSystemMetrics(0), \
            win32api.GetSystemMetrics(1)]
        win.moveTo(int((screen_dim[0] - win.size[0]) / 2) - 60, \
            int((screen_dim[1] - win.size[1]) / 2))

# Profile selection
print('=========== vATISLoad ===========')
profiles, TIMEOUT = read_config()
for i in range(0, len(profiles)):
    facility = list(profiles.keys())[i]
    airports = ', '.join(list(profiles.values())[i])
    print(f'({i}) {facility} - {airports}')
    
if len(profiles) == 0:
    print('No vATISLoadConfig.json file found!')
    print('Add profiles to create a configuration file.\n')
    
print('(A) Add new profile')

for i in range(0, 100):
    idx = input('\nProfile: ')
    if idx.isdigit():
        idx = int(idx)
        if idx < len(profiles):
            break
    elif idx.upper() == 'A':
        os.system('cls')
        print('=========== vATISLoad ===========')
        print('Input the vATIS facility name')
        print('e.g. \'Oakland ARTCC (ZOA)\'')
        print('e.g. \'ZOA\'')
        facility = input('\nFacility: ')
        os.system('cls')
        print('=========== vATISLoad ===========')
        print('Input airports separated by commas')
        print('DEP/ARR ATISes - add \'/D\' or \'/A\'')
        print('e.g. \'MIA/D, MIA/A, FLL, TPA, RSW\'')
        airports = input('\nAirports: ')
        add_profile(facility, airports)
        profiles, TIMEOUT = read_config()
        idx = len(profiles) - 1
        break
    print(f'Invalid input! Selection must be a number ' \
          + f'between 0 and {len(profiles) - 1}.')

os.system('cls')
PROFILE = list(profiles.keys())[idx]
AIRPORTS = list(profiles.values())[idx]

# Create missing D-ATIS presets
check_datis_profile(PROFILE)

print('=========== vATISLoad ===========')
print(f'[{PROFILE}]')

# Open vATIS
open_vATIS()
time.sleep(2.5)
pyautogui.PAUSE = 0.001

# Center 'vATIS Profiles' window and bring to foreground
win = center_win('vATIS.exe', 'vATIS Profiles')

# Select profile chosen above
win_bound = [win.left, win.top]
n_profile = get_profile_pos(PROFILE, sort=True)
loc_profile = [90, 40 + 14 * n_profile]

click_xy(loc_profile, win)
pyautogui.press('enter')

time.sleep(1)

# Center 'vATIS' window and bring to foreground
win = center_win('vATIS.exe', 'vATIS')

for ident in AIRPORTS:
    # Select tab for specified airport
    tab = get_tab(ident, PROFILE)
    if tab == -1:
        print(f'{ident} NOT FOUND.')
        continue
    loc_tab = [38.6 + 53.6 * tab, 64]
    click_xy(loc_tab, win)
    
    # Select first preset
    click_xy([400, 330], win)
    pyautogui.press(['up', 'enter'])
    
    # Get D-ATIS
    atis, code = get_atis(ident)
    
    # Enter ARPT COND
    if len(atis) > 0:
        click_xy([200, 250], win)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        pyperclip.copy(atis[0])
        pyautogui.hotkey('ctrl', 'v')
        click_xy([40, 295], win)
    
    # Enter NOTAMS
    if len(atis) > 1:
        click_xy([600, 250], win)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        pyperclip.copy(atis[1])
        pyautogui.hotkey('ctrl', 'v')
        click_xy([415, 295], win)
        
    if len(code) == 0 or len(atis[0]) == 0:
        print(f'{ident.upper()} - UN')
        continue
    
    # Connect ATIS
    click_xy([720, 330], win)
    state = ''
    for i in range(0, int(10 * TIMEOUT)):
        pix_x = int(scale_factor * 118)
        pix_y = int(scale_factor * 104)
        pix = pyautogui.pixel(win.left + pix_x, win.top + pix_y)
        # Check if METAR loads (white 'K')
        if pix[0] >= 248 or pix[1] >= 248:
            state = 'CON'
            break
        # Check if ATIS is already connected (red 'N')
        elif pix[0] >= 220 and pix[0] <= 230:
            state = 'ON'
            break
        time.sleep(.1)
    # Set ATIS code
    for i in range(0, char_position(code)):
        click_xy([62, 130], win)
    if state == 'CON':
        print(f'{ident.upper()} - {code}')
    elif state == 'ON':
        print(f'{ident.upper()} - OL/{code}')
    else:
        if len(code) > 0:
            print(f'{ident.upper()} - UN/{code}')
        else:  
            print(f'{ident.upper()} - UN')

pyperclip.copy('')
time.sleep(3)
win32gui.ShowWindow(win._hWnd, win32con.SW_MINIMIZE);

# Determine mouse position (and color) on held left click
def mouse_position(color=False):
    prev_xy = [-99999, -99999]
    for i in range(0, 1000):
        time.sleep(1)
        if win32api.GetAsyncKeyState(0x01) >= 0:
            continue
        x, y = pyautogui.position()
        win_bound = [win.left, win.top]
        out = x - win_bound[0], y - win_bound[1]
        if out[0] != prev_xy[0] or out[1] != prev_xy[1]:
            if not color:
                print(out)
            else:
                print(out, pyautogui.pixel(x, y))
            prev_xy = out[:]
            
# mouse_position()