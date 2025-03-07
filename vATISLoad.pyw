#####################################################################
############################# vATISLoad #############################
#####################################################################
import subprocess, sys, os, time, json, re, uuid, ctypes, asyncio
from datetime import datetime

import importlib.util as il
if None in [il.find_spec('pyautogui'), il.find_spec('pyperclip'), \
            il.find_spec('pygetwindow'), il.find_spec('win32api'), \
            il.find_spec('psutil'), il.find_spec('requests'), \
            il.find_spec('pyscreeze'), il.find_spec('websockets'), \
            il.find_spec('pynput')]:

    os.system('cmd /K \"cls & echo Updating required packages for vATISLoad.' + \
        ' & echo Please wait a few minutes for packages to install. & timeout 5 & exit\"')

    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', 'pip'])
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyautogui']);
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyperclip']);
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pygetwindow']);
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pywinutils']);
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pywin32']);
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'psutil']);
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'requests']);
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyscreeze']);
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'websockets']);
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pynput']);
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', 'Pillow']);

os.system('cls')

import pyautogui, pyperclip, pygetwindow, psutil, requests, websockets, pynput
from win32 import win32api, win32gui, win32gui, win32process
from win32.lib import win32con

scale_factor = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 150
tab_sizes = {'small': 70, 'large': 95, 'small_con': 90, 'large_con': 118}

# Set to False for testing
RUN_UPDATE = True

def update_vATISLoad():
    online_file = ''
    url = 'https://raw.githubusercontent.com/glott/vATISLoad/refs/heads/main/vATISLoad.pyw'
    try:
        online_file = requests.get(url).text.split('\n')
    except Exception as ignored:
        return

    up_to_date = True
    with open(sys.argv[0], 'r') as FileObj:
        i = 0
        for line in FileObj:
            if i > len(online_file) or len(line.strip()) != len(online_file[i].strip()):
                up_to_date = False
                break
            i += 1

    if up_to_date:
        return

    try:
        os.rename(sys.argv[0], sys.argv[0] + '.bak')
        with requests.get(url, stream=True) as r:
            r.raise_for_status()
            with open(sys.argv[0], 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192): 
                    f.write(chunk)

        os.remove(sys.argv[0] + '.bak')
        
    except Exception as ignored:
        if not os.path.isfile(sys.argv[0]) and os.path.isfile(sys.argv[0] + '.bak'):
            os.rename(sys.argv[0] + '.bak', sys.argv[0])

    os.execv(sys.executable, ['python'] + sys.argv)

def set_foreground_window(hwnd):
    pyautogui.FAILSAFE = False
    pyautogui.press('alt')
    win32gui.SetForegroundWindow(hwnd)

def get_win(exe_name, window_title):
    for i in range(0, 100):
        try:
            for window in pygetwindow.getAllWindows():
                thread_id, process_id = win32process.GetWindowThreadProcessId(window._hWnd)
                process = psutil.Process(process_id)
                process_name = process.name()
                process_path = process.exe()
                if exe_name in process_path:
                    if window.title == window_title:
                        set_foreground_window(window._hWnd)
                        return window
        except Exception as ignored:
            pass
        time.sleep(0.1)
    return None

def click_xy(xy, win, sf=True, d=0.01):
    x, y = xy
    if sf:
        x *= scale_factor
        y *= scale_factor
    x += win.left
    y += win.top
    time.sleep(d)
    pyautogui.moveTo(x, y)
    pyautogui.click()

def determine_active_profile():
    crc_profiles = os.getenv('LOCALAPPDATA') + '\\CRC\\Profiles'
    crc_name = ''
    crc_data = {}
    crc_lastused_time = '2020-01-01T08:00:00'
    for filename in os.listdir(crc_profiles):
        if filename.endswith('.json'): 
            file_path = os.path.join(crc_profiles, filename)
            with open(file_path, 'r') as f:
                data = json.load(f)
                dt1 = datetime.strptime(crc_lastused_time, '%Y-%m-%dT%H:%M:%S')
                if 'LastUsedAt' not in data or data['LastUsedAt'] == None:
                    continue
                dt2 = datetime.strptime(data['LastUsedAt'].split('.')[0], '%Y-%m-%dT%H:%M:%S')
                if dt2 > dt1:
                    crc_lastused_time = data['LastUsedAt'].split('.')[0]
                    crc_name = data['Name']
                    crc_data = data

    facility_id = ''
    try:
        facility_id = crc_data['DisplayWindowSettings'][0]['DisplaySettings'][0]['FacilityId']
    except Exception as ignored:
        pass
    
    vatis_profiles = os.getenv('LOCALAPPDATA') + '\\org.vatsim.vatis\\Profiles'
    for filename in os.listdir(vatis_profiles):
        if filename.endswith('.json'): 
            file_path = os.path.join(vatis_profiles, filename)
            with open(file_path, 'r') as f:
                data = json.load(f)
                data['name'] = data['name'].replace('[', '(').replace(']', ')')
                if not('(' in  data['name'] and ')' in data['name']):
                    continue
                vatis_abr = data['name'].split('(')[1].split(')')[0]
                if vatis_abr in crc_name or (len(facility_id) > 0 and vatis_abr in facility_id):
                    return data['name']

    config = {}
    try:
        url = 'https://raw.githubusercontent.com/glott/vATISLoad/refs/heads/main/vATISLoadConfig.json'
        config = json.loads(requests.get(url).text)
    except Exception as ignored:
        pass
    
    if 'facility-patches' in config and facility_id in config['facility-patches']:
        patch = config['facility-patches'][facility_id]

        for filename in os.listdir(vatis_profiles):
            if filename.endswith('.json'): 
                file_path = os.path.join(vatis_profiles, filename)
                with open(file_path, 'r') as f:
                    data = json.load(f)
                    data['name'] = data['name'].replace('[', '(').replace(']', ')')
                    if not('(' in  data['name'] and ')' in data['name']):
                        continue
                    vatis_abr = data['name'].split('(')[1].split(')')[0]
                    if vatis_abr in patch:
                        return data['name']
    return ''

def get_active_profile_position(profile):
    vatis_profiles = os.getenv('LOCALAPPDATA') + '\\org.vatsim.vatis\\Profiles'
    profile_names = []
    for filename in os.listdir(vatis_profiles):
        if filename.endswith('.json'): 
            file_path = os.path.join(vatis_profiles, filename)
            with open(file_path, 'r') as f:
                data = json.load(f)
                profile_names.append(data['name'])

    profile_names = sorted(profile_names, key=str.lower)
    if profile in profile_names:
        return profile_names.index(profile)
    return -1

def open_vATIS():
    # Set 'autoFetchAtisLetter' to True
    config_path = os.getenv('LOCALAPPDATA') + '\\org.vatsim.vatis\\AppConfig.json'
    try:
        with open(config_path, 'r') as f:
            data = json.load(f)
            if 'autoFetchAtisLetter' in data:
                data['autoFetchAtisLetter'] = True
        with open(config_path, 'w') as f:
            json.dump(data, f, indent=2)
    except Exception as ignored:
        pass

    # Close vATIS, then reopen vATIS
    os.system('taskkill /f /im vATIS.exe 2>nul 1>nul')
    exe = os.getenv('LOCALAPPDATA') + '\\org.vatsim.vatis\\current\\vATIS.exe'
    subprocess.Popen(exe);

    # Open vATIS dialog
    for i in range(0, 50): 
        for window in pygetwindow.getAllWindows():
            if 'vATIS Profiles' in window.title:
                set_foreground_window(window._hWnd)
                
                t0 = time.time()
                atis_data = get_all_datises()
                dt = time.time() - t0
                if dt < 0.5:
                    time.sleep(0.5 - dt)
                return atis_data
            else:
                time.sleep(0.1)
    return []

async def get_online_atises():
    # url = "https://data.vatsim.net/v3/vatsim-data.json"
    # response = requests.get(url)
    # data = response.json()

    # vat_atises = []
    # for atis in data['atis']:
    #     vat_atises.append(atis['callsign'].replace('_ATIS', ''))
    
    online_atises = {}
    async with websockets.connect('ws://127.0.0.1:49082/', close_timeout=0.05) as websocket:
        for i in range(0, 20):
            await websocket.send(json.dumps({'type': 'getAtis'}))
            m = json.loads(await websocket.recv())['value']
            if m['atisType'] == 'Arrival':
                m['station'] += '_A'
            elif m['atisType'] == 'Departure':
                m['station'] += '_D'

            if m['networkConnectionStatus'] != 'Disconnected' and m['station'] not in online_atises:
                online_atises[m['station']] = m['networkConnectionStatus']
            # elif m['networkConnectionStatus'] == 'Disconnected' and m['station'] in vat_atises:
            #     online_atises[m['station']] = 'Connected'
    
    return online_atises

def read_profile(profile):
    vatis_profiles = os.getenv('LOCALAPPDATA') + '\\org.vatsim.vatis\\Profiles'
    for filename in os.listdir(vatis_profiles):
        if filename.endswith('.json'): 
            file_path = os.path.join(vatis_profiles, filename)
            with open(file_path, 'r') as f:
                data = json.load(f)
                if data['name'] == profile:
                    return data

def get_stations(data):
    stations = []
    ordinal = False
    for station in data['stations']:
        o = 0 if 'ordinal' not in station else station['ordinal']
        if o != 0:
            ordinal = True
        
        s = station['identifier']
        if station['atisType'] == 'Departure':
            s += '_D'
        elif station['atisType'] == 'Arrival':
            s += '_A'
        stations.insert(o, s)

    if not ordinal:
        stations = sorted(stations)

    return stations

def get_atis_replacements(stations):
    stations = list(set(value.replace('_A', '').replace('_D', '') for value in stations))

    config = {}
    try:
        url = 'https://raw.githubusercontent.com/glott/vATISLoad/refs/heads/main/vATISLoadConfig.json'
        config = json.loads(requests.get(url).text)
    except Exception as ignored:
        pass

    if 'replacements' not in config:
        return {}

    replacements = {}
    for a in config['replacements']:
        if a in stations:
            replacements[a] = config['replacements'][a]

    return replacements

def get_contractions(station, data):
    c = {}
    station_data = {}
    for s in data['stations']:
        if station[0:4] in s['identifier']:
            if len(station) == 4:
                station_data = s
                break
            elif (station[5] == 'D' and s['atisType'] == 'Departure') or \
                (station[5] == 'A' and s['atisType'] == 'Arrival'):
                station_data = s
    contractions = station_data['contractions']
    for cont in contractions:
        c[cont['text']] = '@' + cont['variableName']
    c = dict(sorted(c.items(), key=lambda item: len(item[0])))
    c = {key: c[key] for key in reversed(c)}

    return c

async def get_station_position(station, stations):
    online_atises = []
    offline_atises = list(stations)

    for s, v in (await get_online_atises()).items():
        if v == 'Observer' and s == station:
            return -1
        
        if s in stations:
            online_atises.insert(stations.index(s), s)

        if s in offline_atises:
            offline_atises.remove(s)

    left_pad = 20
    for atis in online_atises:
        if station in atis:
            if '_' in atis:
                left_pad += tab_sizes['large_con'] / 2
            else:
                left_pad += tab_sizes['small_con'] / 2
            return left_pad
        else:
            if '_' in atis:
                left_pad += tab_sizes['large_con']
            else:
                left_pad += tab_sizes['small_con']

    for atis in offline_atises:
        if station in atis:
            if '_' in atis:
                left_pad += tab_sizes['large'] / 2
            else:
                left_pad += tab_sizes['small'] / 2
            return left_pad
        else:
            if '_' in atis:
                left_pad += tab_sizes['large']
            else:
                left_pad += tab_sizes['small']
    
    return left_pad

def get_all_datises():
    url = 'https://datis.clowd.io/api/all'
    return json.loads(requests.get(url).text)

def get_datis(ident, atis_data, data, replacements):
    atis_type = 'combined'
    if '_A' in ident:
        atis_type = 'arr'
    elif '_D' in ident:
        atis_type = 'dep'

    atis_info = []
    if 'error' in atis_data:
        return atis_info
    
    for n in range(0, len(atis_data)):
        if atis_data[n]['airport'] != ident[0:4]:
            continue
        if atis_data[n]['type'] != atis_type:
            continue
        
        # Strip beginning and ending D-ATIS text
        datis = atis_data[n]['datis']
        datis = re.sub('.*INFO [A-Z] [0-9][0-9][0-9][0-9]Z. ', '', datis)
        datis = '. '.join(datis.split('. ')[1:])
        datis = re.sub(' ...ADVS YOU HAVE.*', '', datis)
        datis = datis.replace('NOTICE TO AIR MISSIONS, NOTAMS. ', 'NOTAMS... ') \
            .replace('NOTICE TO AIR MISSIONS. ', 'NOTAMS... ')

        # Replace defined replacements
        for r in replacements:
            if '%r' in replacements[r]:
                datis = re.sub(r + '[,.;]{0,2}', replacements[r].replace('%r', ''), datis)
            else:
                datis = re.sub(r + '[,.;]{0,2}', replacements[r], datis)
        datis = re.sub(r'\s+', ' ', datis).strip()

        # Clean up D-ATIS
        datis = datis.replace('...', '/./').replace('..', '.') \
            .replace('/./', '...').replace('  ', ' ').replace(' . ', '. ') \
            .replace(', ,', ',').replace(' ; ', '; ').replace(' .,', ' ,') \
            .replace(' , ', ', ').replace('., ', ', ').replace('&amp;', '&') \
            .replace(' ;.', '.').replace(' ;,', ',')

        # Replace contractions
        contractions = get_contractions(ident, data)
        for c, v in contractions.items():
            if not c.isdigit():
                datis = re.sub(r'(?<!@)\b' + c + r'\b,', v + ',', datis)
                datis = re.sub(r'(?<!@)\b' + c + r'\b\.', v + '.', datis)
                datis = re.sub(r'(?<!@)\b' + c + r'\b ', v + ' ', datis)
                datis = re.sub(r'(?<!@)\b' + c + r'\b;', v + ';', datis)

        # Split at NOTAMs
        if 'NOTAMS' in datis:
            atis_info = datis.split('NOTAMS... ')
        else:
            atis_info = [datis, '']
    
    return atis_info

async def load_atis(station, stations, data, atis_data, atis_replacements):
    left_pad = await get_station_position(station, stations)
    
    if left_pad == -1:
        return 0
    
    station_data = {}
    for elem in data['stations']:
        if station[0:4] in elem['identifier']:
            if len(station) == 4:
                station_data = elem
                break
            elif (station[5] == 'D' and elem['atisType'] == 'Departure') or \
                (station[5] == 'A' and elem['atisType'] == 'Arrival'):
                station_data = elem
                break

    # Make sure profile has D-ATIS preset
    if station_data['presets'][0]['name'] != 'D-ATIS':
        return 0

    
    # Select profile
    click_xy([left_pad, 100], win)

    # Select D-ATIS preset
    click_xy([500, 500], win, d=0.1)
    click_xy([500, 550], win, d=0.1)

    # Add ATIS replacements
    replacements = {}
    if station[0:4] in atis_replacements:
        replacements = atis_replacements[station[0:4]]
    
    atis = get_datis(station, atis_data, data, replacements)
    if len(atis) == 0:
        return 0
    # Fill in AIRPORT CONDITIONS field
    click_xy([335, 390], win)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')
    pyperclip.copy(atis[0])
    pyautogui.hotkey('ctrl', 'v')
    click_xy([540, 295], win, d=0.1)

    # Fill in NOTAMS field
    click_xy([940, 390], win)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')
    pyperclip.copy(atis[1])
    pyautogui.hotkey('ctrl', 'v')
    pyperclip.copy('')
    click_xy([1145, 295], win, d=0.1)

    # This code likely won't work until there's a reliable way to check if something is online
    # Select profile again and connect ATIS
    
    # click_xy([left_pad, 100], win)
    # with pyautogui.hold('shift'):
    #     pyautogui.press('tab', presses=7)
    # pyautogui.press('enter')
    # time.sleep(1)

    print(f'{station} is now loaded.')
    time.sleep(0.05)
    
    return 0

if RUN_UPDATE:
    update_vATISLoad()

active_profile = determine_active_profile()
if len(active_profile) == 0:
    print('Active profile not found.')
    time.sleep(10)
    sys.exit()

# Open vATIS and select first profile
atis_data = open_vATIS()
pyautogui.PAUSE = 0.0001
win = get_win('vATIS.exe', 'vATIS Profiles')
if True:
    # Suppress mouse movements
    mouse_listener = pynput.mouse.Listener()
    def on_move(x,y):
        mouse_listener.suppress_event()
    mouse_listener.on_move = on_move
    mouse_listener.start()
    
    click_xy([0, 0], win)
    pyautogui.press('tab')

    # Select active profile
    for i in range(0, get_active_profile_position(active_profile)):
        pyautogui.press('down')
    pyautogui.press('enter')
    with pyautogui.hold('shift'):
        pyautogui.press('tab', presses=5)
    pyautogui.press('enter')
    
    mouse_listener.stop()

data = read_profile(active_profile)
stations = get_stations(data)

# Bring window to front
win = get_win('vATIS.exe', 'vATIS')

# Get ATIS replacements
t0 = time.time()
atis_replacements = get_atis_replacements(stations)
dt = time.time() - t0
if dt < 1.0:
    time.sleep(1.0 - dt)

# Load ATIS information
i = 0
for station in stations[::-1]:
    if i > 3: 
        break

    # Suppress mouse movements
    mouse_listener = pynput.mouse.Listener()
    def on_move(x,y):
        mouse_listener.suppress_event()
    mouse_listener.on_move = on_move
    mouse_listener.start()

    # Use first line for Desktop, second line for Jupyter
    i += asyncio.run(load_atis(station, stations, data, atis_data, atis_replacements))
    # i += await load_atis(station, stations, data, atis_data, atis_replacements)
    
    mouse_listener.stop()

# Restore these items in the future
# time.sleep(3)
# win32gui.ShowWindow(win._hWnd, win32con.SW_MINIMIZE);