#####################################################################
############################# vATISLoad #############################
#####################################################################

DISABLE_AUTOCONNECT = False     # Set to True to disable auto-connect
DISABLE_AUTOUPDATES = False     # Set to True to disable auto-updates
RUN_UPDATE = True               # Set to False for testing
SHUTDOWN_LIMIT = 60 * 5         # Time delay to exit script

#####################################################################

import subprocess, sys, os, time, json, re, uuid, ctypes, asyncio
from datetime import datetime

import importlib.util as il
if None in [il.find_spec('requests'), il.find_spec('websockets'), il.find_spec('psutil')]:

    os.system('cmd /K \"cls & echo Updating required packages for vATISLoad.' + \
        ' & echo Please wait a few minutes for packages to install. & timeout 5 & exit\"')
    
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', 'pip'])
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'requests']);
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'websockets']);
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'psutil']);

os.system('cls')

import requests, websockets, psutil

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
            if ('DISABLE_AUTOCONNECT =' in line or 'DISABLE_AUTOUPDATES =' in line or 
                'RUN_UPDATE =' in line or 'SHUTDOWN_LIMIT =' in line) and i < 10:
                pass
            elif i > len(online_file) or len(line.strip()) != len(online_file[i].strip()):
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

async def try_websocket(shutdown=RUN_UPDATE, limit=SHUTDOWN_LIMIT):
    t0 = time.time()
    for i in range(0, 250):
        t1 = time.time()
        if t1 - t0 > limit:
            os.system('cmd /K \"cls & vATIS not open and/or a profile is not loaded, exiting vATISLoad.' + 
                      ' & timeout 5 & exit\"')
            if shutdown:
                sys.exit()
            return
        try:
            async with websockets.connect('ws://127.0.0.1:49082/', close_timeout=0.01) as websocket:
                await websocket.send(json.dumps({'type': 'getStations'}))
                try:
                    await asyncio.wait_for(websocket.recv(), timeout=1)
                    if time.time() - t0 > 5:
                        time.sleep(1)
                    return
                except Exception as ignored:
                    pass
        except Exception as ignored:
            dt = time.time() - t1
            if dt < 1:
                time.sleep(1 - dt)
            pass

async def get_datis_stations():
    await try_websocket()
    
    data = {}
    async with websockets.connect('ws://127.0.0.1:49082/', close_timeout=0.01) as websocket:
        await websocket.send(json.dumps({'type': 'getStations'}))
        stations = json.loads(await websocket.recv())['stations']

        for s in stations:
            name = s['name']

            if s['atisType'] == 'Arrival':
                name += '_A'
            elif s['atisType'] == 'Departure':
                name += '_D'
            
            if 'D-ATIS' in s['presets']:
                data[name] = s['id']
            
    return data

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

async def get_current_profile_data():
    stations = list((await get_datis_stations()).keys())
    stations = list(set(value.replace('_A', '').replace('_D', '') for value in stations))
    stations.sort()
    
    vatis_profiles = os.getenv('LOCALAPPDATA') + '\\org.vatsim.vatis\\Profiles'
    data = {}
    for filename in os.listdir(vatis_profiles):
        if not filename.endswith('.json'): 
            continue
        
        file_path = os.path.join(vatis_profiles, filename)
        with open(file_path, 'r') as f:
            data = json.load(f)['stations']
        
        profile_datis_stations = []
        for s in data:
            if s['identifier'] in profile_datis_stations:
                continue
            
            for p in s['presets']:
                if p['name'] == 'D-ATIS':
                    profile_datis_stations.append(s['identifier'])
                    break

        profile_datis_stations.sort()
        if profile_datis_stations == stations:
            return data

    return data
    
async def get_contractions(station):
    data = await get_current_profile_data()

    c = {}
    station_data = {}
    for s in data:
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

def get_datis_data():
    data = {}
    try:
        url = 'https://datis.clowd.io/api/all'
        data =  json.loads(requests.get(url, timeout=2.5).text)
    except Exception as ignored:
        os.system('cmd /K \"cls & echo Unable to fetch D-ATIS data. & timeout 5 & exit\"')
    
    return data

async def get_datis(station, atis_data, replacements):
    atis_type = 'combined'
    if '_A' in station:
        atis_type = 'arr'
    elif '_D' in station:
        atis_type = 'dep'

    atis_info = ['D-ATIS NOT AVBL', '']
    if 'error' in atis_data:
        return atis_info

    datis = ''
    for a in atis_data:
        if a['airport'] != station[0:4] or a['type'] != atis_type:
            continue
        datis = a['datis']

    if len(datis) == 0:
        return atis_info

    # Strip beginning and ending D-ATIS text
    datis = re.sub('.*INFO [A-Z] [0-9][0-9][0-9][0-9]Z. ', '', datis)
    datis = '. '.join(datis.split('. ')[1:])
    datis = re.sub(' ...ADVS YOU HAVE.*', '', datis)
    datis = datis.replace('NOTICE TO AIR MISSIONS, NOTAMS. ', 'NOTAMS... ') \
        .replace('NOTICE TO AIR MISSIONS. ', 'NOTAMS... ') \
        .replace('NOTICE TO AIR MEN. ', 'NOTAMS... ') \
        .replace('NOTICE TO AIRMEN. ', 'NOTAMS... ')

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
    contractions = await get_contractions(station)
    for c, v in contractions.items():
        if not c.isdigit():
            datis = re.sub(r'(?<!@)\b' + c + r'\b,', v + ',', datis)
            datis = re.sub(r'(?<!@)\b' + c + r'\b\.', v + '.', datis)
            datis = re.sub(r'(?<!@)\b' + c + r'\b ', v + ' ', datis)
            datis = re.sub(r'(?<!@)\b' + c + r'\b;', v + ';', datis)

    # Split at NOTAMs
    if 'NOTAMS... ' in datis:
        atis_info = datis.split('NOTAMS... ')
    else:
        atis_info = [datis, '']
    
    return atis_info

async def get_atis_statuses():
    await try_websocket()
    
    data = {}
    
    async with websockets.connect('ws://127.0.0.1:49082/', close_timeout=0.01) as websocket:
        for s, i in (await get_datis_stations()).items():
            await websocket.send(json.dumps({'type': 'getAtis', 'value': {'id': i}}))
            
            m = json.loads(await websocket.recv())['value']
            data[s] = m['networkConnectionStatus']
    
    return data

async def get_num_connections():
    n = 0
    for k, v in (await get_atis_statuses()).items():
        if v == 'Connected':
            n =+ 1
    return n

async def configure_atises(connected_only=False):
    stations = await get_datis_stations()
    replacements = get_atis_replacements(stations)
    atis_data = get_datis_data()

    atis_statuses = await get_atis_statuses()
    
    for s, i in stations.items():
        if connected_only and atis_statuses[s] != 'Connected':
            continue

        rep = []
        if s[0:4] in replacements:
            rep = replacements[s[0:4]]
        
        v = {'id': i, 'preset': 'D-ATIS'}
        v['airportConditionsFreeText'], v['notamsFreeText'] = await get_datis(s, atis_data, rep)
        
        payload = {'type': 'configureAtis', 'value': v}
        async with websockets.connect('ws://127.0.0.1:49082/', close_timeout=0.01) as websocket:
            await websocket.send(json.dumps(payload))

async def connect_atises():
    stations = await get_datis_stations()
    atis_statuses = await get_atis_statuses()
    disconnected_atises = [k for k, v in atis_statuses.items() if v == 'Disconnected']
    
    for s, i in stations.items():
        if s not in disconnected_atises:
            continue
        
        payload = {'type': 'connectAtis', 'value': {'id': i}}
        async with websockets.connect('ws://127.0.0.1:49082/', close_timeout=0.01) as websocket:
            await websocket.send(json.dumps(payload))

            try:
                m = await asyncio.wait_for(websocket.recv(), timeout=0.1)
            except Exception as ignored:
                pass

def kill_open_instances():
    prev_instances = {}

    for q in psutil.process_iter():
        if 'python' in q.name():
            for parameter in q.cmdline():
                if 'vATISLoad' in parameter and parameter.endswith('.pyw'):
                    q_create_time = q.create_time()
                    q_create_datetime = datetime.fromtimestamp(q_create_time)
                    prev_instances[q.pid] = {'process': q, 'start': q_create_datetime}
    
    prev_instances = dict(sorted(prev_instances.items(), key=lambda item: item[1]['start']))
    
    for i in range(0, len(prev_instances) - 1):
        k = list(prev_instances.keys())[i]
        prev_instances[k]['process'].terminate()

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

    # Check if vATIS is open
    for process in psutil.process_iter(['name']):
        if process.info['name'] == 'vATIS.exe':
            return

    exe = os.getenv('LOCALAPPDATA') + '\\org.vatsim.vatis\\current\\vATIS.exe'
    subprocess.Popen(exe);

async def main():
    if RUN_UPDATE:
        update_vATISLoad()

    kill_open_instances()
    open_vATIS()
    await configure_atises()
    if not DISABLE_AUTOCONNECT:
        await connect_atises()

    while not DISABLE_AUTOUPDATE:
        for i in range(0, 15):
            await try_websocket()
            time.sleep(60)
            
        await configure_atises(connected_only=True)

if __name__ == "__main__":
    # Use first line for Desktop, second line for Jupyter
    asyncio.run(main())
    # await main()