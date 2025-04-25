import subprocess, sys, os, time

import importlib.util as il
if None in [il.find_spec('requests'), il.find_spec('websockets'), il.find_spec('psutil'), 
            il.find_spec('pygetwindow')]:

    os.system('cls')
    print('Installing required packages for vATISLoad!')
    time.sleep(2.5)
    
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', 'pip'])
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'requests']);
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'websockets']);
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'psutil']);
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pygetwindow']);

    os.system('cls')
    print('vATISLoad required packages successfully installed!')
    time.sleep(5)