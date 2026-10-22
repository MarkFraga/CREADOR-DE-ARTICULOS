import sys
import os
import shutil
import subprocess

# Ficheros de datos que se empaquetan junto al exe (solo los que existan).
DATA_FILES = [
    'ayuda.xlsx', 'PROVS.xlsx', 'exceltodo.xlsx', 'ROBOT.py',
    'en_proceso.gif', 'completado.png', 'don_articulo.ico', 'toilet_logo.png', 'logo.png',
]

os.chdir(os.path.dirname(os.path.abspath(__file__)))

print("Building DON_ARTICULO.exe directly into root folder...")

command = [
    sys.executable, '-m', 'PyInstaller',
    'DON_ARTICULO.py',
    '--noconfirm',
    '--clean',
    '--name=DON ARTICULO',
]
command += ['--add-data=%s;.' % name for name in DATA_FILES if os.path.exists(name)]
if os.path.exists('don_articulo.ico'):
    command.append('--icon=don_articulo.ico')
command += ['--distpath=.', '--workpath=build_temp']

subprocess.run(command, check=True)

print("Cleaning up PyInstaller leftover files...")
if os.path.exists('DON ARTICULO.spec'):
    os.remove('DON ARTICULO.spec')

dest_internal = os.path.join(os.getcwd(), '_internal')
src_internal = os.path.join(os.getcwd(), 'DON ARTICULO', '_internal')

if os.path.exists(dest_internal):
    shutil.rmtree(dest_internal)
if os.path.exists(src_internal):
    shutil.move(src_internal, dest_internal)

exe_src = os.path.join(os.getcwd(), 'DON ARTICULO', 'DON ARTICULO.exe')
exe_dest = os.path.join(os.getcwd(), 'DON ARTICULO.exe')
if os.path.exists(exe_dest):
    os.remove(exe_dest)
if os.path.exists(exe_src):
    shutil.move(exe_src, exe_dest)

src_folder = os.path.join(os.getcwd(), 'DON ARTICULO')
if os.path.exists(src_folder) and not os.listdir(src_folder):
    os.rmdir(src_folder)

if os.path.exists('build_temp'):
    shutil.rmtree('build_temp')

print("DONE.")
