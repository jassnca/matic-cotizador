"""Entrada para Streamlit Community Cloud; acceso cerrado sin secretos configurados."""
import os,sys,runpy
from pathlib import Path
os.environ['MATIC_CLOUD_MODE']='1'
app_dir=Path(__file__).resolve().parent/'cotizador'
sys.path.insert(0,str(app_dir))
runpy.run_path(str(app_dir/'app.py'),run_name='__main__')
