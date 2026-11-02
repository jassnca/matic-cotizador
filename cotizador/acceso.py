"""Acceso con clave para el uso local compartido. La publicación utiliza Caddy."""
from pathlib import Path
import hashlib,hmac,json,time,os
import streamlit as st

CONFIG=Path(__file__).resolve().parent/'acceso_local.json'

def verificar_acceso():
    if os.environ.get('MATIC_CLOUD_MODE')=='1':
        try: cfg=dict(st.secrets['acceso'])
        except Exception:
            st.error('Publicación protegida: falta configurar la clave en Secrets.'); st.stop()
        if not cfg.get('salt') or not cfg.get('hash'):
            st.error('Configuración de acceso incompleta.'); st.stop()
    else:
        if not CONFIG.exists(): return
        cfg=json.loads(CONFIG.read_text(encoding='utf-8'))
    try:
        salt=bytes.fromhex(cfg['salt'])
        if len(salt)<16 or len(bytes.fromhex(cfg['hash']))!=32: raise ValueError('Formato inválido')
    except (ValueError,TypeError,KeyError):
        st.error('Configuración de acceso no válida.'); st.stop()
    if st.session_state.get('_acceso_hash')==cfg['hash']: return
    st.title('MATIC · Acceso al cotizador')
    st.caption('Introduce la clave del cotizador.' if os.environ.get('MATIC_CLOUD_MODE')=='1' else 'Introduce la clave del acceso compartido. Utiliza una red Wi-Fi de confianza.')
    until=st.session_state.get('_acceso_bloqueo',0)
    if until>time.time():
        st.warning('Espera un minuto antes de volver a intentar.'); st.stop()
    with st.form('acceso_compartido'):
        password=st.text_input('Clave de acceso',type='password')
        submit=st.form_submit_button('Ingresar')
    if submit:
        digest=hashlib.pbkdf2_hmac('sha256',password.encode(),salt,300000).hex()
        if hmac.compare_digest(digest,cfg['hash']):
            st.session_state['_acceso_hash']=cfg['hash']; st.session_state['_acceso_intentos']=0; st.rerun()
        else:
            attempts=st.session_state.get('_acceso_intentos',0)+1
            st.session_state['_acceso_intentos']=attempts
            if attempts>=5:
                st.session_state['_acceso_bloqueo']=time.time()+60; st.session_state['_acceso_intentos']=0
            st.error('Clave incorrecta.')
    st.stop()
