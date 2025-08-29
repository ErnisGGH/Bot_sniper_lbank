# lbank_gui_sniper.py
# Bot LBank Sniper con Interfaz Gráfica de Usuario (GUI)
# Versión con:
# - Endpoint de creación de orden a /v2/supplement/create_order.do
# - Uso del parámetro 'price' para el monto USDT en órdenes 'buy_market'
# - Mantenimiento de firma HmacSHA256 (MD5 -> HMAC -> Base64)
# - Uso de 'timestamp' como nombre del parámetro de marca de tiempo.
# - Ajuste en la generación de 'echostr' a 32 caracteres.
# - EXPERIMENTAL v2: 'signature_version' se envía pero NO se incluye en la base de la firma MD5.

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, scrolledtext
import threading
import queue
import time as pytime 
import yaml
import os
import requests
import hashlib
import random
import string 
import base64
import hmac
import pandas as pd
import smtplib
from email.mime.text import MIMEText
import logging
import traceback
import webbrowser 
import datetime 

# --- Configuración del Logging ---
console_logger = logging.getLogger("ConsoleLogger")
console_logger.setLevel(logging.DEBUG) 
if not console_logger.hasHandlers(): 
    console_handler = logging.StreamHandler()
    console_formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', datefmt='%H:%M:%S')
    console_handler.setFormatter(console_formatter)
    console_logger.addHandler(console_handler)
console_logger.info("Logging de consola inicializado.")

# --- Variables Globales ---
bot_thread = None
stop_bot_event = threading.Event()
log_queue = queue.Queue()
CONFIG_FILE = "config_gui.yml"
EXCEL_LOG_FILE = "log_compras_gui.xlsx"
VERSION = "1.2.7" # Versión del Bot con ajuste experimental en firma v2 (signature_version)

# --- Lógica del Bot ---

def generate_random_string_for_echostr(length=35):
    return ''.join(random.choice(string.ascii_letters + string.digits) for _ in range(length))

def generate_lbank_signature_for_api(params_dict_for_md5_base, secret_key_str, declared_signature_method="HmacSHA256", gui_log_func=None):
    if gui_log_func is None: gui_log_func = console_logger.debug 

    sorted_params_list = sorted(params_dict_for_md5_base.items())
    string_for_md5 = '&'.join([f"{k}={v}" for k, v in sorted_params_list])
    gui_log_func(f"DEBUG: String para MD5 (base para preparedStr): {string_for_md5}")

    md5_prepared_str = hashlib.md5(string_for_md5.encode('utf-8')).hexdigest().upper()
    gui_log_func(f"DEBUG: MD5 Prepared String (preparedStr): {md5_prepared_str}")

    if declared_signature_method.upper() == "HMACSHA256":
        try:
            final_sign_value = hmac.new(secret_key_str.encode('utf-8'), md5_prepared_str.encode('utf-8'), hashlib.sha256).hexdigest().lower()
            gui_log_func(f"DEBUG: HmacSHA256 Signature (Hex encoded): {final_sign_value}")
            return final_sign_value
        except Exception as e:
            gui_log_func(f"ERROR: Excepción al generar firma HmacSHA256: {e}", "error")
            return None
    else:
        gui_log_func(f"ERROR: Método de firma '{declared_signature_method}' no soportado.", "error")
        return None

# #######################################################################################
# ## FUNCIÓN make_lbank_api_request_gui MODIFICADA (signature_version fuera de base MD5) ##
# #######################################################################################
def make_lbank_api_request_gui(endpoint_path_str, params_dict_original, http_method, api_config, gui_log_func, requires_full_signature_flow=False):
    api_key_str = api_config.get('lbank_api_key')
    secret_key_str = api_config.get('lbank_secret_key')
    base_url_str = "https://api.lbank.info" 

    if not api_key_str or not secret_key_str:
        gui_log_func("ERROR: Claves API no disponibles para la solicitud.", "error")
        return None

    # Parámetros que SÍ se incluirán en la cadena base para el MD5 y la firma HmacSHA256
    params_for_signature_generation = params_dict_original.copy() # Endpoint-specific params (symbol, type, price)
    params_for_signature_generation['api_key'] = api_key_str
    
    # Initialize headers
    headers = {'Content-Type': 'application/json', 'User-Agent': f"LBankSniperBotGUI/{VERSION}"}
    
    # params_to_send_final will be the JSON body for POST or query params for GET.
    # It's initialized with business parameters.
    params_to_send_final = params_dict_original.copy()

    if requires_full_signature_flow:
        current_timestamp_val = str(int(pytime.time() * 1000)) 
        current_echostr = generate_random_string_for_echostr() 
        declared_final_method = "HmacSHA256"
        signature_version_value = "2.0"

        # Populate params for signature generation (includes business params, api_key, and sig-specific params)
        params_for_signature_generation['timestamp'] = current_timestamp_val 
        params_for_signature_generation['echostr'] = current_echostr
        params_for_signature_generation['signature_method'] = declared_final_method
        params_for_signature_generation['signature_version'] = signature_version_value
        
        gui_log_func(f"DEBUG: Params for signature base (contents before signing): {params_for_signature_generation}")
        gui_log_func(f"DEBUG: Params for signature base (sorted items before signing): {sorted(params_for_signature_generation.items())}")
        sign = generate_lbank_signature_for_api(params_for_signature_generation, secret_key_str, declared_signature_method=declared_final_method, gui_log_func=gui_log_func)
        if not sign:
            gui_log_func("ERROR: Fallo al generar la firma HmacSHA256.", "error")
            return None
        
        # Populate headers with ALL signature-related parameters
        headers['api_key'] = api_key_str
        headers['signature_version'] = signature_version_value
        headers['sign'] = sign
        headers['timestamp'] = current_timestamp_val
        headers['echostr'] = current_echostr
        headers['signature_method'] = declared_final_method
        
        # For POST requests with full signature flow, the JSON body (params_to_send_final)
        # must ONLY contain the original business parameters.
        # We ensure this by re-assigning it to a fresh copy of params_dict_original here.
        if http_method.upper() == "POST":
            params_to_send_final = params_dict_original.copy()
        # For GET requests, params_to_send_final already holds params_dict_original.
        # All signature-related parameters are now in headers.
        # If the API for some GET endpoints expects these in query params *as well*,
        # they would need to be explicitly added to params_to_send_final here.
        # Assuming they are only needed in headers for GET if requires_full_signature_flow.

    else: # Not requires_full_signature_flow
        # For endpoints not requiring full signature, params_to_send_final is params_dict_original.
        # Specific non-signed endpoints might have different expectations for api_key (e.g., some might need it in query/body).
        # /v2/ticker/24hr.do typically doesn't need api_key.
        if endpoint_path_str == "/v2/ticker/24hr.do": 
            pass # params_to_send_final is already params_dict_original
        else:
            # If other non-signed endpoints might need api_key in the query/body, it should be added here.
            # For now, assuming it's not needed for other non-signed general calls by default.
            gui_log_func(f"ADVERTENCIA: Endpoint {endpoint_path_str} llamado sin flujo de firma HmacSHA256 completo. Params: {params_to_send_final}", "warning")
            # Consider if a generic non-signed request should attempt to include api_key in params_to_send_final or not.
            # The original return for "UNSUPPORTED_SIGN_FLOW_IN_BOT" was removed in a previous step to allow flexibility.

    full_url_str = base_url_str + endpoint_path_str
    
    try:
        if http_method.upper() == "POST":
            # params_to_send_final is now clean for POST if requires_full_signature_flow was true.
            gui_log_func(f"DEBUG: Enviando {http_method} a {full_url_str} con JSON body: {params_to_send_final} y Headers: {headers}")
            response = requests.post(full_url_str, json=params_to_send_final, headers=headers, timeout=20)
        elif http_method.upper() == "GET": 
            # For GET, params_to_send_final contains original business params.
            # Headers contain sig elements if requires_full_signature_flow.
            gui_log_func(f"DEBUG: Enviando {http_method} a {full_url_str} con params: {params_to_send_final} y Headers: {headers}")
            response = requests.get(full_url_str, params=params_to_send_final, headers=headers, timeout=20)
        else:
            gui_log_func(f"ERROR: Método HTTP no soportado: {http_method}", "error")
            return None
            
        response_json = response.json()
        gui_log_func(f"DEBUG: Respuesta JSON de LBank ({endpoint_path_str}): {response_json}")
        
        if isinstance(response_json, dict) and str(response_json.get('result')).lower() == 'false':
            error_code = response_json.get('error_code')
            gui_log_func(f"ERROR: Error API LBank (código {error_code}): {response_json.get('msg', 'Sin mensaje de error específico')}", "error")
            return response_json 
        
        if response.status_code != 200 and not (isinstance(response_json, dict) and str(response_json.get('result')).lower() == 'false'):
             gui_log_func(f"ERROR: Respuesta HTTP no exitosa ({response.status_code}) de LBank: {response.text}", "error")
             return None 

        return response_json
        
    except requests.exceptions.Timeout:
        gui_log_func(f"ERROR: Timeout en la solicitud a LBank ({endpoint_path_str}).", "error")
    except requests.exceptions.RequestException as e:
        gui_log_func(f"ERROR: Error en la solicitud a LBank ({endpoint_path_str}): {e}", "error")
    except ValueError as e: 
        gui_log_func(f"ERROR: Error decodificando JSON de LBank ({endpoint_path_str}): {e}. Respuesta: {response.text if 'response' in locals() else 'N/A'}", "error")
    return None
# #######################################################################################
# ## FIN DE LA FUNCIÓN MODIFICADA ##
# #######################################################################################

def check_if_symbol_is_tradable_gui(target_symbol_str_from_input, gui_log_func):
    # ... (Esta función no ha cambiado) ...
    formatted_api_symbol = target_symbol_str_from_input.lower().strip() + "_usdt"
    base_url = "https://api.lbank.info"
    ticker_endpoint = "/v2/ticker/24hr.do" 
    full_url = base_url + ticker_endpoint
    params = {'symbol': formatted_api_symbol}
    gui_log_func(f"INFO: Verificando símbolo '{formatted_api_symbol}' en LBank...")
    try:
        response = requests.get(full_url, params=params, timeout=10)
        try:
            data_from_api = response.json()
            gui_log_func(f"DEBUG: Respuesta de {ticker_endpoint} para '{formatted_api_symbol}': {data_from_api}")
        except ValueError:
            gui_log_func(f"ERROR: Error decodificando JSON de LBank al verificar '{formatted_api_symbol}'. Status: {response.status_code}, Respuesta: {response.text[:200]}", "error")
            return None 
        if response.status_code != 200:
            gui_log_func(f"INFO: Símbolo '{formatted_api_symbol}' no encontrado o error de servidor (HTTP Status: {response.status_code}). Respuesta: {data_from_api if isinstance(data_from_api, dict) else response.text[:100]}", "warning")
            return False
        condition1_met = (
            isinstance(data_from_api, dict) and
            str(data_from_api.get('result')).lower() == 'true' and 
            'data' in data_from_api and
            isinstance(data_from_api['data'], list) and
            len(data_from_api['data']) > 0
        )
        gui_log_func(f"[DEBUG_DETAIL] Condición 1 (dict con data lista): {condition1_met}")
        if condition1_met:
            ticker_list_from_data_key = data_from_api['data']
            for item_dict in ticker_list_from_data_key:
                if isinstance(item_dict, dict) and 'symbol' in item_dict:
                    api_returned_symbol = item_dict.get('symbol', '').lower()
                    gui_log_func(f"[DEBUG_DETAIL] Comparando (lista 'data'): '{api_returned_symbol}' == '{formatted_api_symbol}'")
                    if api_returned_symbol == formatted_api_symbol:
                        gui_log_func(f"INFO: Símbolo '{formatted_api_symbol}' ENCONTRADO en LBank (en lista bajo clave 'data'). Ticker: {item_dict.get('ticker')}", "success")
                        return True
            gui_log_func(f"INFO: Símbolo '{formatted_api_symbol}' no encontrado específicamente en la lista bajo la clave 'data'.", "info")
            return False 
        condition2_met = isinstance(data_from_api, list) and len(data_from_api) > 0
        gui_log_func(f"[DEBUG_DETAIL] Condición 2 (respuesta es lista directa): {condition2_met}")
        if condition2_met:
            for item_dict in data_from_api:
                if isinstance(item_dict, dict) and 'symbol' in item_dict:
                    api_returned_symbol = item_dict.get('symbol', '').lower()
                    gui_log_func(f"[DEBUG_DETAIL] Comparando (lista directa): '{api_returned_symbol}' == '{formatted_api_symbol}'")
                    if api_returned_symbol == formatted_api_symbol:
                        gui_log_func(f"INFO: Símbolo '{formatted_api_symbol}' ENCONTRADO en LBank (respuesta directa en lista). Ticker: {item_dict.get('ticker')}", "success")
                        return True
            gui_log_func(f"INFO: Símbolo '{formatted_api_symbol}' no encontrado específicamente en la lista de tickers devuelta directamente.", "info")
            return False 
        condition3_met = isinstance(data_from_api, dict) and str(data_from_api.get('result')).lower() == 'false'
        gui_log_func(f"[DEBUG_DETAIL] Condición 3 (dict con result:false): {condition3_met}")
        if condition3_met:
            error_code = data_from_api.get('error_code')
            gui_log_func(f"INFO: Símbolo '{formatted_api_symbol}' no encontrado (API LBank devolvió error {error_code}: {data_from_api.get('msg')}).", "warning")
            return False
        condition4_met = isinstance(data_from_api, dict) and 'symbol' in data_from_api
        gui_log_func(f"[DEBUG_DETAIL] Condición 4 (dict con símbolo nivel superior): {condition4_met}")
        if condition4_met:
            api_returned_symbol = data_from_api.get('symbol', '').lower()
            gui_log_func(f"[DEBUG_DETAIL] Comparando (dict nivel sup.): '{api_returned_symbol}' == '{formatted_api_symbol}'")
            if api_returned_symbol == formatted_api_symbol:
                gui_log_func(f"INFO: Símbolo '{formatted_api_symbol}' ENCONTRADO en LBank (respuesta dict único de nivel superior). Ticker: {data_from_api.get('ticker')}", "success")
                return True
            else: 
                gui_log_func(f"INFO: Símbolo '{formatted_api_symbol}' no encontrado. Respuesta dict de nivel superior con símbolo diferente: '{data_from_api.get('symbol')}'", "info")
                return False
        gui_log_func(f"INFO: Símbolo '{formatted_api_symbol}' no encontrado o formato de respuesta inesperado. Revisar logs DEBUG_DETAIL.", "warning")
        return False
    except requests.exceptions.RequestException as e:
        gui_log_func(f"ERROR: Error de red/solicitud al verificar símbolo '{formatted_api_symbol}': {e}", "error")
        return None 
    except Exception as e: 
        gui_log_func(f"ERROR: Error inesperado al verificar el símbolo '{formatted_api_symbol}' en LBank: {e}", "error")
        gui_log_func(traceback.format_exc(), "debug") 
        return None 

def place_lbank_market_buy_order_gui(target_symbol_str_from_input, usdt_amount_float, api_config, gui_log_func, excel_log_func):
    # ... (Esta función no ha cambiado, usa make_lbank_api_request_gui) ...
    formatted_symbol = target_symbol_str_from_input.lower().strip() + "_usdt"
    gui_log_func(f"INFO: Intentando colocar orden de COMPRA A MERCADO para {formatted_symbol} con {usdt_amount_float} USDT.")
    params_for_order = {
        'symbol': formatted_symbol,
        'type': 'buy_market',
        'price': str(usdt_amount_float), 
    }
    gui_log_func(f"DEBUG: Parámetros base para la orden (antes de añadir params de firma): {params_for_order}")
    endpoint_str = "/v2/supplement/create_order.do" 
    response_data = make_lbank_api_request_gui(endpoint_str, params_for_order, "POST", api_config, gui_log_func, requires_full_signature_flow=True)
    current_time = datetime.datetime.now()
    if response_data and str(response_data.get('result')).lower() == 'true' and response_data.get('order_id'):
        order_id_str = response_data.get('order_id')
        success_message = (f"ÉXITO: ¡COMPRA EXITOSA (ORDEN COLOCADA)! Símbolo: {formatted_symbol}, "
                           f"ID Orden: {order_id_str}, Monto (USDT en 'price'): {usdt_amount_float}.\n"
                           f"IMPORTANTE: Verifica detalles en LBank.")
        gui_log_func(success_message, "success")
        excel_log_func(
            symbol=formatted_symbol, timestamp=current_time, amount_bought_token="VERIFICAR_EN_LBANK",
            price_usdt="VERIFICAR_EN_LBANK", total_usdt_spent=usdt_amount_float, 
            order_id=order_id_str, status="ÉXITO (ORDEN COLOCADA)"
        )
        return {"status": "success", "order_id": order_id_str}
    else:
        error_code_from_api = response_data.get('error_code', 'Desconocido') if isinstance(response_data, dict) else 'Sin respuesta'
        failure_message = (f"ERROR: FALLO AL COLOCAR ORDEN para {formatted_symbol}. "
                           f"Cód API: {error_code_from_api}. Respuesta: {response_data}")
        gui_log_func(failure_message, "error")
        excel_log_func(
            symbol=formatted_symbol, timestamp=current_time, amount_bought_token="N/A",
            price_usdt="N/A", total_usdt_spent=usdt_amount_float, 
            order_id="-", status=f"FALLO (Cód API: {error_code_from_api})"
        )
        return None

def bot_logic_thread_gui(app_gui_instance, current_config):
    # ... (Lógica del hilo principal sin cambios, llama a las funciones actualizadas) ...
    def gui_log_from_thread(message, level="info"):
        log_queue.put((message, level))
    def excel_log_from_thread(symbol, timestamp, amount_bought_token, price_usdt, total_usdt_spent, order_id, status):
        log_queue.put(((symbol, timestamp, amount_bought_token, price_usdt, total_usdt_spent, order_id, status), "excel_data"))
    gui_log_from_thread("INFO: Hilo del bot iniciado.")
    api_cfg = {
        'lbank_api_key': current_config.get('lbank_api_key'),
        'lbank_secret_key': current_config.get('lbank_secret_key')
    }
    target_symbol_input = current_config.get('target_symbol_input') 
    usdt_to_invest = current_config.get('usdt_amount_to_invest')
    poll_interval = current_config.get('poll_interval_seconds')
    stop_after_attempt = current_config.get('stop_after_attempt', True) 

    gui_log_from_thread(f"INFO: Bot configurado para {target_symbol_input}, invirtiendo {usdt_to_invest} USDT.")
    gui_log_from_thread(f"INFO: Opción 'Detener tras 1er intento': {stop_after_attempt}")

    tg_cfg = current_config.get('telegram_config', {})
    email_cfg = current_config.get('email_config', {})
    start_msg = (f"🤖 LBank Sniper Bot (GUI) INICIADO.\n"
                 f"🎯 Símbolo: {target_symbol_input}\n"
                 f"💰 Monto: {usdt_to_invest} USDT\n"
                 f"🕒 Intervalo: {poll_interval} seg.")
    if tg_cfg.get('enabled'):
        app_gui_instance.send_telegram_notification_from_gui(start_msg, tg_cfg) 
    if email_cfg.get('enabled'):
        app_gui_instance.send_email_notification_from_gui(f"Bot LBank Iniciado: {target_symbol_input}", start_msg, email_cfg)

    order_placed_or_failed_critically = False
    max_check_errors = 3
    current_check_errors = 0
    try:
        while not stop_bot_event.is_set() and not order_placed_or_failed_critically:
            gui_log_from_thread(f"INFO: Verificando si '{target_symbol_input}' está listado...")
            is_tradable_result = check_if_symbol_is_tradable_gui(target_symbol_input, gui_log_from_thread)
            if stop_bot_event.is_set(): break
            if is_tradable_result is None: 
                current_check_errors += 1
                gui_log_from_thread(f"ADVERTENCIA: Error verificando símbolo. Intento {current_check_errors}/{max_check_errors}.", "warning")
                if current_check_errors >= max_check_errors:
                    gui_log_from_thread("ERROR: Máximo de errores verificando símbolo. Deteniendo bot.", "error")
                    order_placed_or_failed_critically = True; break
                for _ in range(poll_interval):
                    if stop_bot_event.is_set(): break
                    pytime.sleep(1) 
                if stop_bot_event.is_set(): break
                continue
            current_check_errors = 0 
            if is_tradable_result:
                gui_log_from_thread(f"INFO: ¡SÍMBOLO '{target_symbol_input}' DETECTADO COMO TRADABLE!", "success")
                detect_msg = f"🎯 ¡{target_symbol_input} detectado en LBank! Intentando comprar AHORA..."
                if tg_cfg.get('enabled'): app_gui_instance.send_telegram_notification_from_gui(detect_msg, tg_cfg)
                if email_cfg.get('enabled'): app_gui_instance.send_email_notification_from_gui(f"{target_symbol_input} Detectado!", detect_msg, email_cfg)

                order_result = place_lbank_market_buy_order_gui(target_symbol_input, usdt_to_invest, api_cfg, gui_log_from_thread, excel_log_from_thread)
                
                if stop_after_attempt:
                    gui_log_from_thread("INFO: Opción 'Detener tras 1er intento' activada. El bot se detendrá.", "info")
                    order_placed_or_failed_critically = True 
            else:
                gui_log_from_thread(f"INFO: '{target_symbol_input}' aún no operable. Esperando {poll_interval}s...", "info")
            
            if not order_placed_or_failed_critically:
                for _ in range(poll_interval):
                    if stop_bot_event.is_set(): break
                    pytime.sleep(1) 
        if stop_bot_event.is_set():
            gui_log_from_thread("INFO: Bot detenido por el usuario.", "warning")
    except Exception as e:
        tb_str = traceback.format_exc()
        gui_log_from_thread(f"CRÍTICO: ERROR CRÍTICO en el hilo del bot: {e}\n{tb_str}", "critical_error")
        critical_msg = f"💥 ERROR CRÍTICO Bot LBank: {e}"
        if tg_cfg.get('enabled'): app_gui_instance.send_telegram_notification_from_gui(critical_msg, tg_cfg)
        if email_cfg.get('enabled'): app_gui_instance.send_email_notification_from_gui(f"Error Crítico Bot LBank: {target_symbol_input}", f"{critical_msg}\n\n{tb_str}", email_cfg)
    finally:
        gui_log_from_thread("INFO: Hilo del bot finalizado.", "info")
        log_queue.put(("BOT_THREAD_FINISHED", "control"))

class LBankSniperAppGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"LBank Sniper Bot GUI - v{VERSION}")
        self.geometry("800x850") 
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.bot_running = False
        self.log_data_for_excel_gui = []

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1) 

        main_config_frame = ctk.CTkFrame(self)
        main_config_frame.grid(row=0, column=0, padx=10, pady=(10,5), sticky="ew")
        main_config_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(main_config_frame, text="Símbolo (ej. KING1):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.symbol_entry = ctk.CTkEntry(main_config_frame, placeholder_text="PEPE")
        self.symbol_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(main_config_frame, text="Monto USDT:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.amount_entry = ctk.CTkEntry(main_config_frame, placeholder_text="10.0")
        self.amount_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        ctk.CTkLabel(main_config_frame, text="API Key LBank:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.api_key_entry = ctk.CTkEntry(main_config_frame, placeholder_text="Tu API Key de LBank", width=350)
        self.api_key_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(main_config_frame, text="Secret Key LBank:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.secret_key_entry = ctk.CTkEntry(main_config_frame, show="*", placeholder_text="Tu Secret Key de LBank", width=350)
        self.secret_key_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(main_config_frame, text="Intervalo Sondeo (s):").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.interval_entry = ctk.CTkEntry(main_config_frame, placeholder_text="2", width=100)
        self.interval_entry.grid(row=4, column=1, padx=5, pady=5, sticky="w") 

        self.options_notifications_frame = ctk.CTkFrame(self)
        self.options_notifications_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        
        self.stop_after_attempt_var = tk.BooleanVar(value=True)
        self.stop_after_attempt_check = ctk.CTkCheckBox(self.options_notifications_frame, text="Detener bot tras 1er intento de compra", variable=self.stop_after_attempt_var)
        self.stop_after_attempt_check.pack(anchor="w", padx=10, pady=(5,2))

        self.telegram_enabled_var = tk.BooleanVar(value=False)
        self.telegram_check = ctk.CTkCheckBox(self.options_notifications_frame, text="Activar Notificaciones por Telegram", variable=self.telegram_enabled_var, command=self.toggle_notification_fields)
        self.telegram_check.pack(anchor="w", padx=10, pady=2)
        
        self.telegram_config_frame = ctk.CTkFrame(self.options_notifications_frame) 
        ctk.CTkLabel(self.telegram_config_frame, text="Token Bot Telegram:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.telegram_token_entry = ctk.CTkEntry(self.telegram_config_frame, width=300)
        self.telegram_token_entry.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ctk.CTkLabel(self.telegram_config_frame, text="Chat ID Telegram:").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.telegram_chat_id_entry = ctk.CTkEntry(self.telegram_config_frame, width=150)
        self.telegram_chat_id_entry.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        self.telegram_config_frame.grid_columnconfigure(1, weight=1)

        self.email_enabled_var = tk.BooleanVar(value=False)
        self.email_check = ctk.CTkCheckBox(self.options_notifications_frame, text="Activar Notificaciones por Email (Gmail)", variable=self.email_enabled_var, command=self.toggle_notification_fields)
        self.email_check.pack(anchor="w", padx=10, pady=2)

        self.email_config_frame = ctk.CTkFrame(self.options_notifications_frame) 
        ctk.CTkLabel(self.email_config_frame, text="Email Remitente (Gmail):").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.email_sender_entry = ctk.CTkEntry(self.email_config_frame, width=250)
        self.email_sender_entry.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ctk.CTkLabel(self.email_config_frame, text="Contraseña de Aplicación Gmail:").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.email_password_entry = ctk.CTkEntry(self.email_config_frame, show="*", width=200)
        self.email_password_entry.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        ctk.CTkLabel(self.email_config_frame, text="Email Destinatario Notificaciones:").grid(row=2, column=0, padx=5, pady=2, sticky="w")
        self.email_receiver_entry = ctk.CTkEntry(self.email_config_frame, width=250)
        self.email_receiver_entry.grid(row=2, column=1, padx=5, pady=2, sticky="ew")
        self.email_config_frame.grid_columnconfigure(1, weight=1)
        
        controls_frame = ctk.CTkFrame(self)
        controls_frame.grid(row=2, column=0, padx=10, pady=5, sticky="ew")
        controls_frame.grid_columnconfigure(0, weight=1)
        controls_frame.grid_columnconfigure(1, weight=1)
        controls_frame.grid_columnconfigure(2, weight=1)

        self.save_config_button = ctk.CTkButton(controls_frame, text="Guardar Config", command=self.save_config_button_action)
        self.save_config_button.grid(row=0, column=0, padx=5, pady=10, sticky="ew")
        
        self.start_button = ctk.CTkButton(controls_frame, text="INICIAR BOT", command=self.start_bot, fg_color="green")
        self.start_button.grid(row=0, column=1, padx=5, pady=10, sticky="ew")

        self.stop_button = ctk.CTkButton(controls_frame, text="PARAR BOT", command=self.stop_bot, state=tk.DISABLED, fg_color="red")
        self.stop_button.grid(row=0, column=2, padx=5, pady=10, sticky="ew")
        
        log_frame = ctk.CTkFrame(self)
        log_frame.grid(row=4, column=0, padx=10, pady=(5,10), sticky="nsew") 
        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)

        self.log_textbox = scrolledtext.ScrolledText(log_frame, height=15, state=tk.DISABLED, wrap=tk.WORD, bg="#2B2B2B", fg="white", insertbackground="white", relief=tk.FLAT, borderwidth=0)
        self.log_textbox.grid(row=0, column=0, sticky="nsew")
        self.log_textbox.tag_config("INFO", foreground="#A9A9A9") 
        self.log_textbox.tag_config("DEBUG", foreground="cyan")
        self.log_textbox.tag_config("WARNING", foreground="yellow")
        self.log_textbox.tag_config("ERROR", foreground="red")
        self.log_textbox.tag_config("CRITICAL_ERROR", foreground="red", font=('TkDefaultFont', 9, 'bold'))
        self.log_textbox.tag_config("SUCCESS", foreground="green")

        self.load_config_from_yaml()
        self.toggle_notification_fields() 
        self.process_log_queue()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.gui_log_to_display("INFO: Aplicación iniciada. Configura y presiona INICIAR BOT.", "INFO")

    def toggle_notification_fields(self):
        # ... (sin cambios) ...
        if self.telegram_enabled_var.get():
            self.telegram_config_frame.pack(fill="x", padx=10, pady=(2,5), before=self.email_check)
        else:
            self.telegram_config_frame.pack_forget()
        if self.email_enabled_var.get():
            self.email_config_frame.pack(fill="x", padx=10, pady=(2,5), after=self.email_check)
        else:
            self.email_config_frame.pack_forget()
            
    def send_telegram_notification_from_gui(self, message_text, tg_config):
        # ... (sin cambios) ...
        if not tg_config.get('enabled'): return
        bot_token = tg_config.get('bot_token')
        chat_id = tg_config.get('chat_id')
        if not bot_token or not chat_id:
            self.gui_log_to_display("ADVERTENCIA: Telegram habilitado, pero token o chat_id no configurados.", "WARNING")
            return
        try:
            from telegram import Bot; from telegram.error import TelegramError
            bot = Bot(token=bot_token)
            bot.send_message(chat_id=chat_id, text=message_text, parse_mode='Markdown')
            self.gui_log_to_display("INFO: Notificación de Telegram enviada.", "INFO")
        except Exception as e: self.gui_log_to_display(f"ERROR: Error enviando notificación de Telegram: {e}", "ERROR")

    def send_email_notification_from_gui(self, subject_line, body_text, email_config):
        # ... (sin cambios) ...
        if not email_config.get('enabled'): return
        sender = email_config.get('sender_email')
        password = email_config.get('app_password')
        receiver = email_config.get('receiver_email')
        if not sender or not password or not receiver:
            self.gui_log_to_display("ADVERTENCIA: Email habilitado, pero configuración incompleta.", "WARNING")
            return
        try:
            msg = MIMEText(body_text); msg['Subject'] = subject_line; msg['From'] = sender; msg['To'] = receiver
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s: s.login(sender, password); s.sendmail(sender, receiver, msg.as_string())
            self.gui_log_to_display(f"INFO: Email enviado a '{receiver}'.", "INFO")
        except Exception as e: self.gui_log_to_display(f"ERROR: Error enviando email: {e}", "ERROR")

    def gui_log_to_display(self, message, level="INFO"):
        # ... (sin cambios) ...
        self.log_textbox.configure(state=tk.NORMAL)
        timestamp = datetime.datetime.now().strftime("%H:%M:%S") 
        level_upper = level.upper()
        num_lines = int(self.log_textbox.index('end-1c').split('.')[0])
        if num_lines > 500: 
            self.log_textbox.delete('1.0', f'{num_lines - 499}.0')
        self.log_textbox.insert(tk.END, f"[{timestamp}][{level_upper}] {message}\n", level_upper)
        self.log_textbox.see(tk.END)
        self.log_textbox.configure(state=tk.DISABLED)
        console_logger.info(f"GUI_LOG: [{level_upper}] {message}")

    def add_entry_to_excel_log_gui(self, symbol, timestamp, amount_bought_token, price_usdt, total_usdt_spent, order_id, status):
        # ... (sin cambios) ...
        log_entry = {
            "Símbolo": symbol, "Fecha/Hora": timestamp.strftime("%Y-%m-%d %H:%M:%S"),
            "Cantidad Comprada (Token)": amount_bought_token, "Precio Ejecutado (USDT)": price_usdt,
            "Monto Total Gastado (USDT)": total_usdt_spent, "ID de Orden": order_id, "Estado": status
        }
        self.log_data_for_excel_gui.append(log_entry)
        self.gui_log_to_display(f"INFO: Entrada para Excel: {symbol}, Estado: {status}", "INFO")
        self.save_all_logs_to_excel_file_gui()

    def save_all_logs_to_excel_file_gui(self):
        # ... (sin cambios) ...
        if not self.log_data_for_excel_gui: return
        df_new = pd.DataFrame(self.log_data_for_excel_gui)
        try:
            if os.path.exists(EXCEL_LOG_FILE):
                df_old = pd.read_excel(EXCEL_LOG_FILE)
                df_all = pd.concat([df_old, df_new], ignore_index=True)
                df_all.drop_duplicates(subset=['ID de Orden', 'Fecha/Hora', 'Símbolo'], keep='last', inplace=True)
            else: df_all = df_new
            df_all.to_excel(EXCEL_LOG_FILE, index=False, sheet_name="ComprasBot")
            self.gui_log_to_display(f"INFO: Log guardado/actualizado en '{EXCEL_LOG_FILE}'", "INFO")
            self.log_data_for_excel_gui.clear()
        except Exception as e: self.gui_log_to_display(f"ERROR: Error guardando log Excel: {e}", "ERROR")

    def process_log_queue(self):
        # ... (sin cambios) ...
        try:
            while True:
                message_tuple = log_queue.get_nowait()
                if isinstance(message_tuple, tuple) and len(message_tuple) == 2:
                    message, level_or_type = message_tuple
                    if level_or_type == "control" and message == "BOT_THREAD_FINISHED": self.bot_has_finished()
                    elif level_or_type == "excel_data" and isinstance(message, tuple): self.add_entry_to_excel_log_gui(*message)
                    else: self.gui_log_to_display(message, level_or_type)
                else: console_logger.warning(f"Mensaje de cola malformado: {message_tuple}")
        except queue.Empty: pass
        finally: self.after(100, self.process_log_queue)

    def get_current_config_from_gui(self):
        # ### MODIFICADO para incluir nuevas opciones de GUI ###
        try:
            symbol_input = self.symbol_entry.get().strip().upper()
            if not symbol_input:
                messagebox.showerror("Error de Configuración", "El 'Símbolo de la Moneda' no puede estar vacío."); return None
            
            config = {
                'lbank_api_key': self.api_key_entry.get(),
                'lbank_secret_key': self.secret_key_entry.get(),
                'target_symbol_input': symbol_input, 
                'usdt_amount_to_invest': float(self.amount_entry.get()),
                'poll_interval_seconds': int(self.interval_entry.get()),
                'stop_after_attempt': self.stop_after_attempt_var.get(), 
                'telegram_config': {
                    'enabled': self.telegram_enabled_var.get(),
                    'bot_token': self.telegram_token_entry.get(),
                    'chat_id': self.telegram_chat_id_entry.get()
                },
                'email_config': {
                    'enabled': self.email_enabled_var.get(),
                    'sender_email': self.email_sender_entry.get(),
                    'app_password': self.email_password_entry.get(),
                    'receiver_email': self.email_receiver_entry.get()
                }
            }
            if config['usdt_amount_to_invest'] <= 0: messagebox.showerror("Error", "'Monto a Invertir' > 0."); return None
            if config['poll_interval_seconds'] <= 0: messagebox.showerror("Error", "'Intervalo Sondeo' > 0."); return None
            if not config['lbank_api_key'] or not config['lbank_secret_key']: messagebox.showerror("Error", "Claves API LBank no pueden estar vacías."); return None

            if config['telegram_config']['enabled'] and (not config['telegram_config']['bot_token'] or not config['telegram_config']['chat_id']):
                messagebox.showerror("Error de Configuración", "Si Telegram está activado, el Token y Chat ID son requeridos.")
                return None
            if config['email_config']['enabled'] and (not config['email_config']['sender_email'] or not config['email_config']['app_password'] or not config['email_config']['receiver_email']):
                messagebox.showerror("Error de Configuración", "Si Email está activado, todos los campos de Email son requeridos.")
                return None
            return config
        except ValueError: messagebox.showerror("Error de Configuración", "Asegúrate de que 'Monto a Invertir' e 'Intervalo de Sondeo' sean números válidos."); return None
        except Exception as e: messagebox.showerror("Error de Configuración", f"Error inesperado al leer la configuración: {e}"); return None

    def save_config_button_action(self):
        # ... (sin cambios) ...
        current_gui_config = self.get_current_config_from_gui()
        if current_gui_config:
            self.save_config_to_yaml(current_gui_config)
            self.gui_log_to_display("INFO: Configuración guardada en config_gui.yml", "INFO")
            messagebox.showinfo("Guardado", "Configuración guardada exitosamente.")
        else:
            self.gui_log_to_display("ADVERTENCIA: Configuración no guardada debido a errores.", "WARNING")

    def start_bot(self):
        # ... (sin cambios) ...
        global bot_thread
        if self.bot_running: messagebox.showwarning("Advertencia", "El bot ya se está ejecutando."); return
        current_gui_config = self.get_current_config_from_gui()
        if not current_gui_config: return
        self.gui_log_to_display("INFO: Iniciando bot...", "INFO")
        self.save_config_to_yaml(current_gui_config) 
        stop_bot_event.clear()
        bot_thread = threading.Thread(target=bot_logic_thread_gui, args=(self, current_gui_config), daemon=True)
        bot_thread.start()
        self.bot_running = True
        self.start_button.configure(state=tk.DISABLED)
        self.stop_button.configure(state=tk.NORMAL)
        self.gui_log_to_display("INFO: Bot iniciado. Monitoreando...", "INFO")

    def stop_bot(self):
        # ... (sin cambios) ...
        if not self.bot_running or not bot_thread or not bot_thread.is_alive():
            self.gui_log_to_display("ADVERTENCIA: El bot no se está ejecutando.", "WARNING"); self.bot_has_finished(); return
        self.gui_log_to_display("INFO: Intentando detener el bot...", "INFO")
        stop_bot_event.set()

    def bot_has_finished(self):
        # ... (sin cambios) ...
        self.bot_running = False
        self.start_button.configure(state=tk.NORMAL)
        self.stop_button.configure(state=tk.DISABLED)
        self.gui_log_to_display("INFO: El bot ha sido detenido o ha finalizado.", "INFO")

    def load_config_from_yaml(self):
        # ### MODIFICADO para cargar nuevas opciones de GUI ###
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r") as f: config_data = yaml.safe_load(f)
                if config_data:
                    self.api_key_entry.insert(0, config_data.get('lbank_api_key', ''))
                    self.secret_key_entry.insert(0, config_data.get('lbank_secret_key', ''))
                    self.symbol_entry.insert(0, config_data.get('target_symbol_input', ''))
                    self.amount_entry.insert(0, str(config_data.get('usdt_amount_to_invest', '10.0')))
                    self.interval_entry.insert(0, str(config_data.get('poll_interval_seconds', '2')))
                    self.stop_after_attempt_var.set(config_data.get('stop_after_attempt', True))

                    tg_cfg = config_data.get('telegram_config', {})
                    self.telegram_enabled_var.set(tg_cfg.get('enabled', False))
                    self.telegram_token_entry.delete(0, tk.END) 
                    self.telegram_token_entry.insert(0, tg_cfg.get('bot_token', ''))
                    self.telegram_chat_id_entry.delete(0, tk.END)
                    self.telegram_chat_id_entry.insert(0, tg_cfg.get('chat_id', ''))
                    
                    email_cfg = config_data.get('email_config', {})
                    self.email_enabled_var.set(email_cfg.get('enabled', False))
                    self.email_sender_entry.delete(0, tk.END)
                    self.email_sender_entry.insert(0, email_cfg.get('sender_email', ''))
                    self.email_password_entry.delete(0, tk.END)
                    self.email_password_entry.insert(0, email_cfg.get('app_password', ''))
                    self.email_receiver_entry.delete(0, tk.END)
                    self.email_receiver_entry.insert(0, email_cfg.get('receiver_email', ''))
                    
                    self.gui_log_to_display("INFO: Configuración cargada desde config_gui.yml", "INFO")
            else: self.gui_log_to_display("INFO: config_gui.yml no encontrado. Usar valores por defecto.", "INFO")
        except Exception as e: self.gui_log_to_display(f"ERROR: Error cargando config_gui.yml: {e}", "ERROR")


    def save_config_to_yaml(self, config_to_save):
        # ... (sin cambios) ...
        if not config_to_save: self.gui_log_to_display("ADVERTENCIA: No se guardó config (datos inválidos).", "WARNING"); return
        try:
            with open(CONFIG_FILE, "w") as f: yaml.dump(config_to_save, f, sort_keys=False, default_flow_style=False)
        except Exception as e: self.gui_log_to_display(f"ERROR: Error guardando config_gui.yml: {e}", "ERROR")
            
    def on_closing(self):
        # ... (sin cambios) ...
        if self.bot_running:
            if messagebox.askyesno("Salir", "El bot se está ejecutando. ¿Seguro que quieres salir? El bot se detendrá."):
                self.stop_bot(); self.after(500, self.destroy) 
            else: return
        else:
            current_cfg = self.get_current_config_from_gui() 
            if current_cfg : self.save_config_to_yaml(current_cfg)
            self.destroy()

if __name__ == "__main__":
    console_logger.info("Iniciando aplicación LBank Sniper Bot GUI...")
    abspath = os.path.abspath(__file__)
    dname = os.path.dirname(abspath)
    os.chdir(dname)
    app = LBankSniperAppGUI()
    try:
        app.mainloop()
    except Exception as e:
        console_logger.critical(f"Error fatal en GUI: {e}\n{traceback.format_exc()}")
    finally:
        console_logger.info("Aplicación LBank Sniper Bot GUI cerrada.")

