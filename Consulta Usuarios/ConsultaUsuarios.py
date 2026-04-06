# ================================================================
# IMPORTS INMEDIATOS — solo lo estrictamente necesario para
# mostrar la ventana. Todo lo pesado se carga en background.
# ================================================================
import sys
import os
import threading
import json
import re
import unicodedata
import hashlib
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor

import customtkinter as ctk  # Inevitable para la UI

# ================================================================
# IMPORTS DIFERIDOS — se asignan en _cargar_imports_pesados()
# ================================================================
requests    = None
msal_mod    = None
msalext_mod = None
_SESSION    = None   # se crea tras cargar requests

PublicClientApplication           = None
FilePersistenceWithDataProtection = None
PersistedTokenCache               = None

# Flag que indica que los imports pesados ya están listos
_imports_listos = threading.Event()


def _cargar_imports_pesados():
    """
    Carga todos los módulos costosos en un thread de background.
    El usuario ya ve la ventana splash mientras esto ocurre.
    """
    global requests, msal_mod, msalext_mod, _SESSION
    global PublicClientApplication, FilePersistenceWithDataProtection, PersistedTokenCache

    import requests as _req
    requests = _req
    _SESSION = _req.Session()

    import msal as _msal
    msal_mod = _msal
    PublicClientApplication = _msal.PublicClientApplication

    from msal_extensions import (
        FilePersistenceWithDataProtection as _FPDP,
        PersistedTokenCache as _PTC
    )
    FilePersistenceWithDataProtection = _FPDP
    PersistedTokenCache = _PTC

    _imports_listos.set()
    print("[imports] Módulos pesados cargados")


# ================================================================
# CONFIGURACIÓN ENTRA ID
# ================================================================

CLIENT_ID        = os.getenv("CLIENT_ID")
TENANT_ID        = os.getenv("TENANT_ID")
AUTHORITY        = f"https://login.microsoftonline.com/{TENANT_ID}"

SCOPES = [
    "User.Read.All",
    "GroupMember.Read.All",
    "Group.ReadWrite.All",
    "User.ReadWrite.All",
    "Files.Read.All"
]

GRUPO_REQUERIDO  = os.getenv("GRUPO_REQUERIDO", "ConsultaUsuarios")
TOKEN                = None
USUARIO_LOGADO       = ""
app_msal             = None
PUEDE_GENERAR_CORREO = False

# ================================================================
# WRAPPERS GRAPH
# ================================================================

def graph_get(url, timeout=10):
    headers = {"Authorization": f"Bearer {TOKEN}"}
    return _SESSION.get(url, headers=headers, timeout=timeout)

def graph_post(url, payload, timeout=10):
    headers = {"Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json"}
    return _SESSION.post(url, json=payload, headers=headers, timeout=timeout)

def graph_patch(url, payload, timeout=10):
    headers = {"Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json"}
    return _SESSION.patch(url, json=payload, headers=headers, timeout=timeout)


# ================================================================
# CACHE PROPIETARIOS DE GRUPOS
# ================================================================

_PROPIETARIOS_CACHE = {}
_PROPIETARIOS_LISTO = threading.Event()

GRUPOS = {
    "sap success factors pro": os.getenv("GRUPO_ID_SAP"),
    "empleados_mfa":           os.getenv("GRUPO_ID_MFA"),
    "vpn":                     os.getenv("GRUPO_ID_VPN"),
    "rrhh":                    os.getenv("GRUPO_ID_RRHH"),
}

def precargar_propietarios():
    global _PROPIETARIOS_CACHE
    logged_upn = USUARIO_LOGADO.strip().lower()

    def check_owner(item):
        nombre, gid = item
        try:
            resp = graph_get(f"https://graph.microsoft.com/v1.0/groups/{gid}/owners")
            if resp.status_code == 200:
                owners = [o.get("userPrincipalName", "").lower()
                        for o in resp.json().get("value", [])]
                return gid, logged_upn in owners
        except Exception as e:
            print(f"[propietarios] Error en {nombre}: {e}")
        return gid, False

    with ThreadPoolExecutor(max_workers=4) as ex:
        for gid, es_owner in ex.map(check_owner, GRUPOS.items()):
            _PROPIETARIOS_CACHE[gid] = es_owner

    _PROPIETARIOS_LISTO.set()
    print(f"[propietarios] Cache lista: {_PROPIETARIOS_CACHE}")


# ================================================================
# TOOLTIP
# ================================================================

class ToolTip:
    def __init__(self, widget, text):
        self.widget    = widget
        self.text      = text
        self.tipwindow = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tipwindow:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + 20
        self.tipwindow = tw = ctk.CTkToplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.geometry(f"+{x}+{y}")
        ctk.CTkLabel(
            tw, text=self.text,
            fg_color="#1f2937", text_color="white",
            corner_radius=6, padx=8, pady=4,
            font=("Segoe UI", 11)
        ).pack()

    def hide_tip(self, event=None):
        if self.tipwindow:
            self.tipwindow.destroy()
            self.tipwindow = None


# ================================================================
# SHAREPOINT — INDEX FICHEROS
# ================================================================

CARPETA_CACHE    = "cache"
os.makedirs(CARPETA_CACHE, exist_ok=True)

RUTA_INDEX_LOCAL = os.path.join(CARPETA_CACHE, "index_ficheros.json")
RUTA_SSFF_LOCAL  = os.path.join(CARPETA_CACHE, "ssff_cache.json")

INDEX_ID     = {}
INDEX_CORREO = {}
INDEX_NOMBRE = {}

SSFF_DATA     = {}
SSFF_ID_INDEX = {}
SSFF_CARGADO  = False


def descargar_index_sharepoint():
    global INDEX_ID, INDEX_CORREO, INDEX_NOMBRE

    if not TOKEN:
        return

    try:
        site_url = f"https://graph.microsoft.com/v1.0/sites/{_sp_host}:{_sp_site}"
        resp = graph_get(site_url, timeout=20)
        resp.raise_for_status()
        site_id = resp.json()["id"]

        drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        resp = graph_get(drives_url, timeout=20)
        resp.raise_for_status()

        drive = next((d for d in resp.json().get("value", []) if d.get("name") == "Ficheros de origen DA"), None)
        if not drive:
            mostrar_error("No se encontró la biblioteca 'Ficheros de origen DA' en SharePoint.")
            return

        drive_id = drive["id"]
        file_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/index_ficheros.json:/content"
        resp = graph_get(file_url, timeout=20)

        if resp.status_code != 200:
            mostrar_error(f"No se pudo descargar index:\n{resp.status_code}\n{resp.text}")
            return

        with open(RUTA_INDEX_LOCAL, "wb") as f:
            f.write(resp.content)

        with open(RUTA_INDEX_LOCAL, "r", encoding="utf-8") as f:
            data = json.load(f)

        INDEX_ID     = {}
        INDEX_CORREO = {}
        INDEX_NOMBRE = {}

        for persona in data:
            fichero = persona.get("fichero_origen")
            emp_id  = persona.get("id_empleado")
            if emp_id:
                INDEX_ID[str(emp_id)] = fichero
            correo = persona.get("correo")
            if correo:
                INDEX_CORREO[correo.lower()] = fichero
            clave = f"{persona.get('nombre','')} {persona.get('apellido1','')} {persona.get('apellido2','')}".strip().lower()
            if clave:
                INDEX_NOMBRE[clave] = fichero

        print("✅ Index de ficheros cargado")

    except Exception as e:
        print("Error cargando index:", e)
        mostrar_error(f"Error cargando index:\n{e}")


def cargar_json_ssff():
    """Descarga condicional: solo si el remoto es más nuevo que la copia local."""
    global SSFF_DATA, SSFF_ID_INDEX, SSFF_CARGADO

    if SSFF_CARGADO:
        return
    if not TOKEN:
        return

    def construir_indice():
        global SSFF_ID_INDEX
        SSFF_ID_INDEX = {}
        for correo, datos in SSFF_DATA.items():
            emp_id = str(datos.get("id_empleado", "")).strip()
            if emp_id:
                SSFF_ID_INDEX[emp_id.lstrip("0")] = {
                    "id": emp_id,
                    "nombre": datos.get("nombre"),
                    "correo": correo
                }

    try:
        site_url = f"https://graph.microsoft.com/v1.0/sites/{_sp_host}:{_sp_site}"
        resp = graph_get(site_url)
        resp.raise_for_status()
        site_id = resp.json()["id"]

        drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        resp = graph_get(drives_url)
        resp.raise_for_status()

        drive    = next(d for d in resp.json()["value"] if d["name"] == "Informe SSFF")
        drive_id = drive["id"]

        folder_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"
        resp = graph_get(folder_url)
        resp.raise_for_status()

        json_files = [
            f for f in resp.json()["value"]
            if f["name"].startswith("InformeSSFF") and f["name"].endswith(".json")
        ]
        if not json_files:
            print("[SSFF] No se encontró JSON SSFF")
            return

        ultimo = max(json_files, key=lambda x: x["lastModifiedDateTime"])

        # Descarga condicional
        if os.path.exists(RUTA_SSFF_LOCAL):
            local_mtime  = datetime.fromtimestamp(os.path.getmtime(RUTA_SSFF_LOCAL))
            remote_mtime = datetime.fromisoformat(
                ultimo["lastModifiedDateTime"].replace("Z", "+00:00")
            ).replace(tzinfo=None)

            if local_mtime >= remote_mtime:
                print("[SSFF] Caché local vigente, omitiendo descarga.")
                with open(RUTA_SSFF_LOCAL, "r", encoding="utf-8") as f:
                    SSFF_DATA = json.load(f)
                construir_indice()
                SSFF_CARGADO = True
                print(f"[SSFF] Cargado desde caché: {len(SSFF_DATA)} usuarios")
                return

        print(f"[SSFF] Descargando: {ultimo['name']}")
        r = _SESSION.get(ultimo["@microsoft.graph.downloadUrl"])

        with open(RUTA_SSFF_LOCAL, "wb") as f:
            f.write(r.content)

        with open(RUTA_SSFF_LOCAL, "r", encoding="utf-8") as f:
            SSFF_DATA = json.load(f)

        construir_indice()
        SSFF_CARGADO = True
        print(f"[SSFF] Descargado y cargado: {len(SSFF_DATA)} usuarios")

    except Exception as e:
        print("Error cargando JSON SSFF:", e)


def comprobar_ssff(upn, emp_id, correo_personal=None):
    if not SSFF_DATA:
        return None, None, None

    upn             = (upn or "").lower().strip()
    correo_personal = (correo_personal or "").lower().strip()
    emp_id          = str(emp_id or "").strip()
    emp_id_limpio   = emp_id.lstrip("0")

    usuario = SSFF_DATA.get(upn)
    if usuario:
        id_ssff = str(usuario.get("id_empleado", "")).strip()
        nombre  = usuario.get("nombre") or ""
        return (True if id_ssff.lstrip("0") == emp_id_limpio else False), id_ssff, nombre

    if correo_personal:
        usuario = SSFF_DATA.get(correo_personal)
        if usuario:
            id_ssff = str(usuario.get("id_empleado", "")).strip()
            nombre  = usuario.get("nombre") or ""
            return (True if id_ssff.lstrip("0") == emp_id_limpio else False), id_ssff, nombre

    if emp_id_limpio in SSFF_ID_INDEX:
        datos = SSFF_ID_INDEX[emp_id_limpio]
        return "id", datos["id"], datos["nombre"]

    if len(emp_id_limpio) >= 4:
        ultimos = emp_id_limpio[-4:]
        for id_json, datos in SSFF_ID_INDEX.items():
            if id_json.endswith(ultimos):
                return "parecido", datos["id"], datos["nombre"]

    return None, None, None


# ================================================================
# CACHE DE TOKENS MULTIUSUARIO
# ================================================================

CACHE_DIR = "token_caches"
os.makedirs(CACHE_DIR, exist_ok=True)

def get_cache_path(username):
    h = hashlib.sha256(username.lower().encode()).hexdigest()
    return os.path.join(CACHE_DIR, f"{h}_token_cache.bin")

def build_cache(path):
    persistence = FilePersistenceWithDataProtection(path)
    return PersistedTokenCache(persistence)


# ================================================================
# FUNCIONES AUXILIARES
# ================================================================

from dotenv import load_dotenv
load_dotenv()

CLIENT_ID        = os.getenv("CLIENT_ID")
TENANT_ID        = os.getenv("TENANT_ID")
AUTHORITY        = f"https://login.microsoftonline.com/{TENANT_ID}"
GRUPO_REQUERIDO  = os.getenv("GRUPO_REQUERIDO", "ConsultaUsuarios")

GRUPOS = {
    "sap success factors pro": os.getenv("GRUPO_ID_SAP"),
    "empleados_mfa":           os.getenv("GRUPO_ID_MFA"),
    "vpn":                     os.getenv("GRUPO_ID_VPN"),
    "rrhh":                    os.getenv("GRUPO_ID_RRHH"),
}

_sp_host    = os.getenv("SHAREPOINT_HOST")
_sp_site    = os.getenv("SHAREPOINT_SITE")

CORREOS_AUTORIZADOS = [c.strip() for c in os.getenv("CORREOS_AUTORIZADOS", "").split(",") if c.strip()]


def normalizar_texto(texto):
    texto = texto.lower().strip()
    texto = unicodedata.normalize("NFD", texto)
    texto = texto.encode("ascii", "ignore").decode("utf-8")
    return re.sub(r"[^a-z]", "", texto)


def formatear_fecha(fecha):
    if not fecha:
        return ""
    try:
        dt = datetime.fromisoformat(fecha.replace("Z", "+00:00"))
        return dt.strftime("%d/%m/%Y %H:%M")
    except:
        return fecha


def mostrar_error(texto):
    v = ctk.CTkToplevel()
    v.title("Error")
    v.geometry("420x180")
    v.grab_set()
    ctk.CTkLabel(v, text="⚠ Error", font=("Segoe UI", 18, "bold"), text_color="#ff4d4d").pack(pady=(20, 5))
    ctk.CTkLabel(v, text=texto, wraplength=380, justify="center").pack(padx=20, pady=10)
    ctk.CTkButton(v, text="Cerrar", command=v.destroy, corner_radius=20).pack(pady=15)


def mostrar_exito(texto):
    v = ctk.CTkToplevel()
    v.title("Éxito")
    v.geometry("420x180")
    v.grab_set()
    v.after(5000, v.destroy)
    ctk.CTkLabel(v, text="✅ Éxito", font=("Segoe UI", 18, "bold"), text_color="#16a34a").pack(pady=(20, 5))
    ctk.CTkLabel(v, text=texto, wraplength=380, justify="center").pack(padx=20, pady=10)
    ctk.CTkButton(v, text="Cerrar", command=v.destroy, corner_radius=20).pack(pady=15)


# ================================================================
# VALIDAR GRUPO — $batch: 1 request en vez de 2
# ================================================================

def validar_grupo(token):
    global USUARIO_LOGADO, PUEDE_GENERAR_CORREO

    batch_url = "https://graph.microsoft.com/v1.0/$batch"
    payload = {
        "requests": [
            {"id": "me",     "method": "GET", "url": "/me?$select=id,userPrincipalName"},
            {"id": "grupos", "method": "GET", "url": "/me/memberOf?$select=id,displayName"}
        ]
    }

    try:
        resp = _SESSION.post(
            batch_url, json=payload,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            timeout=15
        )
        resp.raise_for_status()
        responses = {r["id"]: r for r in resp.json()["responses"]}

        if responses["me"]["status"] != 200:
            mostrar_error("Error obteniendo usuario")
            return False

        me          = responses["me"]["body"]
        grupos_data = responses["grupos"]["body"].get("value", []) \
                    if responses["grupos"]["status"] == 200 else []

    except Exception as e:
        mostrar_error(f"Error en batch de validación: {e}")
        return False

    USUARIO_LOGADO = me.get("userPrincipalName", "")
    ids    = [g.get("id") for g in grupos_data]
    nombres = [g.get("displayName") for g in grupos_data if g.get("displayName")]

    if os.getenv("GRUPO_ID_RRHH", "") in ids:
        PUEDE_GENERAR_CORREO = True
    if any("rrhh" in n.lower() for n in nombres):
        PUEDE_GENERAR_CORREO = True
    if USUARIO_LOGADO.lower() in [c.lower() for c in CORREOS_AUTORIZADOS]:
        PUEDE_GENERAR_CORREO = True

    if GRUPO_REQUERIDO not in nombres:
        mostrar_error(f"No tienes permisos.\nDebes estar en '{GRUPO_REQUERIDO}'")
        return False

    return True


# ================================================================
# LOGIN
# ================================================================

def _arrancar_tareas_background():
    threading.Thread(target=precargar_propietarios,     daemon=True).start()
    threading.Thread(target=descargar_index_sharepoint, daemon=True).start()
    threading.Thread(target=cargar_json_ssff,           daemon=True).start()


def login():
    global TOKEN, app_msal

    if app_msal is None:
        username = os.getlogin()
        cache    = build_cache(get_cache_path(username))
        globals()["app_msal"] = PublicClientApplication(
            CLIENT_ID, authority=AUTHORITY, token_cache=cache
        )

    accounts = app_msal.get_accounts()
    if accounts:
        try:
            result = app_msal.acquire_token_silent(SCOPES, account=accounts[0])
        except Exception as e:
            mostrar_error(f"Error al leer caché: {e}")
            result = None

        if result and "access_token" in result:
            TOKEN = result["access_token"]
            _arrancar_tareas_background()
            if validar_grupo(TOKEN):
                return True

    try:
        result = app_msal.acquire_token_interactive(scopes=SCOPES)
    except Exception as e:
        mostrar_error(str(e))
        return False

    if "access_token" not in result:
        return False

    TOKEN    = result["access_token"]
    username = result["id_token_claims"]["preferred_username"]
    cache    = build_cache(get_cache_path(username))
    cache.add(result)

    _arrancar_tareas_background()
    return validar_grupo(TOKEN)


# ================================================================
# BUSCAR USUARIO — $batch: detalle + grupos en paralelo
# ================================================================

def buscar_usuario(valor_busqueda):
    valor = valor_busqueda.lower()

    # 1. Solo obtener id (payload mínimo)
    query = (
        f"https://graph.microsoft.com/v1.0/users?"
        f"$filter=("
        f"userPrincipalName eq '{valor_busqueda}' "
        f"or employeeId eq '{valor_busqueda}' "
        f"or mail eq '{valor_busqueda}' "
        f"or mailNickname eq '{valor_busqueda}' "
        f"or proxyAddresses/any(x:x eq 'smtp:{valor}')"
        f")&$select=id"
    )

    try:
        response = graph_get(query)
        response.raise_for_status()
    except Exception as e:
        return {"error": f"Error consultando Microsoft Graph:\n{e}"}

    data = response.json()
    if not data.get("value"):
        return {"error": "Usuario no encontrado"}

    user_id = data["value"][0]["id"]

    # 2. Batch: detalle completo + grupos en paralelo
    batch_payload = {
        "requests": [
            {
                "id": "detalle",
                "method": "GET",
                "url": (
                    f"/users/{user_id}"
                    f"?$select=givenName,surname,userPrincipalName,employeeId,"
                    f"country,createdDateTime,onPremisesExtensionAttributes,"
                    f"accountEnabled,mailNickname"
                )
            },
            {
                "id": "grupos",
                "method": "GET",
                "url": f"/users/{user_id}/memberOf?$select=id,displayName&$top=100"
            }
        ]
    }

    try:
        batch_resp = graph_post(
            "https://graph.microsoft.com/v1.0/$batch",
            batch_payload, timeout=15
        )
        batch_resp.raise_for_status()
        responses = {r["id"]: r for r in batch_resp.json()["responses"]}
    except Exception as e:
        return {"error": f"Error en batch de usuario:\n{e}"}

    if responses["detalle"]["status"] != 200:
        return {"error": "No se pudo obtener el detalle del usuario"}

    user       = responses["detalle"]["body"]
    grupos_raw = responses["grupos"]["body"].get("value", []) \
                if responses["grupos"]["status"] == 200 else []

    # Paginación de grupos (poco frecuente)
    next_link = responses["grupos"]["body"].get("@odata.nextLink")
    while next_link:
        r = graph_get(next_link)
        if r.status_code != 200:
            break
        gdata = r.json()
        grupos_raw.extend(gdata.get("value", []))
        next_link = gdata.get("@odata.nextLink")

    grupos = sorted(
        [g.get("displayName") for g in grupos_raw if g.get("displayName")],
        key=lambda x: str(x).lower()
    )

    onprem          = user.get("onPremisesExtensionAttributes") or {}
    correo_personal = onprem.get("extensionAttribute1")
    if not correo_personal or "@" not in str(correo_personal):
        correo_personal = "No informado"

    return {
        "Nombre":            user.get("givenName", ""),
        "Apellidos":         user.get("surname", ""),
        "UPN":               user.get("userPrincipalName", ""),
        "Alias":             user.get("mailNickname", ""),
        "ID empleado":       user.get("employeeId", ""),
        "País o región":     user.get("country", ""),
        "Fecha creación":    formatear_fecha(user.get("createdDateTime")),
        "Cuenta habilitada": "Sí" if user.get("accountEnabled") else "No",
        "Correo personal":   correo_personal,
        "Grupos":            grupos,
        "id":                user_id
    }


# ================================================================
# CORREO — comprobar existencia y sugerencias en batch
# ================================================================

def correo_existe(correo):
    query = (
        f"https://graph.microsoft.com/v1.0/users?"
        f"$filter=userPrincipalName eq '{correo}' "
        f"or mail eq '{correo}' "
        f"or proxyAddresses/any(x:x eq 'smtp:{correo.lower()}')"
        f"&$select=userPrincipalName"
    )
    try:
        resp = graph_get(query)
        resp.raise_for_status()
        return len(resp.json().get("value", [])) > 0
    except Exception as e:
        print("Error comprobando correo:", e)
        return True


def _generar_candidatos_alias(nombre, apellidos):
    nombre_n         = normalizar_texto(nombre)
    partes           = apellidos.lower().strip().split()
    if not partes:
        return []
    primer_apellido  = normalizar_texto(partes[0])
    segundo_apellido = normalizar_texto(partes[1]) if len(partes) > 1 else ""
    inicial_segundo  = segundo_apellido[0] if segundo_apellido else ""

    candidatos = []
    for i in range(1, len(primer_apellido) + 1):
        alias = (primer_apellido[0] if i == 1 else primer_apellido[:i]) + inicial_segundo
        candidatos.append(f"{nombre_n}.{alias}")
        if len(candidatos) >= 10:
            break
    return candidatos


def generar_sugerencias(nombre, apellidos, dominio="primaprix.eu"):
    candidatos = _generar_candidatos_alias(nombre, apellidos)
    if not candidatos:
        return []

    correos = [f"{a}@{dominio}" for a in candidatos]
    filtros = " or ".join([f"userPrincipalName eq '{c}'" for c in correos[:15]])

    try:
        resp = graph_get(
            f"https://graph.microsoft.com/v1.0/users?$filter={filtros}&$select=userPrincipalName",
            timeout=10
        )
        resp.raise_for_status()
        existentes = {u["userPrincipalName"].lower() for u in resp.json().get("value", [])}
    except Exception as e:
        print("Error batch sugerencias:", e)
        existentes = set()

    return [a for a in candidatos if f"{a}@{dominio}" not in existentes][:4]


def generar_alias_por_defecto(nombre, apellidos):
    nombre = normalizar_texto(nombre)
    partes = [a.strip() for a in apellidos.strip().split() if a.strip()]
    if not partes:
        return nombre
    iniciales = "".join(normalizar_texto(p)[0] for p in partes[:2] if normalizar_texto(p))
    return f"{nombre}.{iniciales}"


# ================================================================
# VENTANA GENERAR CORREO
# ================================================================

def abrir_ventana_generar_correo():
    ventana = ctk.CTkToplevel(app)
    ventana.title("Generar Correo - Entra ID")
    ventana.geometry("600x600")
    ventana.grab_set()
    ventana.resizable(False, False)

    ctk.CTkLabel(ventana, text="📧 Comprobador de Correos", font=("Segoe UI", 22, "bold")).pack(pady=(20, 20))

    entry_nombre    = ctk.CTkEntry(ventana, width=350, placeholder_text="Nombre")
    entry_nombre.pack(pady=8)
    entry_apellidos = ctk.CTkEntry(ventana, width=350, placeholder_text="Apellidos")
    entry_apellidos.pack(pady=8)

    frame_correo = ctk.CTkFrame(ventana, fg_color="transparent")
    frame_correo.pack(pady=8)
    entry_correo = ctk.CTkEntry(frame_correo, width=250, height=40, placeholder_text="alias (ej: javier.mr)")
    entry_correo.pack(side="left", padx=(0, 5))
    ctk.CTkLabel(frame_correo, text="@primaprix.eu", font=("Segoe UI", 14, "bold"), text_color="gray").pack(side="left")

    autocomplete_after_id = None

    def autocompletar_correo(event=None):
        nonlocal autocomplete_after_id
        if not entry_nombre.get().strip() or not entry_apellidos.get().strip():
            return
        if autocomplete_after_id:
            ventana.after_cancel(autocomplete_after_id)
        autocomplete_after_id = ventana.after(
            1000,
            lambda: completar_alias(entry_nombre.get().strip(), entry_apellidos.get().strip())
        )

    def completar_alias(nombre, apellidos_raw):
        if entry_correo.get().strip():
            return
        alias = generar_alias_por_defecto(nombre, re.sub(r'\s+', ' ', apellidos_raw).strip())
        if alias:
            entry_correo.delete(0, "end")
            entry_correo.insert(0, alias)

    entry_nombre.bind("<KeyRelease>",    autocompletar_correo)
    entry_apellidos.bind("<KeyRelease>", autocompletar_correo)

    def copiar_correo():
        alias  = entry_correo.get().strip().lower()
        correo = alias if "@" in alias else f"{alias}@primaprix.eu"
        app.clipboard_clear()
        app.clipboard_append(correo)
        app.update()
        resultado_label.configure(text=f"📋 Copiado: {correo}", text_color="#16a34a")

    def comprobar():
        nombre    = entry_nombre.get().strip()
        apellidos = entry_apellidos.get().strip()
        alias     = entry_correo.get().strip().lower()

        if "@" in alias:
            alias = alias.split("@")[0]

        correo = f"{alias}@primaprix.eu"

        if not nombre or not apellidos or not alias:
            mostrar_error("Debe completar todos los campos.")
            return

        # 🔄 Mostrar spinner
        spinner.pack(pady=5)
        spinner.start()
        resultado_label.configure(text="Buscando disponibilidad...", text_color="gray")

        sugerencias_box.configure(state="normal")
        sugerencias_box.delete("1.0", "end")
        sugerencias_box.configure(state="disabled")

        boton_copiar.pack_forget()

        def tarea():
            existe = correo_existe(correo)
            sugs   = generar_sugerencias(nombre, apellidos) if existe else []

            def actualizar_ui():
                spinner.stop()
                spinner.pack_forget()

                if existe:
                    resultado_label.configure(
                        text="❌ El correo ya existe en Entra ID",
                        text_color="#dc2626"
                    )

                    sugerencias_box.configure(state="normal")
                    if sugs:
                        sugerencias_box.insert("end", "Correos disponibles sugeridos:\n\n")
                        for s in sugs:
                            sugerencias_box.insert("end", f"• {s}\n")
                    else:
                        sugerencias_box.insert("end", "No se encontraron alternativas libres.")
                    sugerencias_box.configure(state="disabled")

                else:
                    resultado_label.configure(
                        text="✅ El correo está libre y se puede utilizar",
                        text_color="#16a34a"
                    )
                    boton_copiar.pack(pady=5)

            ventana.after(0, actualizar_ui)

        threading.Thread(target=tarea, daemon=True).start()

    ctk.CTkButton(
        ventana, text="Comprobar disponibilidad",
        height=40, width=250, corner_radius=20,
        fg_color="#2563eb", hover_color="#1d4ed8",
        command=comprobar
    ).pack(pady=10)

    resultado_label = ctk.CTkLabel(ventana, text="", font=("Segoe UI", 13, "bold"))
    resultado_label.pack(pady=10)
    spinner = ctk.CTkProgressBar(ventana, mode="indeterminate", width=250)

    boton_copiar = ctk.CTkButton(
        ventana, text="📋 Copiar correo",
        height=35, width=200, corner_radius=15,
        fg_color="#16a34a", hover_color="#15803d",
        command=copiar_correo
    )

    sugerencias_box = ctk.CTkTextbox(ventana, height=150, width=450)
    sugerencias_box.pack(pady=10)
    sugerencias_box.configure(state="disabled")


# ================================================================
# BUSCAR FICHERO ORIGEN
# ================================================================

def buscar_fichero_usuario(emp_id=None, correo=None):
    if emp_id and emp_id in INDEX_ID:
        return INDEX_ID[emp_id]
    if correo and correo.lower() in INDEX_CORREO:
        return INDEX_CORREO[correo.lower()]
    return None


def descargar_fichero(ruta_origen):
    if not ruta_origen or not os.path.exists(ruta_origen):
        mostrar_error("No se encontró fichero para descargar.")
        return
    from tkinter import filedialog
    nombre  = os.path.basename(ruta_origen)
    destino = filedialog.asksaveasfilename(
        initialfile=nombre, defaultextension=".txt", title="Guardar fichero como..."
    )
    if not destino:
        return
    try:
        import shutil
        shutil.copy2(ruta_origen, destino)
        mostrar_exito(f"Fichero descargado correctamente:\n{destino}")
    except Exception as e:
        mostrar_error(f"Error copiando fichero:\n{e}")


# ================================================================
# CREAR TABLA RESULTADO
# ================================================================

def crear_tabla(frame, datos):
    card = ctk.CTkFrame(frame, corner_radius=20, fg_color="white")
    card.pack(fill="x", padx=30, pady=20)

    nombre_completo = f"{datos['Nombre']} {datos['Apellidos']}"

    # ── Cabecera ──
    header_frame = ctk.CTkFrame(card, fg_color="transparent")
    header_frame.pack(fill="x", padx=25, pady=(20, 10))

    ctk.CTkLabel(
        header_frame, text=f"👤 {nombre_completo}",
        font=("Segoe UI", 22, "bold"), text_color="black"
    ).pack(side="left")

    emp_id         = str(datos.get("ID empleado") or "").strip()
    nombre_fichero = buscar_fichero_usuario(emp_id=emp_id, correo=datos.get("UPN"))
    texto_fichero  = f"📄 {nombre_fichero}" if nombre_fichero else "📄 No encontrado"

    label_fichero = ctk.CTkLabel(
        header_frame, text=texto_fichero,
        font=("Segoe UI", 12, "bold"), text_color="black", cursor="hand2"
    )
    label_fichero.pack(side="right", padx=5)

    upn     = datos.get("UPN", "")
    user_id = datos.get("id") or ""

    estado_ssff, id_ssff, nombre_ssff = comprobar_ssff(
        datos["UPN"], emp_id, datos.get("Correo personal")
    )

    # ── Tooltip "copiado" ──
    def mostrar_tooltip(widget, texto="Copiado"):
        widget.update_idletasks()
        tooltip = ctk.CTkToplevel(widget)
        tooltip.overrideredirect(True)
        tooltip.attributes("-topmost", True)
        tooltip.geometry(f"+{widget.winfo_rootx()}+{widget.winfo_rooty() - 28}")
        ctk.CTkLabel(
            tooltip, text=f"✓ {texto}",
            fg_color="#16a34a", text_color="white",
            corner_radius=8, font=("Segoe UI", 11, "bold"),
            padx=10, pady=2
        ).pack()
        tooltip.after(1200, tooltip.destroy)

    def copiar_valor(widget, texto):
        if not texto:
            return
        app.clipboard_clear()
        app.clipboard_append(texto)
        app.update()
        mostrar_tooltip(widget)

    label_fichero.bind("<Button-1>", lambda e: copiar_valor(label_fichero, nombre_fichero))

    # ── Info frame ──
    info_frame = ctk.CTkFrame(card, fg_color="transparent")
    info_frame.pack(fill="x", padx=25, pady=10, side="left")

    campos = [
        ("UPN",             datos["UPN"]),
        ("ID empleado",     emp_id),
        ("País",            datos["País o región"]),
        ("Correo personal", datos["Correo personal"]),
        ("Fecha creación",  datos["Fecha creación"]),
    ]

    def mostrar_info_ssff(estado_ssff, id_ssff, nombre_ssff, user_id, upn):
        ventana = ctk.CTkToplevel(app)
        ventana.title("Información SSFF")
        ancho, alto = 440, 450
        ventana.geometry(
            f"{ancho}x{alto}"
            f"+{app.winfo_x() + app.winfo_width()  // 2 - ancho // 2}"
            f"+{app.winfo_y() + app.winfo_height() // 2 - alto  // 2}"
        )
        ventana.transient(app)
        ventana.grab_set()
        ventana.focus()

        if estado_ssff is True:
            estado_texto, color = "Coincide con SSFF", "#16a34a"
        elif estado_ssff in (False, "id", "parecido"):
            estado_texto, color = "ID diferente en SSFF / Correo no actualizado", "#f59e0b"
        else:
            estado_texto, color = "Usuario no encontrado en SSFF", "#dc2626"

        contenedor = ctk.CTkFrame(ventana, corner_radius=12, fg_color="white")
        contenedor.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(
            contenedor,
            text=nombre_ssff if nombre_ssff else "Sin nombre SSFF",
            font=("Segoe UI", 22, "bold"), text_color="black", anchor="center"
        ).pack(pady=(15, 10))
        ctk.CTkFrame(contenedor, height=2, fg_color="#e5e7eb").pack(fill="x", padx=10, pady=(0, 15))
        ctk.CTkLabel(
            contenedor, text="Información SSFF",
            font=("Segoe UI", 16, "bold"), text_color="#111827"
        ).pack(pady=(0, 15))

        def campo_ssff(titulo, valor, color_valor="black"):
            f = ctk.CTkFrame(contenedor, fg_color="transparent")
            f.pack(fill="x", padx=10, pady=6)
            ctk.CTkLabel(f, text=f"{titulo}:", font=("Segoe UI", 12, "bold"),
                        text_color="#6b7280", width=120, anchor="w").pack(side="left")
            ctk.CTkLabel(f, text=valor if valor else "-", font=("Segoe UI", 13),
                        text_color=color_valor, anchor="w").pack(side="left")

        campo_ssff("Estado", estado_texto, color)

        if id_ssff:
            frame_id = ctk.CTkFrame(contenedor, fg_color="transparent")
            frame_id.pack(fill="x", padx=10, pady=6)
            ctk.CTkLabel(frame_id, text="ID SSFF:", font=("Segoe UI", 12, "bold"),
                        text_color="#6b7280", width=120, anchor="w").pack(side="left")
            ctk.CTkLabel(frame_id, text=id_ssff, font=("Segoe UI", 13), text_color="black").pack(side="left")

            if estado_ssff in (False, "id", "parecido"):
                def asignar_id():
                    from tkinter import messagebox
                    if not messagebox.askyesno("Confirmar cambio", f"¿Asignar ID SSFF {id_ssff} a este usuario?"):
                        return
                    try:
                        resp = graph_patch(
                            f"https://graph.microsoft.com/v1.0/users/{user_id}",
                            {"employeeId": str(id_ssff)}
                        )
                        if resp.status_code == 204:
                            mostrar_exito(f"ID actualizado a {id_ssff}")
                            for w in frame_resultados.winfo_children():
                                w.destroy()
                            crear_tabla(frame_resultados, buscar_usuario(upn))
                            ventana.destroy()
                        else:
                            mostrar_error(f"Error actualizando ID:\n{resp.text}")
                    except Exception as e:
                        mostrar_error(str(e))

                ctk.CTkButton(
                    frame_id, text="Asignar", width=90, height=28,
                    fg_color="#50a2d8", hover_color="#449ad3",
                    text_color="white", font=("Segoe UI", 11, "bold"),
                    corner_radius=8, command=asignar_id
                ).pack(side="left", padx=10)

        if nombre_ssff:
            campo_ssff("Nombre", nombre_ssff)

        correo_ssff = next(
            (c for c, d in SSFF_DATA.items() if d.get("nombre") == nombre_ssff), None
        )
        if correo_ssff and "@" in correo_ssff:
            campo_ssff("Correo", correo_ssff)
        else:
            campo_ssff("Correo", "Pendiente de actualizar", "#6b7280")

        ctk.CTkButton(contenedor, text="Cerrar", width=120, command=ventana.destroy).pack(pady=(20, 10))

    # ── Renderizar campos ──
    for titulo, valor in campos:
        fila = ctk.CTkFrame(info_frame, fg_color="transparent")
        fila.pack(fill="x", pady=4)

        ctk.CTkLabel(
            fila, text=f"{titulo}:", width=150, anchor="w",
            font=("Segoe UI", 12, "bold"), text_color="black"
        ).pack(side="left")

        valor_texto = valor if valor else ""

        if titulo == "ID empleado":
            if estado_ssff is True:
                color_circulo, tooltip_texto = "#16a34a", "Empleado encontrado en SSFF"
            elif estado_ssff in (False, "id", "parecido"):
                color_circulo, tooltip_texto = "#f59e0b", "Distinto en SSFF"
            else:
                color_circulo, tooltip_texto = "#dc2626", "Empleado no encontrado en SSFF"

            valor_label = ctk.CTkLabel(fila, text=valor_texto, anchor="w", text_color="black")
            valor_label.pack(side="left", padx=(0, 10))

            info_icon = ctk.CTkLabel(
                fila, text="ⓘ", font=("Segoe UI", 16, "bold"),
                text_color=color_circulo, cursor="hand2"
            )
            ToolTip(info_icon, tooltip_texto)
            info_icon.pack(side="left", padx=5)
            info_icon.bind(
                "<Button-1>",
                lambda e: mostrar_info_ssff(estado_ssff, id_ssff, nombre_ssff, user_id, datos["UPN"])
            )
        else:
            valor_label = ctk.CTkLabel(fila, text=valor_texto, anchor="w", text_color="black")
            valor_label.pack(side="left", padx=(0, 10))

        valor_label.bind("<Button-1>", lambda e, v=valor_texto, w=valor_label: copiar_valor(w, v))
        valor_label.configure(cursor="hand2")

    # ── Panel lateral accesos ──
    estado_panel = ctk.CTkFrame(
        card, fg_color="#f5f5f5", corner_radius=12,
        border_width=1, border_color="#d1d5db"
    )
    estado_panel.pack(side="right", padx=25, pady=15, fill="y")

    ctk.CTkLabel(
        estado_panel, text="📊 Estado Accesos",
        font=("Segoe UI", 14, "bold"), text_color="black"
    ).pack(pady=(10, 5))

    grupos_lower       = [str(g).lower() for g in datos.get("Grupos", [])]
    es_rrhh            = "rrhh primaprix" in grupos_lower
    propietarios_cache = _PROPIETARIOS_CACHE if _PROPIETARIOS_LISTO.is_set() else {}

    def crear_estado(fondo, texto, correcto=True, grupo_id=None, uid=None):
        color = "#16a34a" if correcto else "#dc2626"
        fila  = ctk.CTkFrame(fondo, fg_color="transparent")
        fila.pack(fill="x", padx=10, pady=6)
        ctk.CTkFrame(fila, width=16, height=16, corner_radius=8, fg_color=color).pack(side="left", padx=(0, 8))
        ctk.CTkLabel(fila, text=texto, anchor="w", text_color="black",
                    font=("Segoe UI", 12, "bold")).pack(side="left")

        if grupo_id and uid and propietarios_cache.get(grupo_id, False):
            def agregar(gid=grupo_id, u=uid):
                url     = f"https://graph.microsoft.com/v1.0/groups/{gid}/members/$ref"
                payload = {"@odata.id": f"https://graph.microsoft.com/v1.0/users/{u}"}
                try:
                    resp = graph_post(url, payload)
                    if resp.status_code in [200, 204]:
                        mostrar_exito("Usuario añadido correctamente al grupo.")
                        for w in frame_resultados.winfo_children():
                            w.destroy()
                        crear_tabla(frame_resultados, buscar_usuario(datos["UPN"]))
                    elif resp.status_code == 400 and "added object references already exist" in resp.text:
                        mostrar_error("El usuario ya pertenece al grupo.")
                    else:
                        mostrar_error(f"Error añadiendo usuario: {resp.status_code} {resp.text}")
                except Exception as e:
                    mostrar_error(f"Error al conectar con Graph: {e}")

            ctk.CTkButton(
                fila, text="+", width=50, height=35,
                font=("Segoe UI", 12, "bold"), corner_radius=15,
                fg_color="#3b82f6", hover_color="#2563eb", text_color="white",
                command=agregar
            ).pack(side="right", padx=10, pady=2)

    falta_sap = "sap success factors pro" not in grupos_lower

    bloque_ssff = ctk.CTkFrame(estado_panel, fg_color="transparent")
    bloque_ssff.pack(fill="x", padx=10, pady=6)
    badge_color = "#16a34a" if (not falta_sap and emp_id != "") else "#dc2626"
    ctk.CTkFrame(bloque_ssff, width=16, height=16, corner_radius=8, fg_color=badge_color).pack(side="left", padx=(0, 8))
    ctk.CTkLabel(bloque_ssff, text="Acceso a SSFF + UKG",
                font=("Segoe UI", 12, "bold"), text_color="black").pack(side="left")

    if falta_sap:
        fila = ctk.CTkFrame(estado_panel, fg_color="transparent")
        fila.pack(fill="x", padx=35, pady=2)
        ctk.CTkLabel(fila, text="• Grupo SAP Success Factors PRO",
                    font=("Segoe UI", 11), text_color="black").pack(side="left")

        if propietarios_cache.get(GRUPOS["sap success factors pro"], False) or es_rrhh:
            def agregar_sap():
                url     = f"https://graph.microsoft.com/v1.0/groups/{GRUPOS['sap success factors pro']}/members/$ref"
                payload = {"@odata.id": f"https://graph.microsoft.com/v1.0/users/{user_id}"}
                resp    = graph_post(url, payload)
                if resp.status_code in (200, 204):
                    mostrar_exito("Usuario añadido a SAP Success Factors")
                    for w in frame_resultados.winfo_children():
                        w.destroy()
                    crear_tabla(frame_resultados, buscar_usuario(datos["UPN"]))
                else:
                    mostrar_error(resp.text)

            ctk.CTkButton(
                fila, text="+", width=50, height=35,
                font=("Segoe UI", 12, "bold"), corner_radius=15,
                fg_color="#3b82f6", hover_color="#2563eb", text_color="white",
                command=agregar_sap
            ).pack(side="right", padx=10, pady=2)

    if emp_id == "":
        fila = ctk.CTkFrame(estado_panel, fg_color="transparent")
        fila.pack(fill="x", padx=35, pady=2)
        ctk.CTkLabel(fila, text="• ID empleado", font=("Segoe UI", 11), text_color="black").pack(side="left")

    crear_estado(estado_panel, "Acceso a MFA",
                correcto="empleados_mfa" in grupos_lower,
                grupo_id=GRUPOS["empleados_mfa"] if "empleados_mfa" not in grupos_lower else None,
                uid=user_id)

    crear_estado(estado_panel, "Acceso a VPN",
                correcto="vpn" in grupos_lower,
                grupo_id=GRUPOS["vpn"] if "vpn" not in grupos_lower else None,
                uid=user_id)

    # ── Listado grupos ──
    grupos_card = ctk.CTkFrame(frame, corner_radius=20, fg_color="white")
    grupos_card.pack(fill="both", expand=True, padx=30, pady=(0, 20))

    ctk.CTkLabel(
        grupos_card, text=f"👥 Grupos ({len(datos['Grupos'])})",
        font=("Segoe UI", 18, "bold"), text_color="black"
    ).pack(anchor="w", padx=25, pady=10)

    entry_buscar_grupo = ctk.CTkEntry(grupos_card, placeholder_text="Buscar grupo...", height=30)
    entry_buscar_grupo.pack(fill="x", padx=25, pady=(0, 10))

    scroll = ctk.CTkScrollableFrame(grupos_card, height=250, fg_color="white")
    scroll.pack(fill="both", expand=True, padx=25, pady=(0, 20))

    grupo_labels = []
    for g in datos["Grupos"]:
        lbl = ctk.CTkLabel(scroll, text=f"• {g}", anchor="w", text_color="black")
        lbl.pack(fill="x", pady=3)
        grupo_labels.append((g.lower(), lbl))

    def filtrar_grupos(event=None):
        texto = entry_buscar_grupo.get().strip().lower()
        for nombre, lbl in grupo_labels:
            if texto in nombre:
                lbl.pack(fill="x", pady=3)
            else:
                lbl.pack_forget()

    entry_buscar_grupo.bind("<KeyRelease>", filtrar_grupos)


# ================================================================
# EJECUTAR BÚSQUEDA
# ================================================================

def ejecutar_busqueda():
    alias    = entry.get().strip()
    busqueda = f"{alias}@primaprix.eu"
    if not alias:
        mostrar_error("Ingrese UPN, ID empleado o alias de correo")
        return

    boton_buscar.configure(state="disabled")
    for widget in frame_resultados.winfo_children():
        widget.destroy()

    spinner = ctk.CTkProgressBar(frame_resultados, mode="indeterminate")
    spinner.pack(padx=100, pady=80)
    spinner.start()

    def tarea():
        resultado = buscar_usuario(busqueda)

        def actualizar_ui():
            spinner.destroy()
            boton_buscar.configure(state="normal")
            if "error" in resultado:
                ctk.CTkLabel(
                    frame_resultados, text=resultado["error"],
                    text_color="#cc0000", font=("Segoe UI", 14, "bold")
                ).pack(pady=40)
            else:
                crear_tabla(frame_resultados, resultado)

        app.after(0, actualizar_ui)

    threading.Thread(target=tarea, daemon=True).start()


# ================================================================
# VENTANA PRINCIPAL + SPLASH
# ================================================================

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Consulta Usuario Entra ID")

ancho_inicial = 860
alto_inicial  = 680
app.geometry(f"{ancho_inicial}x{alto_inicial}")
app.minsize(ancho_inicial, 600)
app.configure(fg_color="#f2f4f7")

# Centrar ventana
app.update_idletasks()
app.geometry(
    f"{ancho_inicial}x{alto_inicial}"
    f"+{app.winfo_screenwidth()  // 2 - ancho_inicial // 2}"
    f"+{app.winfo_screenheight() // 2 - alto_inicial  // 2}"
)

# ── SPLASH: visible en <200 ms, mientras los imports pesan cargados ──
splash_frame = ctk.CTkFrame(app, fg_color="#f2f4f7")
splash_frame.place(relx=0, rely=0, relwidth=1, relheight=1)

ctk.CTkLabel(
    splash_frame,
    text="🔐 Consulta Usuario Entra ID",
    font=("Segoe UI", 32, "bold"), text_color="black"
).pack(expand=True, pady=(0, 10))

splash_sub = ctk.CTkLabel(
    splash_frame, text="Iniciando...",
    font=("Segoe UI", 13), text_color="gray"
)
splash_sub.pack(pady=(0, 20))

splash_progress = ctk.CTkProgressBar(splash_frame, mode="indeterminate", width=300)
splash_progress.pack(pady=(0, 80))
splash_progress.start()

app.update()  # Forzar render inmediato del splash

# ── Arrancar imports pesados en background AHORA ──
_thread_imports = threading.Thread(target=_cargar_imports_pesados, daemon=True)
_thread_imports.start()

# Placeholders (se asignan en _construir_ui)
entry            = None
boton_buscar     = None
frame_resultados = None


def _construir_ui():
    """Destruye el splash y construye la UI real tras el login."""
    global entry, boton_buscar, frame_resultados

    splash_frame.destroy()

    # HEADER
    frame_top = ctk.CTkFrame(app, corner_radius=0, fg_color="transparent")
    frame_top.pack(fill="x")
    ctk.CTkLabel(
        frame_top, text="🔐 Consulta Usuario Entra ID",
        font=("Segoe UI", 32, "bold"), text_color="black"
    ).pack(pady=(25, 5))
    ctk.CTkLabel(
        frame_top, text=f"Sesión iniciada como {USUARIO_LOGADO}",
        font=("Segoe UI", 13), text_color="gray"
    ).pack(pady=(0, 25))

    # BÚSQUEDA
    frame_busqueda = ctk.CTkFrame(app, fg_color="white", corner_radius=20)
    frame_busqueda.pack(pady=10, padx=30)

    entry = ctk.CTkEntry(frame_busqueda, height=40, width=260, placeholder_text="Alias o ID")
    entry.pack(side="left", padx=(15, 5), pady=15)
    entry.bind("<Return>", lambda event: ejecutar_busqueda())

    ctk.CTkLabel(
        frame_busqueda, text="@primaprix.eu",
        font=("Segoe UI", 14, "bold"), text_color="gray"
    ).pack(side="left", padx=(0, 10))

    boton_buscar = ctk.CTkButton(
        frame_busqueda, text="Buscar",
        height=40, corner_radius=20, command=ejecutar_busqueda
    )
    boton_buscar.pack(side="left", padx=5)

    if PUEDE_GENERAR_CORREO:
        ctk.CTkButton(
            frame_busqueda, text="📧 Correo Nuevo",
            height=40, corner_radius=20,
            fg_color="#85b6ee", hover_color="#72abec",
            font=("Segoe UI", 13, "bold"),
            command=abrir_ventana_generar_correo
        ).pack(side="left", padx=(10, 15))

    # RESULTADOS
    frame_resultados = ctk.CTkScrollableFrame(app, fg_color="transparent")
    frame_resultados.pack(fill="both", expand=True, padx=20, pady=20)

    # Actualizar globals para que ejecutar_busqueda los encuentre
    globals()["entry"]            = entry
    globals()["boton_buscar"]     = boton_buscar
    globals()["frame_resultados"] = frame_resultados


def _continuar_tras_imports():
    """
    Polling cada 50 ms hasta que los imports estén listos.
    Luego hace el login y construye la UI.
    """
    if not _imports_listos.is_set():
        app.after(50, _continuar_tras_imports)
        return

    splash_progress.stop()
    splash_sub.configure(text="Iniciando sesión...")
    app.update()

    if not login():
        app.destroy()
        sys.exit()

    _construir_ui()


# Arrancar el flujo tras el primer ciclo del event-loop
app.after(10, _continuar_tras_imports)

app.mainloop()