# Entra ID User Manager

📧 **Herramienta para gestión de usuarios y correos en Entra ID / Microsoft 365**  
Construida con **Python 3.10+**, **CustomTkinter** y la API de **Microsoft Graph**.

---

## Características principales

- ✅ Autenticación segura mediante **MSAL** y caché de tokens persistente
- 📂 Descarga y gestión de ficheros de origen desde **SharePoint**
- 🧾 Indexación de usuarios con información de SSFF y Entra ID
- 📋 Comprobación de disponibilidad de correos y generación de sugerencias automáticas
- 👥 Validación de pertenencia a grupos específicos antes de permitir acciones
- 🔄 Carga de datos y módulos pesados en **background** para una interfaz rápida
- 🖥 Interfaz moderna con **CustomTkinter**, incluyendo tooltips y ventanas de detalle

---

## Requisitos

- Python 3.10 o superior
- Windows (recomendado — la caché de tokens usa protección nativa de Windows)
- Una aplicación registrada en **Azure Entra ID** con los permisos necesarios

### Permisos de Microsoft Graph requeridos

| Permiso | Tipo |
|---|---|
| `User.Read.All` | Delegado |
| `GroupMember.Read.All` | Delegado |
| `Group.ReadWrite.All` | Delegado |
| `User.ReadWrite.All` | Delegado |
| `Files.Read.All` | Delegado |

---

## Instalación

### 1. Clona el repositorio

```bash
git clone https://github.com/tu_usuario/entra-id-user-manager.git
cd entra-id-user-manager
```

### 2. Crea un entorno virtual (recomendado)

```bash
python -m venv venv
venv\Scripts\activate        # Windows
# source venv/bin/activate   # macOS/Linux
```

### 3. Instala las dependencias

```bash
pip install -r requirements.txt
```

### 4. Configura las variables de entorno

Copia el fichero de ejemplo y rellena tus valores:

```bash
copy .env.example .env      # Windows
# cp .env.example .env      # macOS/Linux
```

Edita `.env` con los datos de tu organización:

```env
CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
SHAREPOINT_HOST=tuempresa.sharepoint.com
SHAREPOINT_SITE=/sites/tu-sitio
SHAREPOINT_LIBRARY=Nombre de tu biblioteca
GRUPO_REQUERIDO=ConsultaUsuarios
GRUPO_ID_SAP=...
GRUPO_ID_MFA=...
GRUPO_ID_VPN=...
GRUPO_ID_RRHH=...
CORREOS_AUTORIZADOS=admin@tuempresa.com,otro@tuempresa.com
```

> ⚠️ **El fichero `.env` nunca debe subirse a Git.** Está incluido en `.gitignore`.

### 5. Ejecuta la aplicación

```bash
python ConsultaUsuarios.py
```

---

## Estructura del proyecto

```
entra-id-user-manager/
├── ConsultaUsuarios.py   # Aplicación principal
├── requirements.txt      # Dependencias
├── .env.example          # Plantilla de configuración (sin valores reales)
├── .gitignore
├── README.md
├── cache/                # Caché local de datos (generada automáticamente, no en Git)
└── token_caches/         # Caché de tokens MSAL (generada automáticamente, no en Git)
```

---

## Uso

Al iniciar, la aplicación mostrará una pantalla de carga mientras se autentican los módulos en segundo plano. El usuario debe pertenecer al grupo de Entra ID definido en `GRUPO_REQUERIDO` para acceder.

Una vez autenticado podrás:
- Buscar usuarios por nombre, ID o correo
- Consultar su pertenencia a grupos
- Comprobar disponibilidad de alias de correo
- Gestionar altas en grupos (si tienes permisos de propietario)

---

## Contribuir

1. Haz un fork del repositorio
2. Crea una rama para tu cambio: `git checkout -b feature/mi-mejora`
3. Haz commit de tus cambios: `git commit -m 'Añade mi mejora'`
4. Haz push a tu rama: `git push origin feature/mi-mejora`
5. Abre un Pull Request

---

## Licencia

MIT — consulta el fichero `LICENSE` para más detalles.