import requests
import pandas as pd
import os
from requests.auth import HTTPBasicAuth
from flask import Flask, jsonify, render_template, request, redirect, url_for, flash, send_file
from io import BytesIO
from dotenv import load_dotenv 

# ---------------------------
# Cargar variables
# ---------------------------

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY")

# ---------------------------
# Configuración API 
# ---------------------------

BASE_URL = os.getenv("BASE_URL")
API_KEY = os.getenv("API_KEY")
PASSWORD = os.getenv("PASSWORD")

# ------------- FUNCIONES -----------------------

# ---------------------------
# Función para obtener roles
# ---------------------------

def get_roles():
    url = f"{BASE_URL}/contentRoles"
    params = {
        "itemsPerPage": 200, 
        "page": 1
    }
    roles = []
    while True:
        response = requests.get(url, auth=HTTPBasicAuth(API_KEY, PASSWORD), params=params, verify=False)
        if response.status_code != 200:
            print(f"Error al obtener roles: {response.status_code} - {response.text}")
            break

        data = response.json()
        if not data:  
            break

        for role in data:
            roles.append({
                "id": role.get("roleId"),
                "name": role.get("contentRole")
            })

        if len(data) < params["itemsPerPage"]:
            break
        params["page"] += 1

    return roles


# ---------------------------
# Función para obtener todos los usuarios
# ---------------------------
def get_all_users(items_per_page=200):
    users = []
    page = 1

    # Definir niveles de acceso
    niveles_permitidos = [4, 7]

    while True:
        url = f"{BASE_URL}/users"
        params = {
            "itemsPerPage": items_per_page,
            "page": page
        }
        response = requests.get(url, params=params, auth=HTTPBasicAuth(API_KEY, PASSWORD), verify=False)

        if response.status_code != 200:
            print(f"Error al obtener usuarios: {response.status_code} - {response.text}")
            break

        data = response.json()
        if isinstance(data, list):
            user_list = data
        elif isinstance(data, dict):
            user_list = data.get("data", [])
        else:
            user_list = []

        if not user_list:
            break

        for user in user_list:
            if user.get("accessLevel") in niveles_permitidos:
                users.append({
                    "userId": user.get("userId"),
                    "username": user.get("username"),
                    "firstName": user.get("firstName"),
                    "lastName": user.get("lastName"),
                    "email": user.get("email"),
                    "accessLevel": user.get("accessLevel"),
                    "accessLevelName": user.get("accessLevelName"),
                    "isActive": "Activo" if user.get("isActive", True) else "Inactivo",
                    "Fecha de Inicio": user.get("hireDate"),
                    "Fecha de Inicio 1": user.get("startDate"),
                    "Fecha de Expiración": user.get("expireDate"),
                    "Location Id": user.get("locationId"),
                    "Location Name": user.get("locationName")
                })

        page += 1

    return users

# ---------------------------
# Función para asignar rol a usuario
# ---------------------------

def assign_role(user_id, role_ids, expire_date=None):
    url = f"{BASE_URL}/users"
    payload = {
        "userId": int(user_id),
        "contentRoleAdd": [int(r) for r in role_ids]  
    }
    response = requests.put(url, json=payload, auth=HTTPBasicAuth(API_KEY, PASSWORD), verify=False)
    return response


# ---------------------------
# Función para fecha de expiración
# ---------------------------
def set_account_expiration(user_id, expire_date=None):
    payload = {"userId": int(user_id)}
    if expire_date:
        payload["expireDate"] = f"{expire_date}T23:59:59Z"
    else:
        payload["expireDate"] = None 

    url = f"{BASE_URL}/users"
    response = requests.put(url, json=payload, auth=HTTPBasicAuth(API_KEY, PASSWORD), verify=False)
    return response

# ---------------------------
# Función cambiar de estado a usuarios
# ---------------------------
def cambiar_estado_usuarios(archivo, estado_objetivo):
    try:
        df = pd.read_excel(archivo)
    except Exception:
        return None, None, None, "Error al leer el archivo Excel. Asegúrate que sea válido."

    if "userId" not in df.columns:
        return None, None, None, "El archivo debe contener una columna llamada 'userId'."

    activados = []
    ya_en_estado = []
    errores = []

    for user_id in df["userId"].dropna().astype(int).tolist():
        # Consultar estado actual
        url_get = f"{BASE_URL}/users/{user_id}"
        resp_get = requests.get(url_get, auth=HTTPBasicAuth(API_KEY, PASSWORD), verify=False)

        if resp_get.status_code != 200:
            errores.append((user_id, "No se pudo consultar"))
            continue

        user_data = resp_get.json()
        if user_data.get("isActive") == estado_objetivo:
            ya_en_estado.append(user_id)
            continue

        # Cambiar estado
        payload = {"userId": user_id, "isActive": estado_objetivo}
        url_put = f"{BASE_URL}/users"
        resp_put = requests.put(url_put, auth=HTTPBasicAuth(API_KEY, PASSWORD), json=payload, verify=False)

        if resp_put.status_code == 200:
            activados.append(user_id)
        else:
            errores.append((user_id, resp_put.text))

    return activados, ya_en_estado, errores, None

# ---------------------------
# Función para actualizar usuarios (firstName o middleName)
# ---------------------------
def update_users_to_corporate(archivo):
    try:
        df = pd.read_excel(archivo)
    except Exception:
        return [], ["Error al leer el archivo Excel. Asegúrate que sea válido."]

    # Validamos solo columnas mínimas obligatorias
    if not all(col in df.columns for col in ["userId", "firstName", "lastName"]):
        return [], ["El archivo debe contener al menos las columnas 'userId', 'firstName' y 'lastName'."]

    # Si no existe la columna middleName, la creamos vacía
    if "middleName" not in df.columns:
        df["middleName"] = ""

    actualizados = []
    errores = []

    for _, row in df.iterrows():
        user_id = int(row["userId"])
        first_name = str(row["firstName"]).strip().lower()
        middle_name = str(row["middleName"]).strip().lower() if row["middleName"] else ""
        last_name = str(row["lastName"]).strip().lower()

        # Construcción del correo
        email = f"{first_name}.{last_name}@unacem.ec"

        # Reemplazar caracteres especiales
        email = (
            email.replace(" ", "")
                 .replace("á","a").replace("é","e")
                 .replace("í","i").replace("ó","o").replace("ú","u")
        )

        # Verificar si ya existe en esta misma corrida
        if any(a[1] == email for a in actualizados):
            if middle_name:  
                email = f"{middle_name}.{last_name}@unacem.ec"
            else:  
                base = f"{first_name}.{last_name}"
                i = 1
                while any(a[1] == f"{base}{i}@unacem.ec" for a in actualizados):
                    i += 1
                email = f"{base}{i}@unacem.ec"

        payload = {
            "userId": user_id,
            "email": email
        }

        url = f"{BASE_URL}/users"
        resp = requests.put(url, json=payload, auth=HTTPBasicAuth(API_KEY, PASSWORD), verify=False)

        if resp.status_code == 200:
            actualizados.append((user_id, email))
        else:
            errores.append((user_id, resp.text))

    return actualizados, errores

# ---------------------------
# Renombrar campos de usuarios inactivos
# ---------------------------

def renombrar_usuarios(file):
    try:
        df = pd.read_excel(file)

        # Validación de columnas mínimas
        if "userId" not in df.columns or "isActive" not in df.columns:
            return [], ["El archivo no contiene las columnas necesarias (userId, isActive)"]

        actualizados = []
        errores = []
        contador = 1

        for _, row in df.iterrows():
            try:
                if str(row["isActive"]).strip().lower() == "inactivo":
                    user_id = int(row["userId"])

                    nuevo_username = f"disponible{contador}"
                    nuevo_email = f"disponible{contador}@unacem.ec"
                    nuevo_nombre = f"Disponible{contador}"

                    payload = {
                        "userId": user_id,                
                        "username": nuevo_username,
                        "email": nuevo_email,
                        "firstName": nuevo_nombre,
                        "middleName": nuevo_nombre,
                        "lastName": nuevo_nombre,
                        "lockUsernamePassword": True     
                    }

                    url = f"{BASE_URL}/users"
                    response = requests.put(
                        url,
                        json=payload,
                        auth=HTTPBasicAuth(API_KEY, PASSWORD),
                        verify=False
                    )

                    if response.status_code == 200:
                        actualizados.append(user_id)
                        contador += 1
                    else:
                        errores.append(
                            f"Error {response.status_code} en {user_id}: {response.text}"
                        )

            except Exception as e:
                errores.append(f"Error procesando usuario {row.get('userId')}: {str(e)}")

        return actualizados, errores

    except Exception as e:
        return [], [f"Error leyendo archivo: {str(e)}"]


# ---------------------------
# Función para crear usuarios
# ---------------------------
def crear_usuarios(archivo, access_level=7, location_id=137980, default_password="Temp123"):
    try:
        df = pd.read_excel(archivo)
    except Exception:
        return [], ["Error al leer el archivo Excel. Asegúrate que sea válido."]

    if not all(col in df.columns for col in ["Empleados (Apellidos)", "Empleados (Nombres)"]):
        return [], ["El archivo debe contener las columnas 'Empleados (Apellidos)' y 'Empleados (Nombres)'."]

    creados = []
    errores = []

    def normalize(s):
        """ Normaliza cadenas para correos y usernames """
        return (
            str(s).strip()
            .replace(" ", "")
            .replace("á", "a").replace("Á", "a")
            .replace("é", "e").replace("É", "e")
            .replace("í", "i").replace("Í", "i")
            .replace("ó", "o").replace("Ó", "o")
            .replace("ú", "u").replace("Ú", "u")
            .replace("ñ", "n").replace("Ñ", "n")
        )

    for _, row in df.iterrows():
        try:
            apellidos_raw = str(row["Empleados (Apellidos)"]).strip()
            nombres_raw = str(row["Empleados (Nombres)"]).strip()

            if not apellidos_raw or not nombres_raw:
                errores.append(("??", "Apellidos o Nombres vacíos"))
                continue

            # Separar nombres y apellidos
            partes_nombre = nombres_raw.split()
            real_first_name = partes_nombre[0].title()
            middle_name = " ".join(partes_nombre[1:]).title() if len(partes_nombre) > 1 else ""
            apellidos = apellidos_raw.title()
            primer_apellido = apellidos.split()[0].title()

            # Generar username y correo
            username = f"{real_first_name[0]}{primer_apellido}".upper()
            email_local = f"{real_first_name.lower()}.{primer_apellido.lower()}"
            email_local = normalize(email_local)
            email = f"{email_local}@unacem.ec"

            payload = {
                "username": username,
                "email": email,
                "firstName": real_first_name,
                "middleName": middle_name if middle_name else None,
                "lastName": apellidos,
                "isActive": True,
                "lockUsernamePassword": True,
                "password": default_password,
                "accessLevel": int(access_level),
                "locationId": int(location_id)
            }

            # 🔹 Crear usuario directamente
            url = f"{BASE_URL}/users"
            resp = requests.post(url, json=payload, auth=HTTPBasicAuth(API_KEY, PASSWORD), verify=False)
            if resp.status_code in (200, 201):
                try:
                    body = resp.json()
                    created_id = body.get("userId") or body.get("id") or username
                except Exception:
                    created_id = username
                creados.append((created_id, username, email))
            else:
                try:
                    j = resp.json()
                    if "errors" in j and isinstance(j["errors"], list):
                        msgs = ", ".join(e.get("message", str(e)) for e in j["errors"])
                    else:
                        msgs = j.get("message") or str(j)
                except Exception:
                    msgs = resp.text
                errores.append((username, msgs))

        except Exception as e:
            errores.append((row.get("Empleados (Nombres)", "??"), str(e)))

    return creados, errores

# ---------------------------
# Función resetear contraseñas de usuarios
# ---------------------------
def resetear_passwords_masivo(df, new_password="Temp1234"):
    actualizados = []
    errores = []

    for user_id in df["userId"].dropna().astype(int).tolist():
        try:
            # Consultar usuario
            url_get = f"{BASE_URL}/users/{user_id}"
            resp_get = requests.get(url_get, auth=HTTPBasicAuth(API_KEY, PASSWORD), verify=False)

            if resp_get.status_code != 200:
                errores.append((user_id, "No se pudo consultar"))
                continue

            user_data = resp_get.json()
            username = user_data.get("username")

            payload = {
                "userId": user_id,
                "username": username,
                "password": new_password,
                "locationId": user_data.get("locationId"),
                "lockUsernamePassword": True,
                "forcePasswordUpdate": True,
                "isActive": True
            }

            url_put = f"{BASE_URL}/users"
            resp_put = requests.put(url_put, json=payload, auth=HTTPBasicAuth(API_KEY, PASSWORD), verify=False)

            if resp_put.status_code == 200:
                actualizados.append(user_id)
            else:
                errores.append((user_id, resp_put.text))

        except Exception as e:
            errores.append((user_id, str(e)))

    return actualizados, errores   # <-- 🔹 SOLO DOS VALORES



# --------------------------------------------------- RUTAS ---------------------------------------------

# ---------------------------
# Ruta para exportar usuarios a Excel
# ---------------------------
@app.route("/export_users")
def export_users():
    users = get_all_users()
    if not users:
        flash("No se encontraron usuarios.", "danger")
        return redirect(url_for("index"))

    df = pd.DataFrame(users)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Usuarios")
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="Usuarios.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------------------------
# Ruta para activar/inactivar usuarios
# ---------------------------
@app.route("/activar_usuarios", methods=["GET", "POST"])
def activar_usuarios():
    if request.method == "POST":
        archivo = request.files["archivo"]
        accion = request.form.get("accion")

        if not archivo:
            flash("Debes subir un archivo Excel", "danger")
            return redirect(url_for("activar_usuarios"))

        # Determinar si activar o inactivar
        estado_objetivo = True if accion == "activar" else False

        activados, ya_en_estado, errores, error_msg = cambiar_estado_usuarios(archivo, estado_objetivo)

        if error_msg:
            flash(error_msg, "danger")
        else:
            if estado_objetivo:  # Activar
                flash(f"Usuarios activados: {len(activados)} ({activados})", "success")
                flash(f"Usuarios ya estaban activos: {len(ya_en_estado)} ({ya_en_estado})", "info")
            else:  # Inactivar
                flash(f"Usuarios inactivados: {len(activados)} ({activados})", "warning")
                flash(f"Usuarios ya estaban inactivos: {len(ya_en_estado)} ({ya_en_estado})", "info")

            flash(f"Errores: {len(errores)} ({errores})", "danger")

        return redirect(url_for("activar_usuarios"))

    return render_template("activar_usuarios.html")

# ---------------------------
# Ruta para actualizar emails a corporativos
# ---------------------------
@app.route("/actualizar_usuarios", methods=["GET", "POST"])
def actualizar_usuarios():
    if request.method == "POST":
        archivo = request.files["archivo"]

        if not archivo:
            flash("Debes subir un archivo Excel", "danger")
            return redirect(url_for("actualizar_usuarios"))

        actualizados, errores = update_users_to_corporate(archivo)

        if errores:
            flash(f"Usuarios actualizados: {len(actualizados)}", "success")
            flash(f"Errores en {len(errores)} usuarios: {errores}", "danger")
        else:
            flash(f"Todos los usuarios fueron actualizados correctamente: {len(actualizados)}", "success")

        return redirect(url_for("actualizar_usuarios"))

    return render_template("actualizar_usuarios.html")


# ---------------------------
# Página de carga de Excel para asignar roles
# ---------------------------

@app.route("/roles", methods=["GET", "POST"])
def roles():
    roles = sorted(get_roles(), key=lambda r: r["id"], reverse=True)


    if request.method == "POST":
        role_ids = request.form.getlist("role_id")  
        expire_date = request.form.get("expire_date") 
        file = request.files.get("archivo")

        if not role_ids:
            flash("Debe seleccionar al menos un rol.", "danger")
            return redirect(url_for("roles"))

        if not file:
            flash("Debe cargar un archivo Excel.", "danger")
            return redirect(url_for("roles"))

        try:
            df = pd.read_excel(file)
        except Exception:
            flash("Error al leer el archivo Excel. Asegúrate que sea válido.", "danger")
            return redirect(url_for("roles"))

        if "userId" not in df.columns:
            flash("El archivo debe contener una columna llamada 'userId'.", "danger")
            return redirect(url_for("roles"))

        errors = []
        success_count = 0

        for user_id in df["userId"].dropna().astype(int):
            try:
                resp_role = assign_role(user_id, role_ids)
                if resp_role.status_code != 200:
                    errors.append(f"Error asignando roles a usuario {user_id}: {resp_role.text}")

                if expire_date:
                    resp_exp = set_account_expiration(user_id, expire_date)
                    if resp_exp.status_code != 200:
                        errors.append(f"Error actualizando expiración de usuario {user_id}: {resp_exp.text}")

                success_count += 1

            except Exception as e:
                errors.append(f"Error procesando usuario {user_id}: {str(e)}")

        if errors:
            flash(f"Usuarios procesados correctamente: {success_count}. Errores: {len(errors)}", "warning")
        else:
            flash(f"Todos los usuarios fueron procesados correctamente: {success_count}", "success")

        return redirect(url_for("roles"))

    return render_template("roles.html", roles=roles)

# ---------------------------
# Página principal: Gestión de Usuarios
# ---------------------------

@app.route("/gestion_usuarios")
def gestion_usuarios():
    return render_template("gestion_usuarios.html")


# ---------------------------
# Ruta para Renombrar Usuarios
# ---------------------------
@app.route("/renombrar_usuarios", methods=["GET", "POST"])
def anonymize_users():
    if request.method == "POST":
        archivo = request.files["archivo"]

        if not archivo:
            flash("Debes subir un archivo Excel", "danger")
            return redirect(url_for("anonymize_users"))

        actualizados, errores = renombrar_usuarios(archivo)

        if errores:
            flash(f"Usuarios actualizados: {len(actualizados)}", "success")
            flash(f"Errores en {len(errores)} usuarios: {errores}", "danger")
        else:
            flash(f"Todos los usuarios inactivos fueron anonimizados correctamente: {len(actualizados)}", "success")

        return redirect(url_for("anonymize_users"))

    return render_template("renombrar_usuarios.html")


# ---------------------------
# Ruta para crear usuarios
# ---------------------------
@app.route("/crear_usuarios", methods=["GET", "POST"])
def usuarios():
    if request.method == "POST":
        file = request.files.get("archivo")
        access_level = request.form.get("access_level", 7)  # lo elige el usuario

        if not file:
            flash("Debe cargar un archivo Excel.", "danger")
            return redirect(url_for("usuarios"))

        # Forzar valores 
        location_id = 137980
        default_password = "Temp123"

        creados, errores = crear_usuarios(
            file,
            access_level=access_level,
            location_id=location_id,
            default_password=default_password
        )

        # Mostrar resumen en pantalla
        if creados:
            usuarios_ok = ", ".join([u[1] for u in creados])  # u[1] = username
            flash(f"Usuarios creados exitosamente: {usuarios_ok}", "success")

        if errores:
            for usuario, msg in errores:
                flash(f"Error con {usuario}: {msg}", "danger")

        return redirect(url_for("usuarios"))

    # Valores por defecto para renderizar el formulario
    return render_template(
        "crear_usuarios.html",
        default_access=7,
        default_location=137980,
        default_password="Temp123"
    )


# ---------------------------
# Ruta: Resetear contraseñas masivo
# ---------------------------
@app.route("/resetear_passwords", methods=["GET", "POST"])
def resetear_passwords_route():
    if request.method == "GET":
        # Mostrar formulario
        return render_template("resetear_passwords.html")

    if "archivo" not in request.files:
        flash("No se subió ningún archivo.", "danger")
        return redirect(url_for("resetear_passwords_route"))

    archivo = request.files["archivo"]

    if archivo.filename == "":
        flash("El archivo está vacío.", "danger")
        return redirect(url_for("resetear_passwords_route"))

    try:
        df = pd.read_excel(archivo)
        actualizados, errores = resetear_passwords_masivo(df)

        flash(f"Se actualizaron {len(actualizados)} usuarios. Errores: {len(errores)}", "success")
    except Exception as e:
        flash(f"Error al procesar archivo: {e}", "danger")

    return redirect(url_for("resetear_passwords_route"))




# ---------------------------
# Página principal: Home Page
# ---------------------------
@app.route("/")
def home():
    return render_template("home.html")


# ---------------------------
# Ejecutar
# ---------------------------
if __name__ == "__main__":
    app.run(debug=True)
    