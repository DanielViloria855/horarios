import os
import calendar
import locale
from datetime import date, datetime, timedelta
from itertools import groupby
from urllib.parse import quote_plus
from flask import Flask, flash, redirect, render_template, send_file, url_for, request, jsonify
from config import Config
from utils import calcular_turno, estructurar_reporte, estructurar_reporte_con_columnas, generar_excel_con_columnas, generar_excel_tabla, generar_pdf_con_columnas, generar_pdf_tabla, generar_pdf_tabla, preparar_datos_reporte, generar_pdf_detallado, generar_excel_detallado
from extensions import db
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required
from flask_bcrypt import Bcrypt

# Configurar el idioma para fechas en español
try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
    except locale.Error:
        print("Advertencia: No se pudo configurar el locale a español.")

app = Flask(__name__)
app.config.from_object(Config)

# Inicialización de extensiones de login
bcrypt = Bcrypt(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'
login_manager.login_message = "Debes iniciar sesión para ver esta página."
login_manager.login_message_category = "warning"

db.init_app(app)

with app.app_context():
    import models
    Empleado = models.Empleado
    Turno = models.Turno
    ResumenMensual = models.ResumenMensual

# Clase de usuario y usuario "quemado"
class User(UserMixin):
    def __init__(self, id, email, password_hash):
        self.id = id
        self.email = email
        self.password_hash = password_hash

# Contraseña original: "waira2025restaurantebar"
hashed_password = bcrypt.generate_password_hash("waira2025").decode('utf-8')
user = User(id=1, email="adminwaira@danvilo.com", password_hash=hashed_password)

@login_manager.user_loader
def load_user(user_id):
    if int(user_id) == user.id:
        return user
    return None

# --- RUTAS DE AUTENTICACIÓN ---

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email')
        password = request.form.get('password')
        if email == user.email and bcrypt.check_password_hash(user.password_hash, password):
            login_user(user)
            next_page = request.args.get('next')
            return redirect(next_page) if next_page else redirect(url_for('horario_diario'))
        else:
            flash('Usuario o contraseña incorrectos.', 'danger')
    return render_template('login.html')

@app.route('/forgot-password')
def forgot_password():
    return render_template('forgot_password.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('Has cerrado sesión exitosamente.', 'success')
    return redirect(url_for('login'))

@app.route("/")
@login_required
def index():
    return redirect(url_for('horario_diario'))

@app.route("/empleados")
@login_required
def listar_empleados():
    empleados = Empleado.query.order_by(Empleado.nombre).all()
    return render_template("empleados_list.html", empleados=empleados)

@app.route("/turnos/nuevo", methods=["GET", "POST"])
@login_required
def nuevo_turno():
    if request.method == "POST":
        try:
            id_empleado = int(request.form["id_empleado"])
            fecha_str = request.form["fecha"]
            entrada_str = request.form["hora_entrada"]
            salida_str = request.form["hora_salida"]

            fecha = datetime.strptime(fecha_str, "%Y-%m-%d").date()
            hora_entrada = datetime.strptime(entrada_str, "%H:%M").time()
            hora_salida = datetime.strptime(salida_str, "%H:%M").time()

            res = calcular_turno(fecha, hora_entrada, hora_salida)
            horas_nocturnas = res["nocturnas_base"] + res["nocturnas_dia_siguiente"]
            horas_totales = res["total_horas"]
            horas_normales = round(horas_totales - horas_nocturnas, 2)

            turno = Turno(
                id_empleado=id_empleado,
                fecha=fecha,
                hora_entrada=hora_entrada,
                hora_salida=hora_salida,
                horas_normales=horas_normales,
                horas_nocturnas=round(horas_nocturnas, 2),
                horas_extras=0.0
            )

            db.session.add(turno)
            db.session.commit()
            
            flash("Turno guardado correctamente.", "success")
            return redirect(url_for("nuevo_turno", fecha_inicio=fecha_str, fecha_fin=fecha_str))
        except Exception as e:
            db.session.rollback()
            flash(f"Error guardando turno: {e}", "danger")
            return redirect(url_for("nuevo_turno"))

    fecha_inicio_str = request.args.get('fecha_inicio')
    fecha_fin_str = request.args.get('fecha_fin')
    empleado_id_str = request.args.get('id_empleado')

    turnos_detallados = []
    resumen_periodo = {
        "normales": 0.0, "nocturnas": 0.0, "extras": 0.0, "total": 0.0
    }
    
    if fecha_inicio_str and fecha_fin_str:
        try:
            fecha_inicio = datetime.strptime(fecha_inicio_str, "%Y-%m-%d").date()
            fecha_fin = datetime.strptime(fecha_fin_str, "%Y-%m-%d").date()
            
            query = Turno.query.filter(Turno.fecha.between(fecha_inicio, fecha_fin))
            
            if empleado_id_str and empleado_id_str.isdigit():
                query = query.filter_by(id_empleado=int(empleado_id_str))

            turnos_en_periodo = query.order_by(Turno.fecha).all()

            for turno in turnos_en_periodo:
                if not turno.hora_entrada:
                    continue
                    
                res = calcular_turno(turno.fecha, turno.hora_entrada, turno.hora_salida)
                
                total_turno = res["total_horas"]
                nocturnas_turno = res["nocturnas_base"] + res["nocturnas_dia_siguiente"]
                extras_turno = max(0, total_turno - 8)
                normales_turno = total_turno - nocturnas_turno - extras_turno

                resumen_periodo["nocturnas"] += nocturnas_turno
                resumen_periodo["extras"] += extras_turno
                resumen_periodo["normales"] += normales_turno
                resumen_periodo["total"] += total_turno
                
                turnos_detallados.append({
                    "id": turno.id_turno,
                    "empleado": turno.empleado.nombre,
                    "fecha": turno.fecha,
                    "entrada": turno.hora_entrada,
                    "salida": turno.hora_salida,
                    "normales": normales_turno,
                    "nocturnas": nocturnas_turno,
                    "extras": extras_turno,
                    "total": total_turno,
                    "observaciones": turno.observaciones
                })
        except ValueError:
            flash("Formato de fecha inválido.", "warning")

    empleados_activos = Empleado.query.filter_by(estado='Activo').order_by(Empleado.nombre).all()
    todos_los_empleados = Empleado.query.order_by(Empleado.nombre).all()

    return render_template(
        "turno_form.html", 
        empleados=empleados_activos,
        todos_los_empleados=todos_los_empleados,
        turnos_detallados=turnos_detallados,
        resumen_periodo=resumen_periodo,
        filtros=request.args
    )

@app.route("/ver-horario/<fecha>")
@login_required
def ver_horario(fecha):
    try:
        fecha_obj = datetime.strptime(fecha, "%Y-%m-%d").date()
    except ValueError:
        return "Formato de fecha inválido. Usa YYYY-MM-DD", 400

    turnos = Turno.query.filter_by(fecha=fecha_obj).join(Empleado).order_by(Empleado.nombre).all()
    return render_template("horario_limpio.html", fecha=fecha_obj, turnos=turnos)

@app.route("/generar_whatsapp/<fecha>")
@login_required
def generar_whatsapp(fecha):
    try:
        fecha_obj = datetime.strptime(fecha, "%Y-%m-%d").date()
    except ValueError:
        return "Formato de fecha inválido. Usa YYYY-MM-DD", 400

    # Definir el orden específico de los cargos
    orden_cargos = ['Caja', 'Meser@', 'Cocina', 'Bar', 'Oficios varios', 'Vigilante']
    
    empleados_activos = Empleado.query.filter_by(estado='Activo').order_by(Empleado.cargo, Empleado.nombre).all()
    
    # Agrupar empleados por cargo
    empleados_agrupados = {}
    for empleado in empleados_activos:
        cargo = empleado.cargo if empleado.cargo else 'Sin Cargo'
        if cargo not in empleados_agrupados:
            empleados_agrupados[cargo] = []
        empleados_agrupados[cargo].append(empleado)
    
    # Ordenar según el orden especificado
    empleados_ordenados = []
    for cargo in orden_cargos:
        if cargo in empleados_agrupados:
            empleados_ordenados.append((cargo, empleados_agrupados[cargo]))
    
    # Agregar los cargos que no están en la lista ordenada
    for cargo, empleados in empleados_agrupados.items():
        if cargo not in orden_cargos:
            empleados_ordenados.append((cargo, empleados))
    
    turnos_raw = Turno.query.filter_by(fecha=fecha_obj).all()
    turnos_del_dia = {t.id_empleado: t for t in turnos_raw}

    mensaje = f"*{fecha_obj.strftime('%A %d de %B del %Y').capitalize()}*\n\n"

    for cargo, empleados in empleados_ordenados:
        mensaje += f"*{cargo.upper() if cargo else 'SIN CARGO'}*\n"
        for empleado in empleados:
            turno = turnos_del_dia.get(empleado.id_empleado)
            nombre_corto = empleado.nombre.split()[0]

            if turno:
                if turno.observaciones:
                    mensaje += f"{nombre_corto} {turno.observaciones}\n"
                elif turno.hora_entrada and turno.hora_salida:
                    entrada_h = turno.hora_entrada.strftime("%I").lstrip('0')
                    salida_h = turno.hora_salida.strftime("%I").lstrip('0')
                    entrada_m = turno.hora_entrada.strftime("%M")
                    salida_m = turno.hora_salida.strftime("%M")
                    
                    # Formatear la hora para mostrar minutos solo si son diferentes de 00
                    entrada_str = f"{entrada_h}:{entrada_m}" if entrada_m != "00" else entrada_h
                    salida_str = f"{salida_h}:{salida_m}" if salida_m != "00" else salida_h
                    
                    mensaje += f"{nombre_corto} {entrada_str}-{salida_str}\n"
                else:
                     mensaje += f"{nombre_corto}\n"
            else:
                 mensaje += f"{nombre_corto}\n"
        mensaje += "\n"

    url = "https://web.whatsapp.com/send?text=" + quote_plus(mensaje)
    return redirect(url)

@app.route("/reportes")
@login_required
def vista_reportes():
    empleados = Empleado.query.order_by(Empleado.nombre).all()
    
    # Obtener año actual y rango de años (5 años atrás y adelante)
    current_year = datetime.now().year
    years = range(current_year - 5, current_year + 6)
    
    # Pasar la fecha actual al template
    now = datetime.now()
    
    return render_template("reportes.html", 
                         empleados=empleados,
                         years=years,
                         current_year=current_year,
                         now=now)

@app.route("/reportes/generar", methods=["POST"])
@login_required
def generar_reporte():
    try:
        tipo_periodo = request.form.get('tipo_periodo')
        mes = int(request.form.get('mes'))
        anio = int(request.form.get('anio'))
        formato = request.form.get('formato')
        empleado_id = request.form.get('id_empleado')
        tipo_datos = request.form.get('tipo_datos', 'totales')

        fecha_inicio, fecha_fin = None, None
        
        if tipo_periodo == 'personalizado':
            fecha_inicio_str = request.form.get('fecha_inicio')
            fecha_fin_str = request.form.get('fecha_fin')
            if not fecha_inicio_str or not fecha_fin_str:
                flash("Debes seleccionar un rango de fechas para el período personalizado.", "danger")
                return redirect(url_for('vista_reportes'))
            
            fecha_inicio = datetime.strptime(fecha_inicio_str, "%Y-%m-%d").date()
            fecha_fin = datetime.strptime(fecha_fin_str, "%Y-%m-%d").date()
        
        elif tipo_periodo == 'mensual':
            _, ultimo_dia = calendar.monthrange(anio, mes)
            fecha_inicio = date(anio, mes, 1)
            fecha_fin = date(anio, mes, ultimo_dia)
        
        elif tipo_periodo == 'quincenal':
            quincena = int(request.form.get('quincena'))
            if quincena == 1:
                fecha_inicio = date(anio, mes, 1)
                fecha_fin = date(anio, mes, 15)
            else:
                _, ultimo_dia = calendar.monthrange(anio, mes)
                fecha_inicio = date(anio, mes, 16)
                fecha_fin = date(anio, mes, ultimo_dia)
        
        elif tipo_periodo == 'semanal':
            semana_num = int(request.form.get('semana'))
            
            # Calcular fecha inicio y fin de semana (siempre lunes a domingo)
            cal = calendar.Calendar(firstweekday=0)  # Lunes como primer día
            month_weeks = cal.monthdatescalendar(anio, mes)
            
            if semana_num < 1 or semana_num > len(month_weeks):
                flash(f"La semana {semana_num} no existe en este mes.", "warning")
                return redirect(url_for('vista_reportes'))
            
            semana_seleccionada = month_weeks[semana_num - 1]
            fecha_inicio = semana_seleccionada[0]  # Primer día de la semana (lunes)
            fecha_fin = semana_seleccionada[6]     # Último día de la semana (domingo)

        # Verificar que las fechas son válidas
        if fecha_inicio is None or fecha_fin is None:
            flash("Error al calcular el período seleccionado.", "danger")
            return redirect(url_for('vista_reportes'))

        # Obtener datos y generar reporte
        datos = preparar_datos_reporte(fecha_inicio, fecha_fin, tipo_datos)

        if empleado_id and empleado_id.isdigit():
            empleado_id_int = int(empleado_id)
            if empleado_id_int in datos:
                datos = {empleado_id_int: datos[empleado_id_int]}
            else:
                datos = {}
        
        if not datos:
            flash("No se encontraron turnos para los filtros seleccionados.", "warning")
            return redirect(url_for('vista_reportes'))

        # Estructurar datos para el reporte con las nuevas columnas
        datos_estructurados = estructurar_reporte_con_columnas(datos, fecha_inicio, fecha_fin, tipo_periodo, tipo_datos)

        # Generar nombre de archivo
        periodo_nombre = f"{tipo_periodo}_{anio}_{mes}"
        if tipo_periodo == 'quincenal':
            periodo_nombre += f"_q{request.form.get('quincena')}"
        elif tipo_periodo == 'semanal':
            periodo_nombre += f"_s{request.form.get('semana')}"
        elif tipo_periodo == 'personalizado':
            periodo_nombre = f"personalizado_{fecha_inicio.strftime('%Y%m%d')}_{fecha_fin.strftime('%Y%m%d')}"
        
        periodo_nombre += f"_{tipo_datos}"
        filename = f"reporte_{periodo_nombre}"
        
        # Generar reporte en el formato solicitado
        if formato == 'pdf':
            buffer = generar_pdf_con_columnas(datos_estructurados)
            mimetype = "application/pdf"
            filename += ".pdf"
        else:
            buffer = generar_excel_con_columnas(datos_estructurados)
            mimetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            filename += ".xlsx"

        return send_file(buffer, as_attachment=True, download_name=filename, mimetype=mimetype)

    except Exception as e:
        flash(f"Ocurrió un error al generar el reporte: {e}", "danger")
        return redirect(url_for('vista_reportes'))

@app.route("/empleados/nuevo", methods=["GET", "POST"])
@login_required
def nuevo_empleado():
    if request.method == "POST":
        try:
            nombre = request.form["nombre"]
            documento = request.form["documento"]
            cargo = request.form["cargo"]
            fecha_ingreso_str = request.form["fecha_ingreso"]
            estado = request.form["estado"]

            fecha_ingreso = datetime.strptime(fecha_ingreso_str, "%Y-%m-%d").date()

            empleado = Empleado(
                nombre=nombre,
                documento=documento,
                cargo=cargo,
                fecha_ingreso=fecha_ingreso,
                estado=estado
            )
            
            db.session.add(empleado)
            db.session.commit()
            flash("Empleado registrado con éxito.", "success")
            return redirect(url_for("listar_empleados"))

        except Exception as e:
            db.session.rollback()
            if "UNIQUE constraint failed" in str(e) or "duplicate key value" in str(e):
                 flash("Error: El documento ya está registrado.", "danger")
            else:
                 flash(f"Error al registrar empleado: {e}", "danger")
            return redirect(url_for("nuevo_empleado"))
            
    return render_template("empleado_form.html")

@app.route("/empleados/editar/<int:id>", methods=["GET", "POST"])
@login_required
def editar_empleado(id):
    empleado = Empleado.query.get_or_404(id)

    if request.method == "POST":
        try:
            empleado.nombre = request.form["nombre"]
            empleado.documento = request.form["documento"]
            empleado.cargo = request.form["cargo"]
            empleado.fecha_ingreso = datetime.strptime(request.form["fecha_ingreso"], "%Y-%m-%d").date()
            empleado.estado = request.form["estado"]
            
            db.session.commit()
            flash("Empleado actualizado con éxito.", "success")
            return redirect(url_for("listar_empleados"))

        except Exception as e:
            db.session.rollback()
            flash(f"Error al actualizar empleado: {e}", "danger")
            return redirect(url_for("editar_empleado", id=id))

    return render_template("empleado_editar.html", empleado=empleado)

@app.route("/empleados/cambiar_estado/<int:id>", methods=["POST"])
@login_required
def cambiar_estado_empleado(id):
    try:
        empleado = Empleado.query.get_or_404(id)
        
        if empleado.estado == "Activo":
            empleado.estado = "Inactivo"
            mensaje = f"El empleado {empleado.nombre} ha sido desactivado."
        else:
            empleado.estado = "Activo"
            mensaje = f"El empleado {empleado.nombre} ha sido activado."
            
        db.session.commit()
        flash(mensaje, "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Error al cambiar el estado del empleado: {e}", "danger")
    
    return redirect(url_for("listar_empleados"))

@app.route("/turnos/eliminar/<int:id>", methods=["POST"])
@login_required
def eliminar_turno(id):
    try:
        turno = Turno.query.get_or_404(id)
        empleado_id = turno.id_empleado
        
        if turno.hora_entrada and turno.hora_salida:
            empleado = Empleado.query.get(empleado_id)
            res = calcular_turno(turno.fecha, turno.hora_entrada, turno.hora_salida)
            balance_a_revertir = res["total_horas"] - 8.0
            empleado.balance_horas -= balance_a_revertir
        
        db.session.delete(turno)
        db.session.commit()
        
        # Obtener el balance actualizado del empleado
        empleado_actualizado = Empleado.query.get(empleado_id)
        nuevo_balance = empleado_actualizado.balance_horas
        
        # Devolver JSON en lugar de redireccionar para AJAX
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return jsonify({
                'success': True, 
                'message': 'Turno eliminado correctamente.',
                'nuevo_balance': nuevo_balance,
                'empleado_id': empleado_id
            })
        else:
            flash("El turno ha sido eliminado y el balance actualizado.", "success")
            return redirect(url_for("horario_diario"))
    except Exception as e:
        db.session.rollback()
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return jsonify({'success': False, 'message': f'Error al eliminar el turno: {e}'}), 500
        else:
            flash(f"Error al eliminar el turno: {e}", "danger")
            return redirect(url_for("horario_diario"))

@app.route("/turnos/editar/<int:id>", methods=["GET", "POST"])
@login_required
def editar_turno(id):
    turno = Turno.query.get_or_404(id)

    if request.method == "POST":
        try:
            empleado_id_anterior = turno.id_empleado
            empleado_id_nuevo = int(request.form["id_empleado"])
            
            # Revertir balance del empleado anterior si tenía horas registradas
            if turno.hora_entrada and turno.hora_salida:
                empleado_anterior = Empleado.query.get(empleado_id_anterior)
                res_anterior = calcular_turno(turno.fecha, turno.hora_entrada, turno.hora_salida)
                balance_a_revertir = res_anterior["total_horas"] - 8.0
                empleado_anterior.balance_horas -= balance_a_revertir

            # Actualizar datos del turno
            turno.id_empleado = empleado_id_nuevo
            turno.fecha = datetime.strptime(request.form["fecha"], "%Y-%m-%d").date()
            turno.hora_entrada = datetime.strptime(request.form["hora_entrada"], "%H:%M").time()
            turno.hora_salida = datetime.strptime(request.form["hora_salida"], "%H:%M").time()
            
            # Recalcular horas para el nuevo empleado
            res = calcular_turno(turno.fecha, turno.hora_entrada, turno.hora_salida)
            horas_nocturnas = res["nocturnas_base"] + res["nocturnas_dia_siguiente"]
            turno.horas_nocturnas = round(horas_nocturnas, 2)
            turno.horas_normales = round(res["total_horas"] - horas_nocturnas, 2)
            
            # Aplicar nuevo balance al empleado
            empleado_nuevo = Empleado.query.get(empleado_id_nuevo)
            balance_a_aplicar = res["total_horas"] - 8.0
            empleado_nuevo.balance_horas += balance_a_aplicar
            
            db.session.commit()
            
            # Obtener balances actualizados de ambos empleados
            balance_empleado_anterior = empleado_anterior.balance_horas if empleado_id_anterior != empleado_id_nuevo else None
            balance_empleado_nuevo = empleado_nuevo.balance_horas
            
            # Devolver JSON en lugar de redireccionar para AJAX
            if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                response_data = {
                    'success': True, 
                    'message': 'Turno actualizado con éxito.',
                    'empleado_id_nuevo': empleado_id_nuevo,
                    'balance_empleado_nuevo': balance_empleado_nuevo
                }
                
                if balance_empleado_anterior is not None:
                    response_data['empleado_id_anterior'] = empleado_id_anterior
                    response_data['balance_empleado_anterior'] = balance_empleado_anterior
                
                return jsonify(response_data)
            else:
                flash("Turno actualizado con éxito.", "success")
                return redirect(url_for("nuevo_turno", fecha_inicio=turno.fecha.strftime('%Y-%m-%d'), fecha_fin=turno.fecha.strftime('%Y-%m-%d')))
        except Exception as e:
            db.session.rollback()
            if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                return jsonify({'success': False, 'message': f'Error al actualizar el turno: {e}'}), 500
            else:
                flash(f"Error al actualizar el turno: {e}", "danger")
                return redirect(url_for("editar_turno", id=id))

    empleados_activos = Empleado.query.filter_by(estado='Activo').order_by(Empleado.nombre).all()
    return render_template("turno_editar.html", turno=turno, empleados=empleados_activos)

@app.route("/empleados/eliminar/<int:id>", methods=["POST"])
@login_required
def eliminar_empleado(id):
    try:
        empleado_a_eliminar = Empleado.query.get_or_404(id)
        nombre_empleado = empleado_a_eliminar.nombre

        Turno.query.filter_by(id_empleado=id).delete()
        ResumenMensual.query.filter_by(id_empleado=id).delete()
        
        db.session.delete(empleado_a_eliminar)
        
        db.session.commit()
        
        flash(f"El empleado '{nombre_empleado}' y todos sus registros han sido eliminados permanentemente.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Error al eliminar al empleado: {e}", "danger")
    
    return redirect(url_for("listar_empleados"))

@app.route("/api/empleado/<int:id>/balance")
@login_required
def get_balance_empleado(id):
    empleado = Empleado.query.get_or_404(id)
    return {"balance": empleado.balance_horas}

@app.route("/horario/diario", methods=["GET", "POST"])
@login_required
def horario_diario():
    if request.method == "POST":
        fecha_str_post = request.form.get('fecha_actual')
        fecha_obj_post = datetime.strptime(fecha_str_post, "%Y-%m-%d").date()
        
        with db.session.no_autoflush:
            empleados_activos = Empleado.query.filter_by(estado='Activo').all()
            for empleado in empleados_activos:
                entrada = request.form.get(f"entrada_{empleado.id_empleado}")
                salida = request.form.get(f"salida_{empleado.id_empleado}")
                obs_select = request.form.get(f"obs_select_{empleado.id_empleado}")
                obs_text = request.form.get(f"obs_text_{empleado.id_empleado}")
                usar_balance = request.form.get(f"usar_balance_{empleado.id_empleado}")

                observacion_final = None
                if obs_select == "Otro" and obs_text:
                    observacion_final = obs_text
                elif obs_select and obs_select != "Otro":
                    observacion_final = obs_select

                turno_existente = Turno.query.filter_by(id_empleado=empleado.id_empleado, fecha=fecha_obj_post).first()

                if observacion_final:
                    if turno_existente:
                        turno_existente.hora_entrada = None
                        turno_existente.hora_salida = None
                        turno_existente.observaciones = observacion_final
                    else:
                        nuevo_turno = Turno(
                            id_empleado=empleado.id_empleado,
                            fecha=fecha_obj_post,
                            observaciones=observacion_final
                        )
                        db.session.add(nuevo_turno)
                else:
                    if entrada and salida:
                        hora_entrada = datetime.strptime(entrada, "%H:%M").time()
                        hora_salida = datetime.strptime(salida, "%H:%M").time()
                        res_nuevo = calcular_turno(fecha_obj_post, hora_entrada, hora_salida)
                        total_horas_nuevo = res_nuevo["total_horas"]
                        
                        # Calcular diferencia con 8 horas
                        diferencia = total_horas_nuevo - 8.0
                        
                        # Si se marca usar balance, compensar las horas
                        if usar_balance:
                            if diferencia < 0:  # Faltan horas
                                # Usar balance para cubrir el déficit
                                horas_a_usar = min(abs(diferencia), empleado.balance_horas)
                                empleado.balance_horas -= horas_a_usar
                                # No se aplica balance negativo porque ya se compensó
                                balance_a_sumar = 0.0
                            else:  # Horas extra
                                # Agregar las horas extra al balance
                                balance_a_sumar = diferencia
                        else:
                            # Aplicar el balance normal
                            balance_a_sumar = diferencia
                        
                        empleado.balance_horas += balance_a_sumar
                        
                        if turno_existente:
                            turno_existente.hora_entrada = hora_entrada
                            turno_existente.hora_salida = hora_salida
                            turno_existente.observaciones = None
                        else:
                            db.session.add(Turno(
                                id_empleado=empleado.id_empleado,
                                fecha=fecha_obj_post,
                                hora_entrada=hora_entrada,
                                hora_salida=hora_salida
                            ))
                    elif turno_existente:
                        db.session.delete(turno_existente)
        
        db.session.commit()
        flash("Horario guardado y balances actualizados.", "success")
        return redirect(url_for("horario_diario", fecha=fecha_str_post))

    fecha_inicio_str = request.args.get('fecha_inicio')
    fecha_fin_str = request.args.get('fecha_fin')
    empleado_id_str = request.args.get('id_empleado_filtro')

    turnos_detallados = []
    resumen_periodo = {"normales": 0.0, "nocturnas": 0.0, "nocturnas_dominicales": 0.0, "total": 0.0}
    resumen_semanal_empleados = {}
    nombres_empleados = {}

    if fecha_inicio_str and fecha_fin_str:
        try:
            fecha_inicio = datetime.strptime(fecha_inicio_str, "%Y-%m-%d").date()
            fecha_fin = datetime.strptime(fecha_fin_str, "%Y-%m-%d").date()
            query = Turno.query.filter(Turno.fecha.between(fecha_inicio, fecha_fin))
            if empleado_id_str and empleado_id_str.isdigit():
                query = query.filter_by(id_empleado=int(empleado_id_str))
            turnos_en_periodo = query.order_by(Turno.fecha).all()
            
            # Calcular horas por empleado para el resumen semanal
            horas_por_empleado = {}
            todos_los_empleados = Empleado.query.all()
            for emp in todos_los_empleados:
                horas_por_empleado[emp.id_empleado] = 0
                nombres_empleados[emp.id_empleado] = emp.nombre
            
            for turno in turnos_en_periodo:
                if not turno.hora_entrada: 
                    continue
                res = calcular_turno(turno.fecha, turno.hora_entrada, turno.hora_salida)
                total_turno = res["total_horas"]
                nocturnas_base = res["nocturnas_base"]
                nocturnas_dia_siguiente = res["nocturnas_dia_siguiente"]

                # Determinar si el día del turno y el día siguiente son domingo
                es_domingo_turno = turno.fecha.weekday() == 6  # 6 es domingo
                es_domingo_siguiente = (turno.fecha + timedelta(days=1)).weekday() == 6

                nocturnas_dominicales = (nocturnas_base if es_domingo_turno else 0) + (nocturnas_dia_siguiente if es_domingo_siguiente else 0)
                nocturnas_normales = (nocturnas_base if not es_domingo_turno else 0) + (nocturnas_dia_siguiente if not es_domingo_siguiente else 0)

                horas_normales = total_turno - nocturnas_normales - nocturnas_dominicales

                resumen_periodo["nocturnas"] += nocturnas_normales
                resumen_periodo["nocturnas_dominicales"] += nocturnas_dominicales
                resumen_periodo["normales"] += horas_normales
                resumen_periodo["total"] += total_turno
                
                # Acumular horas por empleado
                horas_por_empleado[turno.id_empleado] += total_turno
                
                turnos_detallados.append({
                    "id": turno.id_turno, 
                    "id_empleado": turno.id_empleado,
                    "empleado": turno.empleado.nombre, 
                    "fecha": turno.fecha, 
                    "entrada": turno.hora_entrada, 
                    "salida": turno.hora_salida, 
                    "normales": horas_normales,
                    "nocturnas": nocturnas_normales,
                    "nocturnas_dominicales": nocturnas_dominicales,
                    "total": total_turno,
                    "observaciones": turno.observaciones
                })
                
            # Calcular balance semanal por empleado (horas trabajadas - horas esperadas)
            dias_rango = (fecha_fin - fecha_inicio).days + 1
            horas_esperadas = (44 / 7) * dias_rango  # 44 horas semanales prorrateadas
            
            for emp_id, horas_trabajadas in horas_por_empleado.items():
                resumen_semanal_empleados[emp_id] = horas_trabajadas - horas_esperadas
                
        except ValueError:
            flash("Formato de fecha inválido.", "warning")

    fecha_planificacion_str = request.args.get('fecha', default=date.today().strftime('%Y-%m-%d'))
    fecha_planificacion_obj = datetime.strptime(fecha_planificacion_str, "%Y-%m-%d").date()
    
    # Calcular semana actual
    semana_actual = fecha_planificacion_obj.isocalendar()[1]
    fecha_inicio_semana = fecha_planificacion_obj - timedelta(days=fecha_planificacion_obj.weekday())
    fecha_fin_semana = fecha_inicio_semana + timedelta(days=6)
    
    empleados_activos = Empleado.query.filter_by(estado='Activo').order_by(Empleado.cargo, Empleado.nombre).all()
    empleados_agrupados = {k: list(v) for k, v in groupby(empleados_activos, key=lambda emp: emp.cargo)}
    turnos_existentes_raw = Turno.query.filter_by(fecha=fecha_planificacion_obj).all()
    turnos_existentes = {t.id_empleado: t for t in turnos_existentes_raw}
    todos_los_empleados = Empleado.query.order_by(Empleado.nombre).all()

    # Calcular horas trabajadas en la semana actual para cada empleado
    for empleado in empleados_activos:
        # Obtener turnos de la semana actual
        turnos_semana = Turno.query.filter(
            Turno.id_empleado == empleado.id_empleado,
            Turno.fecha >= fecha_inicio_semana,
            Turno.fecha <= fecha_fin_semana,
            Turno.hora_entrada.isnot(None)
        ).all()
        
        # Calcular horas trabajadas en la semana
        horas_semana = 0
        for turno in turnos_semana:
            res = calcular_turno(turno.fecha, turno.hora_entrada, turno.hora_salida)
            horas_semana += res['total_horas']
        
        empleado.horas_semana_actual = horas_semana
        empleado.balance_semanal = horas_semana - 44  # Base de 44 horas semanales
        empleado.balance_acumulado = empleado.balance_horas  # Agregar esta línea

    return render_template("horario_diario.html", 
                           fecha=fecha_planificacion_obj, 
                           empleados_agrupados=empleados_agrupados,
                           turnos_existentes=turnos_existentes, 
                           turnos_detallados=turnos_detallados,
                           resumen_periodo=resumen_periodo, 
                           todos_los_empleados=todos_los_empleados,
                           filtros=request.args,
                           semana_actual=semana_actual,
                           fecha_inicio_semana=fecha_inicio_semana,
                           fecha_fin_semana=fecha_fin_semana,
                           resumen_semanal_empleados=resumen_semanal_empleados,
                           nombres_empleados=nombres_empleados)

@app.route("/api/consultar_turnos")
@login_required
def api_consultar_turnos():
    try:
        fecha_inicio_str = request.args.get('fecha_inicio')
        fecha_fin_str = request.args.get('fecha_fin')
        empleado_id_str = request.args.get('id_empleado_filtro')

        if not (fecha_inicio_str and fecha_fin_str):
            return {"error": "Fechas de inicio y fin son requeridas"}, 400

        fecha_inicio = datetime.strptime(fecha_inicio_str, "%Y-%m-%d").date()
        fecha_fin = datetime.strptime(fecha_fin_str, "%Y-%m-%d").date()
        
        query = Turno.query.filter(Turno.fecha.between(fecha_inicio, fecha_fin))
        if empleado_id_str and empleado_id_str.isdigit():
            query = query.filter_by(id_empleado=int(empleado_id_str))

        turnos_en_periodo = query.order_by(Turno.fecha).all()
        
        turnos_detallados = []
        resumen_periodo = {"normales": 0.0, "nocturnas": 0.0, "nocturnas_dominicales": 0.0, "total": 0.0}
        resumen_semanal_empleados = {}
        nombres_empleados = {}

        # Obtener todos los empleados para el resumen semanal
        todos_empleados = Empleado.query.all()
        for emp in todos_empleados:
            resumen_semanal_empleados[emp.id_empleado] = 0
            nombres_empleados[emp.id_empleado] = emp.nombre

        for turno in turnos_en_periodo:
            if not turno.hora_entrada: 
                turnos_detallados.append({
                    "id": turno.id_turno, 
                    "empleado": turno.empleado.nombre, 
                    "fecha": turno.fecha.strftime('%d/%m/%Y'), 
                    "entrada": "",
                    "salida": "",
                    "nocturnas": "0.00",
                    "nocturnas_dominicales": "0.00",
                    "total": "0.00 h",
                    "observaciones": turno.observaciones or ""
                })
                continue

            res = calcular_turno(turno.fecha, turno.hora_entrada, turno.hora_salida)
            total_turno = res["total_horas"]
            nocturnas_base = res["nocturnas_base"]
            nocturnas_dia_siguiente = res["nocturnas_dia_siguiente"]

            es_domingo_turno = turno.fecha.weekday() == 6
            es_domingo_siguiente = (turno.fecha + timedelta(days=1)).weekday() == 6

            nocturnas_dominicales = (nocturnas_base if es_domingo_turno else 0) + (nocturnas_dia_siguiente if es_domingo_siguiente else 0)
            nocturnas_normales = (nocturnas_base if not es_domingo_turno else 0) + (nocturnas_dia_siguiente if not es_domingo_siguiente else 0)

            horas_normales = total_turno - nocturnas_normales - nocturnas_dominicales

            resumen_periodo["nocturnas"] += nocturnas_normales
            resumen_periodo["nocturnas_dominicales"] += nocturnas_dominicales
            resumen_periodo["normales"] += horas_normales
            resumen_periodo["total"] += total_turno
            
            # Acumular horas para el resumen semanal por empleado
            resumen_semanal_empleados[turno.id_empleado] += total_turno

            turnos_detallados.append({
                "id": turno.id_turno, 
                "empleado": turno.empleado.nombre, 
                "fecha": turno.fecha.strftime('%d/%m/%Y'), 
                "entrada": turno.hora_entrada.strftime('%I:%M %p'), 
                "salida": turno.hora_salida.strftime('%I:%M %p'),
                "nocturnas": f"{nocturnas_normales:.2f}",
                "nocturnas_dominicales": f"{nocturnas_dominicales:.2f}",
                "total": f"{total_turno:.2f} h",
                "observaciones": turno.observaciones or ""
            })

        # Calcular balance semanal (horas trabajadas - horas esperadas)
        dias_rango = (fecha_fin - fecha_inicio).days + 1
        horas_esperadas = (44 / 7) * dias_rango  # 44 horas semanales prorrateadas
        
        for emp_id in resumen_semanal_empleados:
            resumen_semanal_empleados[emp_id] -= horas_esperadas

        return {
            "resumen": resumen_periodo, 
            "detalle": turnos_detallados,
            "resumen_semanal": resumen_semanal_empleados,
            "nombres_empleados": nombres_empleados
        }
    except Exception as e:
        return {"error": str(e)}, 500

@app.route("/ajustes")
@login_required
def vista_ajustes():
    empleados = Empleado.query.order_by(Empleado.nombre).all()
    return render_template("ajustes_balance.html", empleados=empleados)

@app.route("/api/ajustar_balance", methods=["POST"])
@login_required
def api_ajustar_balance():
    try:
        data = request.json
        empleado_id = data.get('empleado_id')
        cantidad = float(data.get('cantidad'))
        tipo = data.get('tipo')

        empleado = Empleado.query.get_or_404(empleado_id)

        if tipo == 'sumar':
            empleado.balance_horas += cantidad
        elif tipo == 'restar':
            empleado.balance_horas -= cantidad
        elif tipo == 'establecer':
            empleado.balance_horas = cantidad
        else:
            return jsonify({"error": "Tipo de ajuste no válido"}), 400

        db.session.commit()

        return jsonify({
            "success": True, 
            "nuevo_balance": empleado.balance_horas
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": str(e)}), 500

# --- RUTAS PARA BALANCE ACUMULADO Y SEMANAS ---

@app.route("/api/empleado/<int:id>/semanas_balance")
@login_required
def get_semanas_balance_empleado(id):
    try:
        empleado = Empleado.query.get_or_404(id)
        
        # Obtener todos los turnos del empleado
        turnos = Turno.query.filter_by(id_empleado=id).filter(Turno.hora_entrada.isnot(None)).order_by(Turno.fecha).all()
        
        # Agrupar turnos por semana
        semanas_dict = {}
        for turno in turnos:
            # Calcular número de semana ISO (lunes a domingo)
            semana_num = turno.fecha.isocalendar()[1]
            año = turno.fecha.isocalendar()[0]
            clave = f"{año}-{semana_num}"
            
            if clave not in semanas_dict:
                # Calcular fechas de inicio y fin de la semana (lunes a domingo)
                fecha_inicio = turno.fecha - timedelta(days=turno.fecha.weekday())
                fecha_fin = fecha_inicio + timedelta(days=6)
                
                semanas_dict[clave] = {
                    'numero_semana': semana_num,
                    'año': año,
                    'fecha_inicio': fecha_inicio.strftime('%d/%m/%Y'),
                    'fecha_fin': fecha_fin.strftime('%d/%m/%Y'),
                    'horas_trabajadas': 0.0,
                    'horas_esperadas': 44.0,  # 44 horas semanales
                    'turnos': []
                }
            
            # Calcular horas del turno
            res = calcular_turno(turno.fecha, turno.hora_entrada, turno.hora_salida)
            horas_turno = res['total_horas']
            semanas_dict[clave]['horas_trabajadas'] += horas_turno
            semanas_dict[clave]['turnos'].append({
                'fecha': turno.fecha.strftime('%d/%m/%Y'),
                'horas': horas_turno
            })
        
        # Convertir a lista y ordenar por año y semana
        semanas_list = sorted(semanas_dict.values(), key=lambda x: (x['año'], x['numero_semana']))
        
        # Calcular balances acumulados
        acumulado = 0.0
        semanas_positivas = 0
        semanas_negativas = 0
        
        for semana in semanas_list:
            semana['balance_semanal'] = semana['horas_trabajadas'] - semana['horas_esperadas']
            acumulado += semana['balance_semanal']
            semana['acumulado'] = acumulado
            
            if semana['balance_semanal'] >= 0:
                semanas_positivas += 1
            else:
                semanas_negativas += 1
        
        return jsonify({
            'nombre_empleado': empleado.nombre,
            'semanas': semanas_list,
            'balance_total': acumulado,
            'semanas_positivas': semanas_positivas,
            'semanas_negativas': semanas_negativas,
            'total_semanas': len(semanas_list)
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route("/api/empleado/<int:id>/balance_detalle")
@login_required
def get_balance_detalle_empleado(id):
    try:
        empleado = Empleado.query.get_or_404(id)
        
        # Obtener los últimos 3 meses de turnos para el detalle
        tres_meses_atras = date.today() - timedelta(days=90)
        turnos_recientes = Turno.query.filter(
            Turno.id_empleado == id,
            Turno.fecha >= tres_meses_atras,
            Turno.hora_entrada.isnot(None)
        ).order_by(Turno.fecha.desc()).limit(50).all()
        
        detalle_turnos = []
        for turno in turnos_recientes:
            res = calcular_turno(turno.fecha, turno.hora_entrada, turno.hora_salida)
            horas_totales = res['total_horas']
            balance_turno = horas_totales - 8.0
            
            detalle_turnos.append({
                'fecha': turno.fecha.strftime('%d/%m/%Y'),
                'entrada': turno.hora_entrada.strftime('%H:%M'),
                'salida': turno.hora_salida.strftime('%H:%M'),
                'horas_totales': horas_totales,
                'balance_turno': balance_turno,
                'horas_nocturnas': res['nocturnas_base'] + res['nocturnas_dia_siguiente']
            })
        
        return jsonify({
            'nombre_empleado': empleado.nombre,
            'balance_actual': empleado.balance_horas,
            'detalle_turnos': detalle_turnos,
            'total_turnos': len(detalle_turnos),
            'periodo': 'Últimos 3 meses'
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
@app.route("/graficos")
@login_required
def vista_graficos():
    empleados = Empleado.query.order_by(Empleado.nombre).all()
    
    # Obtener año actual y rango de años
    current_year = datetime.now().year
    years = range(current_year - 2, current_year + 3)  # 5 años en total
    
    # Pasar la fecha actual al template
    now = datetime.now()
    
    # Obtener cargos únicos para filtros
    cargos = db.session.query(Empleado.cargo).distinct().filter(Empleado.cargo.isnot(None)).all()
    cargos = [cargo[0] for cargo in cargos if cargo[0]]
    
    return render_template("graficos.html", 
                         empleados=empleados,
                         years=years,
                         current_year=current_year,
                         now=now,
                         cargos=cargos)

@app.route('/api/graficos/datos')
def api_datos_graficos():
    try:
        # Obtener parámetros del request
        tipo_periodo = request.args.get('tipo_periodo', 'mensual')
        mes = int(request.args.get('mes', datetime.now().month))
        anio = int(request.args.get('anio', datetime.now().year))
        id_empleado = request.args.get('id_empleado') or None
        cargo = request.args.get('cargo') or None
        
        # Calcular fechas según el tipo de período
        fecha_inicio, fecha_fin = calcular_rango_fechas(tipo_periodo, mes, anio, 
                                                      request.args.get('quincena'),
                                                      request.args.get('semana'),
                                                      request.args.get('fecha_inicio'),
                                                      request.args.get('fecha_fin'))
        
        # Asegurar que las fechas sean del tipo correcto (date)
        if isinstance(fecha_inicio, datetime):
            fecha_inicio = fecha_inicio.date()
        if isinstance(fecha_fin, datetime):
            fecha_fin = fecha_fin.date()
        
        app.logger.debug(f"Fechas calculadas: {fecha_inicio} a {fecha_fin}")
        
        # Obtener datos del reporte
        datos_reporte = preparar_datos_reporte(fecha_inicio, fecha_fin)
        
        # Aplicar filtros adicionales si se especificaron
        if id_empleado or cargo:
            datos_reporte_filtrado = {}
            for emp_id, emp_data in datos_reporte.items():
                # Filtrar por empleado
                if id_empleado and str(emp_data.get('id_empleado')) != str(id_empleado):
                    continue
                # Filtrar por cargo
                if cargo and emp_data.get('cargo') != cargo:
                    continue
                datos_reporte_filtrado[emp_id] = emp_data
            datos_reporte = datos_reporte_filtrado
        
        # Generar datos para gráficos
        datos_graficos = generar_datos_graficos(datos_reporte, fecha_inicio, fecha_fin)
        
        # Log para debug
        app.logger.debug(f"Datos gráficos generados: {list(datos_graficos.keys())}")
        if 'estadisticas' in datos_graficos:
            app.logger.debug(f"Estadísticas: {datos_graficos['estadisticas']}")
        
        return jsonify(datos_graficos)
        
    except Exception as e:
        app.logger.error(f"Error en api_datos_graficos: {str(e)}", exc_info=True)
        return jsonify({'error': str(e)}), 500

def calcular_rango_fechas(tipo_periodo, mes, anio, quincena=None, semana=None, fecha_inicio_str=None, fecha_fin_str=None):
    """Calcula el rango de fechas según el tipo de período - DEVUELVE date objects"""
    try:
        # Para período personalizado
        if tipo_periodo == 'personalizado' and fecha_inicio_str and fecha_fin_str:
            fecha_inicio = datetime.strptime(fecha_inicio_str, '%Y-%m-%d').date()
            fecha_fin = datetime.strptime(fecha_fin_str, '%Y-%m-%d').date()
            return fecha_inicio, fecha_fin
        
        # Para otros tipos de período
        if tipo_periodo == 'mensual':
            fecha_inicio = date(anio, mes, 1)
            if mes == 12:
                fecha_fin = date(anio + 1, 1, 1) - timedelta(days=1)
            else:
                fecha_fin = date(anio, mes + 1, 1) - timedelta(days=1)
        
        elif tipo_periodo == 'quincenal':
            if quincena == '1':
                fecha_inicio = date(anio, mes, 1)
                fecha_fin = date(anio, mes, 15)
            else:
                fecha_inicio = date(anio, mes, 16)
                if mes == 12:
                    fecha_fin = date(anio + 1, 1, 1) - timedelta(days=1)
                else:
                    fecha_fin = date(anio, mes + 1, 1) - timedelta(days=1)
        
        elif tipo_periodo == 'semanal':
            # Para semanal, usamos el mes completo por simplicidad
            fecha_inicio = date(anio, mes, 1)
            if mes == 12:
                fecha_fin = date(anio + 1, 1, 1) - timedelta(days=1)
            else:
                fecha_fin = date(anio, mes + 1, 1) - timedelta(days=1)
        
        else:
            # Default mensual
            fecha_inicio = date(anio, mes, 1)
            if mes == 12:
                fecha_fin = date(anio + 1, 1, 1) - timedelta(days=1)
            else:
                fecha_fin = date(anio, mes + 1, 1) - timedelta(days=1)
        
        app.logger.debug(f"Rango calculado: {fecha_inicio} - {fecha_fin}")
        return fecha_inicio, fecha_fin
        
    except Exception as e:
        app.logger.error(f"Error en calcular_rango_fechas: {str(e)}", exc_info=True)
        # Fallback: mes actual como date
        hoy = date.today()
        primer_dia_mes = date(hoy.year, hoy.month, 1)
        return primer_dia_mes, hoy

def generar_datos_graficos(datos, fecha_inicio, fecha_fin):
    """Genera la estructura de datos para todos los gráficos - Acepta date objects"""
    
    # Asegurar que fecha_inicio y fecha_fin sean date objects
    if isinstance(fecha_inicio, datetime):
        fecha_inicio = fecha_inicio.date()
    if isinstance(fecha_fin, datetime):
        fecha_fin = fecha_fin.date()
    
    # Si no hay datos, devolver estructura vacía pero completa
    if not datos:
        return crear_estructura_vacia()
    
    # Datos básicos
    labels_empleados = []
    datos_totales = []
    datos_diurnas = []
    datos_nocturnas = []
    datos_nocturnas_dom = []

    # Recorremos empleados
    for emp_id, emp_data in datos.items():
        labels_empleados.append(emp_data.get('nombre', 'Sin Nombre'))

        # Calcular totales desde los días
        total_horas_emp = sum(v.get('total', 0.0) for v in emp_data.get('dias', {}).values())
        hn_emp = sum(v.get('hn', 0.0) for v in emp_data.get('dias', {}).values())
        hn_dom_emp = sum(v.get('hn_dom', 0.0) for v in emp_data.get('dias', {}).values())

        datos_totales.append(total_horas_emp)
        datos_diurnas.append(total_horas_emp - hn_emp - hn_dom_emp)  # Horas diurnas como diferencia
        datos_nocturnas.append(hn_emp)
        datos_nocturnas_dom.append(hn_dom_emp)

    # Estadísticas generales con valores por defecto
    total_horas_periodo = sum(datos_totales) if datos_totales else 0.0
    total_diurnas = sum(datos_diurnas) if datos_diurnas else 0.0
    total_nocturnas = sum(datos_nocturnas) if datos_nocturnas else 0.0
    total_nocturnas_dom = sum(datos_nocturnas_dom) if datos_nocturnas_dom else 0.0

    dias_periodo = max((fecha_fin - fecha_inicio).days + 1, 1)  # Evitar división por cero
    promedio_por_empleado = (total_horas_periodo / len(datos)) if datos else 0.0

    # Gráfico principal
    grafico_principal = {
        'labels': labels_empleados,
        'diurnas': datos_diurnas,
        'nocturnas': datos_nocturnas,
        'nocturnasDom': datos_nocturnas_dom
    }

    # Top 5 nocturnas
    empleados_nocturnas = list(zip(labels_empleados, datos_nocturnas))
    empleados_nocturnas.sort(key=lambda x: x[1], reverse=True)
    top_5_nocturnas = empleados_nocturnas[:5]
    top_nocturnas = {
        'labels': [emp[0] for emp in top_5_nocturnas] if top_5_nocturnas else [],
        'data': [emp[1] for emp in top_5_nocturnas] if top_5_nocturnas else []
    }

    # Datos por cargo
    cargos_data = {}
    for emp_data in datos.values():
        cargo = emp_data.get('cargo', 'Sin Cargo') or 'Sin Cargo'
        total_cargo = sum(v.get('total', 0.0) for v in emp_data.get('dias', {}).values())
        cargos_data[cargo] = cargos_data.get(cargo, 0.0) + total_cargo

    por_cargo = {
        'labels': list(cargos_data.keys()),
        'data': list(cargos_data.values())
    }

    # Evolución semanal (datos de ejemplo)
    evolucion = {
        'labels': [f"Semana {i+1}" for i in range(4)],
        'data': [total_horas_periodo * factor for factor in [0.2, 0.3, 0.35, 0.15]]
    }

    # Ratio diurno/nocturno
    ratio = {
        'diurnas': total_diurnas,
        'nocturnas': total_nocturnas + total_nocturnas_dom
    }

    # Promedio por día
    promedio_dia = {
        'labels': labels_empleados[:10],
        'data': [(total / dias_periodo) for total in datos_totales[:10]]
    }

    # Tendencia mensual (datos de ejemplo)
    tendencia = {
        'labels': ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun'],
        'data': [total_horas_periodo * factor for factor in [0.8, 0.9, 1.0, 1.1, 1.2, 1.15]]
    }

    # Tabla de detalles
    detalles = []
    for i, emp_data in enumerate(datos.values()):
        total_horas_emp = datos_totales[i] if i < len(datos_totales) else 0.0
        nocturnas_emp = datos_nocturnas[i] if i < len(datos_nocturnas) else 0.0
        nocturnas_dom_emp = datos_nocturnas_dom[i] if i < len(datos_nocturnas_dom) else 0.0
        porcentaje_nocturno = ((nocturnas_emp + nocturnas_dom_emp) / total_horas_emp * 100.0) if total_horas_emp > 0 else 0.0

        detalles.append({
            'empleado': labels_empleados[i],
            'cargo': emp_data.get('cargo', 'Sin Cargo'),
            'totalHoras': round(total_horas_emp, 1),
            'horasDiurnas': round(datos_diurnas[i], 1) if i < len(datos_diurnas) else 0.0,
            'horasNocturnas': round(nocturnas_emp, 1),
            'horasNocturnasDom': round(nocturnas_dom_emp, 1),
            'promedioDia': round((total_horas_emp / dias_periodo), 1),
            'porcentajeNocturno': round(porcentaje_nocturno, 1)
        })

    return {
        'estadisticas': {
            'totalHoras': round(total_horas_periodo, 1),
            'horasDiurnas': round(total_diurnas, 1),
            'horasNocturnas': round(total_nocturnas, 1),
            'horasNocturnasDom': round(total_nocturnas_dom, 1),
            'totalEmpleados': len(datos),
            'promedioPorEmpleado': round(promedio_por_empleado, 1)
        },
        'graficoPrincipal': grafico_principal,
        'topNocturnas': top_nocturnas,
        'porCargo': por_cargo,
        'evolucion': evolucion,
        'ratio': ratio,
        'promedio': promedio_dia,
        'tendencia': tendencia,
        'detalles': detalles
    }

def crear_estructura_vacia():
    """Crea una estructura vacía para cuando no hay datos"""
    return {
        'estadisticas': {
            'totalHoras': 0.0,
            'horasDiurnas': 0.0,
            'horasNocturnas': 0.0,
            'horasNocturnasDom': 0.0,
            'totalEmpleados': 0,
            'promedioPorEmpleado': 0.0
        },
        'graficoPrincipal': {'labels': [], 'diurnas': [], 'nocturnas': [], 'nocturnasDom': []},
        'topNocturnas': {'labels': [], 'data': []},
        'porCargo': {'labels': [], 'data': []},
        'evolucion': {'labels': [], 'data': []},
        'ratio': {'diurnas': 0.0, 'nocturnas': 0.0},
        'promedio': {'labels': [], 'data': []},
        'tendencia': {'labels': [], 'data': []},
        'detalles': []
    }

def debug_fechas(fecha_inicio, fecha_fin):
    """Función temporal para debug de fechas"""
    app.logger.debug(f"Tipo fecha_inicio: {type(fecha_inicio)}, valor: {fecha_inicio}")
    app.logger.debug(f"Tipo fecha_fin: {type(fecha_fin)}, valor: {fecha_fin}")
    
    # Usar temporalmente en api_datos_graficos:
    debug_fechas(fecha_inicio, fecha_fin)

    
if __name__ == "__main__":
    app.run(debug=True)