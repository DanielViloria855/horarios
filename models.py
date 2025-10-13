from extensions import db
from datetime import datetime

# Modelo empleados
class Empleado(db.Model):
    __tablename__ = "empleados"

    id_empleado = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    documento = db.Column(db.String(50), unique=True, nullable=False)
    cargo = db.Column(db.String(100))
    fecha_ingreso = db.Column(db.Date)
    estado = db.Column(db.String(20))
    balance_horas = db.Column(db.Float, nullable=False, default=0.0) # <-- CAMPO AÑADIDO

    # Relación con turnos y resumen
    turnos = db.relationship("Turno", backref="empleado", lazy=True)
    resumenes = db.relationship("ResumenMensual", backref="empleado", lazy=True)


# Modelo turnos
class Turno(db.Model):
    __tablename__ = "turnos"

    id_turno = db.Column(db.Integer, primary_key=True)
    id_empleado = db.Column(db.Integer, db.ForeignKey("empleados.id_empleado"), nullable=False)
    fecha = db.Column(db.Date, nullable=False)
    
    # --- LÍNEAS MODIFICADAS ---
    hora_entrada = db.Column(db.Time, nullable=True) # Se permite que sea nulo
    hora_salida = db.Column(db.Time, nullable=True)  # Se permite que sea nulo
    # --- FIN DE MODIFICACIÓN ---

    horas_normales = db.Column(db.Float, default=0)
    horas_extras = db.Column(db.Float, default=0)
    horas_nocturnas = db.Column(db.Float, default=0)
    observaciones = db.Column(db.String(100), nullable=True)

# Modelo resumen mensual
class ResumenMensual(db.Model):
    __tablename__ = "resumen_mensual"

    id_resumen = db.Column(db.Integer, primary_key=True)
    id_empleado = db.Column(db.Integer, db.ForeignKey("empleados.id_empleado"), nullable=False)
    anio = db.Column(db.Integer, nullable=False)
    mes = db.Column(db.Integer, nullable=False)
    total_horas_normales = db.Column(db.Float, default=0)
    total_horas_extras = db.Column(db.Float, default=0)
    total_horas_nocturnas = db.Column(db.Float, default=0)


# Modelo usuario
class Usuario(db.Model):
    __tablename__ = "usuario"

    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)