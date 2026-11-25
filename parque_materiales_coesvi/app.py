from datetime import datetime
import os
import sqlite3
from exportador_excel import generar_excel_entrada, generar_excel_salida
from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)

app = Flask(__name__)
app.secret_key = 'coesvi_parque_materiales_secret_key'
DB_NAME = 'database.db'


def get_db_connection():
  """Establece la conexión con la base de datos SQLite."""
  conn = sqlite3.connect(DB_NAME)
  conn.row_factory = sqlite3.Row
  return conn


def init_db():
  """Inicializa la base de datos ejecutando el archivo schema.sql si no existe."""
  if not os.path.exists(DB_NAME):
    conn = get_db_connection()
    with open('schema.sql', 'r', encoding='utf-8') as f:
      conn.executescript(f.read())
    conn.commit()
    conn.close()
    print('Base de datos inicializada correctamente.')


# ==========================================
# RUTAS DE LA APLICACIÓN
# ==========================================


@app.route('/')
def index():
  """Dashboard principal con resumen general de movimientos e inventario."""
  conn = get_db_connection()

  total_entradas = conn.execute('SELECT COUNT(*) FROM entradas').fetchone()[0]
  total_salidas = conn.execute('SELECT COUNT(*) FROM salidas').fetchone()[0]
  total_materiales = conn.execute(
      'SELECT COUNT(*) FROM materiales'
  ).fetchone()[0]

  ultimas_entradas = conn.execute(
      'SELECT * FROM entradas ORDER BY id DESC LIMIT 5'
  ).fetchall()
  ultimas_salidas = conn.execute(
      'SELECT * FROM salidas ORDER BY id DESC LIMIT 5'
  ).fetchall()

  conn.close()
  return render_template(
      'index.html',
      total_entradas=total_entradas,
      total_salidas=total_salidas,
      total_materiales=total_materiales,
      ultimas_entradas=ultimas_entradas,
      ultimas_salidas=ultimas_salidas,
  )


@app.route('/inventario')
def inventario():
  """Muestra el catálogo de materiales y sus existencias actuales en patio."""
  conn = get_db_connection()
  materiales = conn.execute(
      'SELECT * FROM materiales ORDER BY nombre ASC'
  ).fetchall()
  conn.close()
  return render_template('inventario.html', materiales=materiales)


@app.route('/materiales/agregar', methods=['POST'])
def agregar_material():
  """Permite dar de alta un nuevo material base en el catálogo."""
  codigo = request.form['codigo']
  nombre = request.form['nombre']
  unidad = request.form['unidad']

  conn = get_db_connection()
  conn.execute(
      'INSERT INTO materiales (codigo, nombre, unidad, stock_actual) VALUES'
      ' (?, ?, ?, 0.0)',
      (codigo, nombre, unidad),
  )
  conn.commit()
  conn.close()
  flash('Material agregado al catálogo con éxito.', 'success')
  return redirect(url_for('inventario'))


@app.route('/entradas', methods=['GET', 'POST'])
def entradas():
  """Muestra el listado de notas de entrada y procesa el registro de nuevas entradas."""
  conn = get_db_connection()

  if request.method == 'POST':
    folio = request.form['folio']
    fecha = request.form['fecha'] or datetime.now().strftime('%Y-%m-%d')
    proveedor = request.form['proveedor']
    factura_no = request.form['factura_no']
    remision_no = request.form['remision_no']

    material_id = request.form['material_id']
    cantidad = float(request.form['cantidad'])
    importe = float(request.form['importe'] or 0.0)

    mat = conn.execute(
        'SELECT * FROM materiales WHERE id = ?', (material_id,)
    ).fetchone()
    if mat:
      descripcion = mat['nombre']
      unidad = mat['unidad']

      cursor = conn.cursor()
      cursor.execute(
          'INSERT INTO entradas (folio, fecha, proveedor, factura_no,'
          ' remision_no) VALUES (?, ?, ?, ?, ?)',
          (folio, fecha, proveedor, factura_no, remision_no),
      )
      entrada_id = cursor.lastrowid

      cursor.execute(
          'INSERT INTO detalle_entradas (entrada_id, cantidad, unidad,'
          ' descripcion, importe) VALUES (?, ?, ?, ?, ?)',
          (entrada_id, cantidad, unidad, descripcion, importe),
      )

      cursor.execute(
          'UPDATE materiales SET stock_actual = stock_actual + ? WHERE id = ?',
          (cantidad, material_id),
      )

      conn.commit()
      flash('Nota de entrada registrada y stock actualizado.', 'success')
    conn.close()
    return redirect(url_for('entradas'))

  lista_entradas = conn.execute(
      'SELECT e.*, d.cantidad, d.unidad, d.descripcion, d.importe FROM entradas'
      ' e JOIN detalle_entradas d ON e.id = d.entrada_id ORDER BY e.id DESC'
  ).fetchall()
  materiales = conn.execute(
      'SELECT * FROM materiales ORDER BY nombre ASC'
  ).fetchall()
  conn.close()

  return render_template(
      'entradas.html', entradas=lista_entradas, materiales=materiales
  )


@app.route('/entradas/<int:entrada_id>/excel')
def descargar_excel_entrada(entrada_id):
  conn = get_db_connection()
  registro = conn.execute(
      """
        SELECT e.*, d.cantidad, d.unidad, d.descripcion, d.importe
        FROM entradas e
        JOIN detalle_entradas d ON e.id = d.entrada_id
        WHERE e.id = ?
    """,
      (entrada_id,),
  ).fetchone()
  conn.close()

  if not registro:
    flash('No se encontró la nota de entrada solicitada.', 'danger')
    return redirect(url_for('entradas'))

  ruta_excel = generar_excel_entrada(registro)
  return send_file(ruta_excel, as_attachment=True)


@app.route('/salidas', methods=['GET', 'POST'])
def salidas():
  """Muestra el listado de notas de salida y procesa las entregas a beneficiarios."""
  conn = get_db_connection()

  if request.method == 'POST':
    folio = request.form['folio']
    fecha = request.form['fecha'] or datetime.now().strftime('%Y-%m-%d')
    beneficiario = request.form['beneficiario']
    recibe_en = request.form['recibe_en']

    material_id = request.form['material_id']
    cantidad = float(request.form['cantidad'])

    mat = conn.execute(
        'SELECT * FROM materiales WHERE id = ?', (material_id,)
    ).fetchone()
    if mat:
      descripcion = mat['nombre']
      unidad = mat['unidad']
      stock_actual = mat['stock_actual']

      if cantidad > stock_actual:
        flash(
            f'Error: Stock insuficiente. Existencia actual: {stock_actual}'
            f' {unidad}',
            'danger',
        )
      else:
        cursor = conn.cursor()
        cursor.execute(
            'INSERT INTO salidas (folio, fecha, beneficiario, recibe_en)'
            ' VALUES (?, ?, ?, ?)',
            (folio, fecha, beneficiario, recibe_en),
        )
        salida_id = cursor.lastrowid

        cursor.execute(
            'INSERT INTO detalle_salidas (salida_id, cantidad, unidad,'
            ' descripcion) VALUES (?, ?, ?, ?)',
            (salida_id, cantidad, unidad, descripcion),
        )

        cursor.execute(
            'UPDATE materiales SET stock_actual = stock_actual - ? WHERE id = ?',
            (cantidad, material_id),
        )

        conn.commit()
        flash(
            'Nota de salida registrada y inventario actualizado con éxito.',
            'success',
        )

    conn.close()
    return redirect(url_for('salidas'))

  lista_salidas = conn.execute(
      'SELECT s.*, d.cantidad, d.unidad, d.descripcion FROM salidas s JOIN'
      ' detalle_salidas d ON s.id = d.salida_id ORDER BY s.id DESC'
  ).fetchall()
  materiales = conn.execute(
      'SELECT * FROM materiales ORDER BY nombre ASC'
  ).fetchall()
  conn.close()

  return render_template(
      'salidas.html', salidas=lista_salidas, materiales=materiales
  )


@app.route('/salidas/<int:salida_id>/excel')
def descargar_excel_salida(salida_id):
  conn = get_db_connection()
  registro = conn.execute(
      """
        SELECT s.*, d.cantidad, d.unidad, d.descripcion
        FROM salidas s
        JOIN detalle_salidas d ON s.id = d.salida_id
        WHERE s.id = ?
    """,
      (salida_id,),
  ).fetchone()
  conn.close()

  if not registro:
    flash('No se encontró la nota de salida solicitada.', 'danger')
    return redirect(url_for('salidas'))

  ruta_excel = generar_excel_salida(registro)
  return send_file(ruta_excel, as_attachment=True)


if __name__ == '__main__':
  init_db()
  app.run(debug=True, port=5000)