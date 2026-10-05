from flask import Flask, render_template,request,session,redirect,send_file
import sqlite3

from openpyxl import Workbook, load_workbook

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch

app = Flask(__name__)
app.secret_key = "LR2026"
@app.route("/")
def inicio():

    if "usuario" not in session:

       return """
       <html>

       <head>
       <link rel="stylesheet" href="/static/estilo.css">
       </head>

       <body>

       <img src="/static/Logo LR.jpeg" width="220">

       <h1>Acceso restringido</h1>

       <p>Debes iniciar sesión para acceder al sistema.</p>

       <a href="/login">
       <button>🔑 Iniciar sesión</button>
       </a>

       </body>

       </html>
       """
    return render_template("index.html")
@app.route("/login", methods=["GET","POST"])
def login():

    if request.method == "POST":

        usuario = request.form["usuario"]
        password = request.form["password"]

        print("Usuario recibido:",repr(usuario))
        print("Password recibido:",repr(password))

        if usuario =="LR" and password =="LR2010":
            print("Usuario recibido:", usuario)
            print("Password recibido:", password)

            session["usuario"] = usuario

            return """
            <html>

            <head>
            <link rel="stylesheet" href="/static/estilo.css">
            </head>

            <body>

            <img src="/static/Logo LR.jpeg" class="logo">

            <h1>¡Bienvenido!</h1>

            <h2>Acceso correcto</h2>

            <p>Has iniciado sesión correctamente.</p>

            <a href="/">
            <button>🏠 Ir al menú principal</button>
            </a>

            </body>

            </html>
            """

        return """
        <h1>Usuario o contraseña incorrectos</h1>

        <a href='/login'>
            <button>Intentar de nuevo</button>
        </a>
        """

    return """
        <html>

        <head>
        <link rel="stylesheet" href="/static/estilo.css">
        </head>

        <body>

        <h1>Iniciar sesión</h1>

        <form method='POST'>

        Usuario:
        <input type='text' name='usuario'>

        <br><br>

        Contraseña:
        <input type='password' name='password'>

        <br><br>

        <button type='submit'>
        Ingresar
        </button>

        </form>

        </body>
        </html>
    """
@app.route("/logout")
def logout():

    session.pop("usuario", None)

    return """
    <html>

    <head>
    <link rel="stylesheet" href="/static/estilo.css">
    </head>

    <body>

    <img src="/static/Logo LR.jpeg" class="logo">

    <h1>Sesión cerrada</h1>

    <h2>Hasta pronto</h2>

    <p>La sesión se cerró correctamente.</p>

    <a href="/login">
        <button>🔑 Iniciar sesión nuevamente</button>
    </a>

    </body>

    </html>
    """
def crear_base_datos():

    conexion = sqlite3.connect("inventario.db")
    cursor = conexion.cursor()

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS productos (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        nombre TEXT NOT NULL,

        marca TEXT,

        cantidad INTEGER,

        costo REAL,

        precio REAL

    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS ventas (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        nombre TEXT NOT NULL,

        marca TEXT,

        cantidad INTEGER,

        precio REAL,

        fecha TEXT

    )
    """)

    # Comprobar si la tabla ventas ya tenía la columna precio
    cursor.execute("PRAGMA table_info(ventas)")
    columnas = cursor.fetchall()

    nombres_columnas = [columna[1] for columna in columnas]

    if "precio" not in nombres_columnas:
        cursor.execute("ALTER TABLE ventas ADD COLUMN precio REAL")

    conexion.commit()
    conexion.close()

@app.route("/eliminar/<int:id>")
def eliminar(id):

    conexion = sqlite3.connect("inventario.db")
    cursor = conexion.cursor()

    cursor.execute(
        "DELETE FROM productos WHERE id = ?",
        (id,)
    )

    conexion.commit()
    conexion.close()

    return """
    <h1>Producto eliminado</h1>
    <a href='/inventario'>
        <button>Volver</button>
    </a>
    """
@app.route("/editar/<int:id>", methods=["GET", "POST"])
def editar(id):

    if request.method == "POST":

        nombre = request.form["nombre"]
        marca = request.form["marca"]
        costo = request.form["costo"]
        precio = request.form["precio"]

        conexion = sqlite3.connect("inventario.db")
        cursor = conexion.cursor()

        cursor.execute("""
        UPDATE productos
        SET nombre = ?,
            marca = ?,
            costo = ?,
            precio = ?
        WHERE id = ?
        """, (nombre, marca, costo, precio, id))

        conexion.commit()
        conexion.close()

        return redirect("/inventario")
        
    conexion = sqlite3.connect("inventario.db")
    cursor = conexion.cursor()

    cursor.execute(
        "SELECT * FROM productos WHERE id = ?",
        (id,)
    )

    producto = cursor.fetchone()

    conexion.close()

    pagina = f"""
    <html>

    <head>
    <link rel="stylesheet" href="/static/estilo.css">
    </head>

    <body>

    <h1>✏️ Editar producto</h1>

    <div class="formulario">

    <form method="POST">

        Nombre:
        <input type="text" name="nombre" value="{producto[1]}">

        <br><br>

        Marca:
        <input type="text" name="marca" value="{producto[2]}">

        <br><br>

        Costo:
        <input type="number" step="0.01" name="costo" value="{producto[4]}">

        <br><br>

        Precio:
       <input type="number" step="0.01" name="precio" value="{producto[5]}">
        <br><br>

        <button>Guardar cambios</button>

    </form>

    </div>

    </body>

    </html>
    """

    return pagina
@app.route("/inventario")
def inventario():

    buscar = request.args.get("buscar")

    conexion = sqlite3.connect("inventario.db")
    cursor = conexion.cursor()

    if buscar:
        cursor.execute(
            "SELECT * FROM productos WHERE nombre LIKE ?",
            ('%' + buscar + '%',)
        )
    else:
        cursor.execute("SELECT * FROM productos")

    productos = cursor.fetchall()

    conexion.close()

    tabla = """
    <html>

    <head>
    <link rel="stylesheet" href="/static/estilo.css">
    </head>

    <body>

    <h1>📦 Inventario L&R</h1>

    <a href="/">
        <button>🏠 Menú principal</button>
    </a>

<br><br>

<form method="GET">

    <div class="buscador">

        Buscar:
        <input type="text" name="buscar">

        <button type="submit">
            🔍 Buscar
        </button>

    </div>

</form>

    <br>

    <table>
<tr>
    <th>ID</th>
    <th>Nombre</th>
    <th>Marca</th>
    <th>Cantidad</th>
    <th>Costo</th>
    <th>Precio</th>
    <th>Acciones</th>
</tr>
"""

    for producto in productos:
        tabla += f"""
<tr>
    <td>{producto[0]}</td>
    <td>{producto[1]}</td>
    <td>{producto[2]}</td>
    <td>{producto[3]}</td>
    <td>{producto[4]}</td>
    <td>{producto[5]}</td>
    <td>

    <a href="/editar/{producto[0]}">
        <button>✏️ Editar</button>
    </a>

    <a href="/eliminar/{producto[0]}">
        <button class="rosa">Eliminar</button>

</td>
</tr>
"""

    tabla += """
    </body>
    </html>
    """

    return tabla
@app.route("/agregar", methods=["GET","POST"])
def agregar():
    if request.method == "POST":

        nombre = request.form["nombre"]
        marca = request.form["marca"]
        cantidad = request.form["cantidad"]
        costo = request.form["costo"]
        precio = request.form["precio"]

        conexion = sqlite3.connect("inventario.db")
        cursor = conexion.cursor()

        cursor.execute("""
            INSERT INTO productos
            (nombre, marca, cantidad, costo, precio)
            VALUES (?, ?, ?, ?, ?)
        """, (nombre, marca, cantidad, costo, precio))

        conexion.commit()
        conexion.close()

        from flask import redirect
        return redirect("/inventario")
    return """
        <html>

        <head>
        <link rel="stylesheet" href="/static/estilo.css">
        </head>

        <body>

    <h1>📦 Agregar producto</h1>

<a href="/">
    <button>🏠 Menú principal</button>
</a>

<br><br>

<div class="formulario">
<form method="POST">

    Nombre:
    <input type="text" name="nombre">

    <br><br>

    Marca:
    <input type="text" name="marca">

    <br><br>

    Cantidad:
    <input type="number" name="cantidad">

    <br><br>

    Costo:
    <input type="number" step="0.01" name="costo">

    <br><br>

    Precio:
    <input type="number" step="0.01" name="precio">

    <br><br>

    <button type="submit">
        ➕ Guardar producto
    </button>

</form>
</div>
</body>
</html>
    """
crear_base_datos()
@app.route("/entrada", methods=["GET","POST"])
def entrada():

    if request.method == "POST":

        producto = request.form["nombre"]
        nombre, marca = producto.split(" | ")
        cantidad = request.form["cantidad"]

        conexion = sqlite3.connect("inventario.db")
        cursor = conexion.cursor()

        cursor.execute(
        "SELECT id FROM productos WHERE nombre= ? AND marca= ?",
        (nombre, marca)
        )

        resultado = cursor.fetchone()

        if resultado:
            id_producto = resultado[0]
        else:
            conexion.close()
            return "<h1>Producto no encontrado</h1>"

        cursor.execute(
            """
        UPDATE productos
        SET cantidad = cantidad + ?
        WHERE id = ?
        """,
     (cantidad, id_producto)
        )

        conexion.commit()
        conexion.close()     
        return redirect("/inventario")

    conexion = sqlite3.connect("inventario.db")
    cursor = conexion.cursor()

    cursor.execute("SELECT id, nombre, marca FROM productos")

    productos = cursor.fetchall()

    conexion.close()

    pagina = """

    <html>
<head>
<link rel="stylesheet" href="/static/estilo.css">
</head>
<body>

<div class="formulario">
    <h1>Producto resurtido</h1>

    <form method="POST">

        Producto:
        <input list="productos" name="nombre">
        <datalist id="productos">
    """

    for producto in productos:

        pagina += f"""
        <option value="{producto[1]} | {producto[2]}"></option>
        """
    pagina += """
    </datalist>

    <br><br>

    Cantidad:
    <input type="number" name="cantidad">

    <br><br>

    <button type="submit">
        Registrar
    </button>

    </form>

    </div>

    </body>
    </html>
    """

    print("ENTRE A LA NUEVA PAGINA DE RESURTIDO")
    return pagina
@app.route("/salida", methods=["GET","POST"])
def salida():
    print(request.method)

    if request.method == "POST":

        producto = request.form["nombre"]
        nombre, marca = producto.split(" | ")
        cantidad = request.form["cantidad"]
        
        conexion = sqlite3.connect("inventario.db")
        cursor = conexion.cursor()

        conexion = sqlite3.connect("inventario.db")
        cursor = conexion.cursor()

        cursor.execute(
            """
            SELECT id, precio
            FROM productos
            WHERE nombre = ? AND marca = ?
            """,
            (nombre, marca)
        )

        resultado = cursor.fetchone()

        if resultado:
            id_producto = resultado[0]
            precio = resultado[1]
        else:
            conexion.close()
            return "<h1>Producto no encontrado</h1>"

        cursor.execute(
            """
            UPDATE productos
            SET cantidad = cantidad - ?
            WHERE id = ?
            """,
            (cantidad, id_producto)
        )

        cursor.execute(
           """
           INSERT INTO ventas (nombre, marca, cantidad, precio, fecha)
           VALUES (?, ?, ?, ?, datetime('now', 'localtime'))
           """,
           (nombre, marca, cantidad, precio)
        )

        conexion.commit()
        conexion.close()

        return """
        <html>

        <head>
        <link rel="stylesheet" href="/static/estilo.css">
        </head>

        <body>

        <h1>✅ Salida registrada</h1>

        <a href="/inventario">
            <button>📦 Volver al inventario</button>
        </a>

        </body>
        </html>
        """

    conexion = sqlite3.connect("inventario.db")
    cursor = conexion.cursor()

    cursor.execute("SELECT id, nombre, marca FROM productos")

    productos = cursor.fetchall()

    conexion.close()

    pagina = """
    <html>

    <head>
    <link rel="stylesheet" href="/static/estilo.css">
    </head>

    <body>

    <h1>⬇️ Ventas</h1>

    <a href="/">
        <button>🏠 Menú principal</button>
    </a>

    <br><br>

    <div class="formulario">

    <form method="POST">

        Producto:
        <input list="productos" name="nombre">

        <datalist id="productos">
    """

    for producto in productos:
        pagina += f"""
        <option value="{producto[1]} | {producto[2]}"></option>
        """

    pagina += """
    </datalist>

    <br><br>

    Cantidad:
    <input type="number" name="cantidad">

    <br><br>

    <button type="submit">
        ⬇️ Ventas
    </button>

    </form>

    </div>

    </body>
    </html>
    """

    return pagina

@app.route("/limpiar_historial", methods=["POST"])
def limpiar_historial():

    conexion = sqlite3.connect("inventario.db")
    cursor = conexion.cursor()

    cursor.execute("DELETE FROM ventas")

    conexion.commit()
    conexion.close()

    return """
    <html>

    <head>
        <meta name="viewport"
              content="width=device-width, initial-scale=1.0">
        <link rel="stylesheet"
              href="/static/estilo.css">
    </head>

    <body>

        <h1>🧹 Historial limpiado</h1>

        <p>
            Todas las ventas han sido eliminadas.
        </p>

        <p>
            El inventario no fue modificado.
        </p>

        <br>

        <a href="/historial">
            <button>📜 Volver al historial</button>
        </a>

        <br><br>

        <a href="/">
            <button>🏠 Menú principal</button>
        </a>

    </body>

    </html>
    """

@app.route("/historial")
def historial():

    dia = request.args.get("dia", "")
    mes = request.args.get("mes", "")
    año = request.args.get("año", "")

    conexion = sqlite3.connect("inventario.db")
    cursor = conexion.cursor()

    if dia and mes and año:

        fecha_busqueda = f"{año}-{mes.zfill(2)}-{dia.zfill(2)}"

        cursor.execute("""
            SELECT fecha, nombre, marca, cantidad, precio
            FROM ventas
            WHERE date(fecha) = ?
            ORDER BY id DESC
        """, (fecha_busqueda,))

    else:

        cursor.execute("""
            SELECT fecha, nombre, marca, cantidad, precio
            FROM ventas
            ORDER BY id DESC
        """)

    ventas = cursor.fetchall()

    conexion.close()

    pagina = """
    <html>

    <head>
        <link rel="stylesheet" href="/static/estilo.css">
    </head>

    <body>

    <h1>📜 Historial de ventas</h1>

    <a href="/">
        <button>🏠 Menú principal</button>
    </a>

    <br><br>

    <div class="formulario">

        <h2>Buscar por fecha</h2>

        <form method="GET">

            Día:
            <input type="number"
                   name="dia"
                   min="1"
                   max="31"
                   placeholder="DD">

            Mes:
            <input type="number"
                   name="mes"
                   min="1"
                   max="12"
                   placeholder="MM">

            Año:
            <input type="number"
                   name="año"
                   min="2000"
                   max="2100"
                   placeholder="AAAA">

            <button type="submit">
                🔍 Buscar
            </button>

        </form>

        <br>

          <a href="/historial">
        <button>📋 Mostrar todas</button>
    </a>

    <br><br>

    <form method="POST"
          action="/limpiar_historial"
          onsubmit="return confirm('⚠️ ¿Estás seguro de que quieres eliminar TODO el historial de ventas? Esta acción no se puede deshacer.');">

        <button type="submit" class="rosa">
            🧹 Limpiar historial
        </button>

    </form>

</div>
    <br>

    <table>

    <tr>
        <th>Fecha</th>
        <th>Producto</th>
        <th>Marca</th>
        <th>Cantidad</th>
        <th>Precio de venta</th>
    </tr>
    """

    for venta in ventas:

        fecha = venta[0]

        try:
            fecha_formateada = fecha[:10]
            año_fecha, mes_fecha, dia_fecha = fecha_formateada.split("-")
            hora = fecha[11:]

            fecha_formateada = f"{dia_fecha}/{mes_fecha}/{año_fecha} {hora}"

        except:
            fecha_formateada = fecha

        if venta[4] is not None:
            precio_formateado = f"${venta[4]:.2f}"
        else:
            precio_formateado = "Sin registro"

        pagina += f"""
        <tr>

            <td>{fecha_formateada}</td>

            <td>{venta[1]}</td>

            <td>{venta[2]}</td>

            <td>{venta[3]}</td>

            <td>{precio_formateado}</td>

        </tr>
        """

    pagina += """
    </table>

    </body>
    </html>
    """

    return pagina
@app.route("/exportar_excel")
def exportar_excel():

    conexion = sqlite3.connect("inventario.db")
    cursor = conexion.cursor()

    # Obtener inventario
    cursor.execute("""
        SELECT id, nombre, marca, cantidad, costo, precio
        FROM productos
        ORDER BY id
    """)

    productos = cursor.fetchall()

    # Obtener historial de ventas
    cursor.execute("""
        SELECT id, fecha, nombre, marca, cantidad, precio
        FROM ventas
        ORDER BY id
    """)

    ventas = cursor.fetchall()

    conexion.close()

    # Crear archivo Excel
    wb = Workbook()

    # Hoja de inventario
    ws_inventario = wb.active
    ws_inventario.title = "Inventario"

    ws_inventario.append([
        "ID",
        "Nombre",
        "Marca",
        "Cantidad",
        "Costo",
        "Precio"
    ])

    for producto in productos:
        ws_inventario.append(producto)

    # Hoja de historial de ventas
    ws_ventas = wb.create_sheet("Historial de ventas")

    ws_ventas.append([
        "ID",
        "Fecha",
        "Nombre",
        "Marca",
        "Cantidad",
        "Precio de venta"
    ])

    for venta in ventas:
        ws_ventas.append(venta)

    # Guardar archivo
    archivo = "inventario_respaldo.xlsx"

    wb.save(archivo)

    return send_file(
        archivo,
        as_attachment=True
    )

@app.route("/exportar_pdf", methods=["GET", "POST"])
def exportar_pdf():

    if request.method == "GET":

        return """
        <!DOCTYPE html>
        <html>

        <head>

            <meta name="viewport"
                  content="width=device-width, initial-scale=1.0">

            <title>Reporte de ventas</title>

            <link rel="stylesheet"
                  href="/static/estilo.css">

        </head>

        <body>

            <h1>📄 Reporte de ventas</h1>

            <form method="POST">

                <h2>Selecciona el periodo</h2>

                <p>Fecha inicial:</p>

                <input
                    type="date"
                    name="fecha_inicio"
                    required
                >

                <br><br>

                <p>Fecha final:</p>

                <input
                    type="date"
                    name="fecha_fin"
                    required
                >

                <br><br>

                <button type="submit">
                    Generar PDF
                </button>

            </form>

            <br>

            <a href="/">
                <button>Volver</button>
            </a>

        </body>

        </html>
        """

    fecha_inicio = request.form["fecha_inicio"]
    fecha_fin = request.form["fecha_fin"]

    conexion = sqlite3.connect("inventario.db")
    cursor = conexion.cursor()

    cursor.execute("""
        SELECT fecha, nombre, marca, cantidad, precio
        FROM ventas
        WHERE date(fecha) BETWEEN date(?) AND date(?)
        ORDER BY fecha
    """, (fecha_inicio, fecha_fin))

    ventas = cursor.fetchall()

    conexion.close()

    archivo = "reporte_ventas.pdf"

    documento = SimpleDocTemplate(
        archivo,
        pagesize=landscape(letter),
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=30
    )

    estilos = getSampleStyleSheet()

    elementos = []

    titulo = Paragraph(
        "<b>Reporte de ventas - Papelería L&R</b>",
        estilos["Title"]
    )

    elementos.append(titulo)

    periodo = Paragraph(
        f"Periodo: {fecha_inicio} al {fecha_fin}",
        estilos["Normal"]
    )

    elementos.append(periodo)
    elementos.append(Spacer(1, 20))

    datos = [
        [
            "Fecha",
            "Producto",
            "Marca",
            "Cantidad",
            "Precio de venta",
            "Total"
        ]
    ]

    total_periodo = 0

    for venta in ventas:

        fecha = venta[0]
        nombre = venta[1]
        marca = venta[2]
        cantidad = venta[3]
        precio = venta[4]

        if precio is not None:

            precio_formateado = f"${precio:.2f}"

            total = cantidad * precio

            total_formateado = f"${total:.2f}"

            total_periodo += total

        else:

            precio_formateado = "Sin registro"
            total_formateado = "Sin registro"

        datos.append([
            fecha,
            nombre,
            marca,
            cantidad,
            precio_formateado,
            total_formateado
        ])

    if len(ventas) == 0:

        datos.append([
            "No hay ventas",
            "",
            "",
            "",
            "",
            ""
        ])

    tabla = Table(
        datos,
        repeatRows=1
    )

    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (3, 1), (5, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("TOPPADDING", (0, 0), (-1, 0), 8)
    ]))

    elementos.append(tabla)

    elementos.append(Spacer(1, 20))

    total_texto = Paragraph(
        f"<b>Total de ventas del periodo: ${total_periodo:.2f}</b>",
        estilos["Heading2"]
    )

    elementos.append(total_texto)

    documento.build(elementos)

    return send_file(
        archivo,
        as_attachment=True
    )

@app.route("/restaurar_excel", methods=["GET", "POST"])
def restaurar_excel():

    if request.method == "POST":

        archivo = request.files["archivo"]

        if not archivo:
            return """
            <h2>No se seleccionó ningún archivo.</h2>
            <a href="/">Volver</a>
            """

        try:
            wb = load_workbook(archivo)

            # Verificar que exista la hoja de inventario
            if "Inventario" not in wb.sheetnames:
                return """
                <h2>Archivo no válido</h2>
                <p>El archivo debe contener una hoja llamada "Inventario".</p>
                <a href="/">Volver</a>
                """

            ws_inventario = wb["Inventario"]

            conexion = sqlite3.connect("inventario.db")
            cursor = conexion.cursor()

            try:

                # Borrar únicamente el inventario actual
                cursor.execute("DELETE FROM productos")

                # Restaurar productos
                for fila in ws_inventario.iter_rows(
                    min_row=2,
                    values_only=True
                ):

                    id_producto, nombre, marca, cantidad, costo, precio = fila

                    if nombre is None:
                        continue

                    cursor.execute("""
                        INSERT INTO productos
                        (id, nombre, marca, cantidad, costo, precio)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (
                        id_producto,
                        nombre,
                        marca,
                        cantidad,
                        costo,
                        precio
                    ))

                # Si el archivo también tiene historial,
                # restaurarlo.
                if "Historial de ventas" in wb.sheetnames:

                    ws_ventas = wb["Historial de ventas"]

                    cursor.execute("DELETE FROM ventas")

                    for fila in ws_ventas.iter_rows(
                        min_row=2,
                        values_only=True
                    ):

                        id_venta, fecha, nombre, marca, cantidad, precio = fila

                        if nombre is None:
                            continue

                        cursor.execute("""
                            INSERT INTO ventas
                            (id, nombre, marca, cantidad, precio, fecha)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, (
                            id_venta,
                            nombre,
                            marca,
                            cantidad,
                            precio,
                            fecha
                        ))

                conexion.commit()

            except Exception as error:

                conexion.rollback()
                conexion.close()

                return f"""
                <h2>Error al restaurar</h2>
                <p>{error}</p>
                <a href="/">Volver</a>
                """

            conexion.close()

            if "Historial de ventas" in wb.sheetnames:

                mensaje = """
                <h2>Restauración completada</h2>
                <p>Se restauró el inventario y el historial de ventas.</p>
                """

            else:

                mensaje = """
                <h2>Inventario restaurado correctamente</h2>
                <p>Se cargaron los productos del archivo.</p>
                <p>El historial de ventas actual no fue modificado.</p>
                """

            return mensaje + """
            <br>
            <a href="/">Volver al menú</a>
            """

        except Exception as error:

            return f"""
            <h2>Error al abrir el archivo</h2>
            <p>{error}</p>
            <a href="/">Volver</a>
            """

    return """
    <!DOCTYPE html>
    <html>
    <head>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Restaurar inventario</title>
        <link rel="stylesheet"
              href="/static/estilo.css">
    </head>

    <body>

        <h1>♻️ Restaurar</h1>

        <form method="POST"
              enctype="multipart/form-data">

            <h2>Selecciona el archivo Excel</h2>

            <input type="file"
                   name="archivo"
                   accept=".xlsx"
                   required>

            <br><br>

            <button type="submit">
                Restaurar
            </button>

        </form>

        <br>

        <a href="/">
            <button>Volver</button>
        </a>

    </body>
    </html>
    """

if __name__ == "__main__":
    app.run(debug=True)