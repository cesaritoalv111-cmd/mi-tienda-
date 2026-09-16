import io
import sqlite3
import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from PIL import Image

# Librería para lectura de códigos de barra
try:
    from pyzbar.pyzbar import decode as decode_barcode
    HAS_PYZBAR = True
except ImportError:
    HAS_PYZBAR = False

# ---------------- CONFIGURACIÓN Y BD ----------------
st.set_page_config(page_title="Sistema de Gestión - Tienda", layout="wide", page_icon="🏷️")

def get_connection():
    return sqlite3.connect("tienda.db")

def init_db():
    with get_connection() as conn:
        cursor = conn.cursor()
        # Tabla Inventario con soporte de código de barras
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS inventario (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                barcode TEXT,
                sku TEXT UNIQUE,
                nombre TEXT,
                marca TEXT,
                color TEXT,
                talle TEXT,
                stock INTEGER,
                costo_compra REAL,
                costo_flete REAL,
                precio_venta REAL,
                url TEXT
            )
        """)
        # Tabla Ventas
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS ventas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                producto_id INTEGER,
                fecha TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                cantidad INTEGER,
                precio_unitario REAL,
                total_venta REAL,
                ganancia_estimada REAL,
                FOREIGN KEY(producto_id) REFERENCES inventario(id)
            )
        """)
        # Tabla Gastos Operativos
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS gastos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                concepto TEXT,
                categoria TEXT,
                monto REAL
            )
        """)
        conn.commit()

init_db()

# ---------------- FUNCIÓN DE METADATOS POR LINK ----------------
def scrape_metadata(url):
    data = {"nombre": "", "marca": "", "color": ""}
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        res = requests.get(url, headers=headers, timeout=6)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, "html.parser")
            og_title = soup.find("meta", property="og:title") or soup.find("title")
            if og_title:
                data["nombre"] = og_title.get("content", og_title.text).strip()
            og_brand = soup.find("meta", property="og:brand") or soup.find("meta", property="product:brand")
            if og_brand:
                data["marca"] = og_brand.get("content", "").strip()
            desc = soup.find("meta", property="og:description")
            desc_text = desc.get("content", "").lower() if desc else ""
            colores_comunes = ["negro", "blanco", "azul", "rojo", "verde", "gris", "beige", "black", "white", "blue", "red", "navy", "pink", "marron", "brown"]
            for c in colores_comunes:
                if c in data["nombre"].lower() or c in desc_text:
                    data["color"] = c.capitalize()
                    break
    except Exception:
        pass
    return data

# ---------------- FUNCIÓN DE EXPORTACIÓN A EXCEL ----------------
def generar_excel_completo():
    conn = get_connection()
    df_inv = pd.read_sql("SELECT barcode, sku, nombre, marca, color, talle, stock, costo_compra, costo_flete, precio_venta FROM inventario", conn)
    df_ventas = pd.read_sql("""
        SELECT v.id, v.fecha, i.sku, i.nombre, v.cantidad, v.precio_unitario, v.total_venta, v.ganancia_estimada 
        FROM ventas v LEFT JOIN inventario i ON v.producto_id = i.id
    """, conn)
    df_gastos = pd.read_sql("SELECT fecha, concepto, categoria, monto FROM gastos", conn)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_inv.to_excel(writer, sheet_name="Inventario", index=False)
        df_ventas.to_excel(writer, sheet_name="Ventas", index=False)
        df_gastos.to_excel(writer, sheet_name="Gastos", index=False)
    return output.getvalue()

# ---------------- MENÚ DE NAVEGACIÓN ----------------
menu = st.sidebar.radio("Navegación", [
    "📊 Panel General", 
    "📦 Cargar Producto (Link / Código)", 
    "📋 Inventario y Exportar Excel", 
    "🏷️ Registrar Venta (Escanear)", 
    "💸 Registrar Gastos"
])

# ---------------- 1. PANEL GENERAL ----------------
if menu == "📊 Panel General":
    st.title("📊 Resumen del Negocio")
    conn = get_connection()
    df_ventas = pd.read_sql("SELECT * FROM ventas", conn)
    df_gastos = pd.read_sql("SELECT * FROM gastos", conn)
    df_inv = pd.read_sql("SELECT * FROM inventario", conn)
    
    total_ingresos = df_ventas["total_venta"].sum() if not df_ventas.empty else 0.0
    margen_bruto = df_ventas["ganancia_estimada"].sum() if not df_ventas.empty else 0.0
    total_gastos = df_gastos["monto"].sum() if not df_gastos.empty else 0.0
    ganancia_neta = margen_bruto - total_gastos
    valor_stock = ((df_inv["costo_compra"] + df_inv["costo_flete"]) * df_inv["stock"]).sum() if not df_inv.empty else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Ingresos Totales", f"${total_ingresos:,.2f}")
    c2.metric("Gastos Operativos", f"${total_gastos:,.2f}")
    c3.metric("Ganancia Neta Real", f"${ganancia_neta:,.2f}", delta=f"{ganancia_neta:,.2f}")
    c4.metric("Capital en Stock", f"${valor_stock:,.2f}")

    st.markdown("---")
    st.subheader("Últimas 10 Ventas Realizadas")
    st.dataframe(df_ventas.tail(10), use_container_width=True)

# ---------------- 2. CARGA DE PRODUCTO ----------------
elif menu == "📦 Cargar Producto (Link / Código)":
    st.title("📦 Agregar Nuevo Producto")

    # Opción: Escanear código de barras para dar de alta
    codigo_detectado = ""
    with st.expander("📷 Escanear Código de Barras / QR con la cámara"):
        if not HAS_PYZBAR:
            st.info("Para usar la cámara necesitas instalar pyzbar (`pip install pyzbar`). Mientras tanto, puedes escribir el código manualmente.")
        else:
            img_file = st.camera_input("Apunta el código de barras a la cámara")
            if img_file is not None:
                img = Image.open(img_file)
                decoded_objs = decode_barcode(img)
                if decoded_objs:
                    codigo_detectado = decoded_objs[0].data.decode("utf-8")
                    st.success(f"Código detectado: **{codigo_detectado}**")
                else:
                    st.warning("No se detectó un código legible. Acerca más el producto.")

    url_input = st.text_input("Enlace del producto (opcional para autocompletar):", placeholder="https://...")
    auto_data = {"nombre": "", "marca": "", "color": ""}
    if st.button("Extraer Datos del Enlace"):
        if url_input:
            with st.spinner("Consultando información en línea..."):
                auto_data = scrape_metadata(url_input)
                st.success("Información extraída. Verifica los campos a continuación:")
        else:
            st.warning("Pega un enlace web primero.")

    with st.form("form_alta"):
        col1, col2 = st.columns(2)
        with col1:
            barcode_val = st.text_input("Código de Barras / EAN", value=codigo_detectado)
            sku = st.text_input("Código Interno / SKU *", placeholder="POLO-TH-01")
            nombre = st.text_input("Nombre del Producto *", value=auto_data["nombre"])
            marca = st.text_input("Marca", value=auto_data["marca"])
            color = st.text_input("Color", value=auto_data["color"])
            talle = st.text_input("Talle / Medida", placeholder="S, M, L, XL, 42mm...")
        with col2:
            stock = st.number_input("Cantidad en Stock", min_value=1, value=1)
            costo_compra = st.number_input("Precio de Compra Unitario ($)", min_value=0.0, format="%.2f")
            costo_flete = st.number_input("Costo de Envío / Courier Unitario ($)", min_value=0.0, format="%.2f")
            precio_venta = st.number_input("Precio de Venta al Público ($)", min_value=0.0, format="%.2f")

        guardar = st.form_submit_button("Guardar en Inventario")
        if guardar:
            if not sku or not nombre:
                st.error("El SKU y el Nombre son obligatorios.")
            else:
                try:
                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute("""
                        INSERT INTO inventario (barcode, sku, nombre, marca, color, talle, stock, costo_compra, costo_flete, precio_venta, url)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (barcode_val, sku, nombre, marca, color, talle, stock, costo_compra, costo_flete, precio_venta, url_input))
                    conn.commit()
                    st.success(f"Producto '{nombre}' añadido con éxito.")
                except sqlite3.IntegrityError:
                    st.error("Ya existe un producto registrado con ese SKU.")

# ---------------- 3. INVENTARIO Y EXCEL ----------------
elif menu == "📋 Inventario y Exportar Excel":
    st.title("📋 Inventario y Reportes")
    conn = get_connection()
    df_inv = pd.read_sql("SELECT barcode, sku, nombre, marca, color, talle, stock, costo_compra, costo_flete, precio_venta FROM inventario", conn)

    c_btn, _ = st.columns([1, 3])
    with c_btn:
        excel_data = generar_excel_completo()
        st.download_button(
            label="📥 Descargar Reporte Completo en Excel",
            data=excel_data,
            file_name="reporte_tienda_completo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.subheader("Lista de Productos")
    st.dataframe(df_inv, use_container_width=True)

# ---------------- 4. REGISTRAR VENTA (ESCANEAR O MANUAL) ----------------
elif menu == "🏷️ Registrar Venta (Escanear)":
    st.title("🏷️ Registro de Venta")
    conn = get_connection()
    df_disp = pd.read_sql("SELECT id, barcode, sku, nombre, stock, costo_compra, costo_flete, precio_venta FROM inventario WHERE stock > 0", conn)

    if df_disp.empty:
        st.info("No hay productos con stock disponible.")
    else:
        # Modo selección o escáner
        codigo_busqueda = ""
        with st.expander("📷 Escanear con la cámara para cobrar"):
            if HAS_PYZBAR:
                cam_venta = st.camera_input("Apunta el producto")
                if cam_venta:
                    img_v = Image.open(cam_venta)
                    dec = decode_barcode(img_v)
                    if dec:
                        codigo_busqueda = dec[0].data.decode("utf-8")
                        st.success(f"Código detectado: {codigo_busqueda}")
            else:
                st.info("pyzbar no detectado. Puedes ingresar el código o seleccionar el producto de la lista.")

        # Buscar si coincide con el código escaneado
        item_index = 0
        opciones = [f"{row['sku']} | {row['nombre']} | Stock: {row['stock']}" for _, row in df_disp.iterrows()]
        
        if codigo_busqueda:
            matches = df_disp[df_disp["barcode"] == codigo_busqueda]
            if not matches.empty:
                target_sku = matches.iloc[0]["sku"]
                for idx, op in enumerate(opciones):
                    if op.startswith(target_sku):
                        item_index = idx
                        break

        seleccion = st.selectbox("Selecciona o confirma el producto:", opciones, index=item_index)
        prod_sel = df_disp.iloc[opciones.index(seleccion)]

        col1, col2 = st.columns(2)
        with col1:
            cant = st.number_input("Cantidad a vender", min_value=1, max_value=int(prod_sel["stock"]), value=1)
            precio_cobrado = st.number_input("Precio unitario cobrado ($)", value=float(prod_sel["precio_venta"]))
        with col2:
            costo_unit = prod_sel["costo_compra"] + prod_sel["costo_flete"]
            ganancia_op = (precio_cobrado - costo_unit) * cant
            st.metric("Total Cobrado", f"${(precio_cobrado * cant):,.2f}")
            st.metric("Ganancia Estimada de la Venta", f"${ganancia_op:,.2f}")

        if st.button("Cobrar y Descontar Stock"):
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO ventas (producto_id, cantidad, precio_unitario, total_venta, ganancia_estimada)
                VALUES (?, ?, ?, ?, ?)
            """, (prod_sel["id"], cant, precio_cobrado, precio_cobrado * cant, ganancia_op))
            cursor.execute("UPDATE inventario SET stock = stock - ? WHERE id = ?", (cant, prod_sel["id"]))
            conn.commit()
            st.success("¡Venta completada! El inventario y las ganancias se actualizaron.")

# ---------------- 5. REGISTRAR GASTOS ----------------
elif menu == "💸 Registrar Gastos":
    st.title("💸 Registro de Gastos Operativos")
    with st.form("form_gastos"):
        c1, c2 = st.columns(2)
        with c1:
            concepto = st.text_input("Concepto del gasto", placeholder="Alquiler, Pauta en Instagram, Bolsas...")
            categoria = st.selectbox("Categoría", ["Publicidad / Marketing", "Logística / Flete", "Packaging", "Alquiler / Servicios", "Otros"])
        with c2:
            monto = st.number_input("Monto total ($)", min_value=0.0, format="%.2f")

        guardar_gasto = st.form_submit_button("Registrar Gasto")
        if guardar_gasto:
            if not concepto or monto <= 0:
                st.error("Ingresa un concepto y un monto mayor a cero.")
            else:
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("INSERT INTO gastos (concepto, categoria, monto) VALUES (?, ?, ?)", (concepto, categoria, monto))
                conn.commit()
                st.success("Gasto guardado correctamente.")

    st.subheader("Historial de Gastos")
    conn = get_connection()
    df_gastos = pd.read_sql("SELECT fecha, concepto, categoria, monto FROM gastos ORDER BY id DESC", conn)
    st.dataframe(df_gastos, use_container_width=True)
