import io
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

def generar_recibo_pdf(pedido, items, datos_vendedor):
    """
    Genera un recibo PDF profesional en memoria (BytesIO) con el logo,
    colores de marca y el desglose completo del pedido.
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    story = []
    styles = getSampleStyleSheet()

    # --- PALETA DE COLORES MARY KAY / BEAUTY STORE ---
    COLOR_PRIMARIO = colors.HexColor("#C2185B")   # Rosa Deep
    COLOR_SECUNDARIO = colors.HexColor("#222222") # Oscuro
    COLOR_FONDO_TABLA = colors.HexColor("#F9F0F5")# Rosa Claro
    COLOR_LINEA = colors.HexColor("#E8A0BF")      # Rosa Suave

    # --- ESTILOS PERSONALIZADOS ---
    style_titulo = ParagraphStyle(
        'TituloComprobante',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=20,
        leading=24,
        textColor=COLOR_PRIMARIO,
        alignment=2 # Alineación derecha
    )

    style_subtitulo = ParagraphStyle(
        'SubtituloComprobante',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        textColor=COLOR_SECUNDARIO,
        alignment=2
    )

    style_normal = ParagraphStyle(
        'TextoNormal',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=12,
        textColor=COLOR_SECUNDARIO
    )

    style_bold = ParagraphStyle(
        'TextoBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        textColor=COLOR_SECUNDARIO
    )

    style_header_tabla = ParagraphStyle(
        'HeaderTabla',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=11,
        textColor=colors.white
    )

    # --- ENCABEZADO: DATOS DE VENDEDOR Y COMPROBANTE ---
    col1 = [
        Paragraph(f"<b>{datos_vendedor['nombre']}</b>", style_bold),
        Paragraph(f"Consultora de Belleza: Analía A. Domínguez", style_normal),
        Paragraph(f"{datos_vendedor['direccion']}", style_normal),
        Paragraph(f"{datos_vendedor['localidad']} - {datos_vendedor['provincia']} (CP {datos_vendedor['codigo_postal']})", style_normal),
        Paragraph(f"WhatsApp: {datos_vendedor['telefono']}", style_normal),
        Paragraph(f"Email: analia.alejandra74@gmail.com", style_normal),
    ]

    col2 = [
        Paragraph("RECIBO DE COMPRA", style_titulo),
        Paragraph(f"<b>N° de Pedido:</b> #{pedido.id:06d}", style_subtitulo),
        Paragraph(f"<b>Fecha:</b> {pedido.fecha.strftime('%d/%m/%Y %H:%M') if pedido.fecha else ''}", style_subtitulo),
        Paragraph(f"<b>Método de Pago:</b> {pedido.metodo_pago.upper()}", style_subtitulo),
        Paragraph(f"<b>Estado:</b> {pedido.estado_pago}", style_subtitulo),
    ]

    header_table = Table([[col1, col2]], colWidths=[3.25 * inch, 4.0 * inch])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 15))

    story.append(HRFlowable(width="100%", thickness=1.5, color=COLOR_LINEA, spaceAfter=15))

    # --- INFORMACIÓN DEL CLIENTE Y ENVÍO ---
    datos_cliente = [
        [
            Paragraph("<b>DATOS DEL CLIENTE</b>", style_bold),
            Paragraph("<b>DIRECCIÓN Y DESTINO DE ENVÍO</b>", style_bold)
        ],
        [
            Paragraph(f"<b>Nombre:</b> {pedido.cliente_nombre}<br/>"
                      f"<b>DNI / CUIT:</b> {pedido.cliente_dni}<br/>"
                      f"<b>Teléfono:</b> {pedido.cliente_telefono}<br/>"
                      f"<b>Email:</b> {pedido.cliente_email or 'No especificado'}", style_normal),
            Paragraph(f"<b>Dirección:</b> {pedido.cliente_direccion}<br/>"
                      f"<b>Localidad:</b> {pedido.cliente_localidad}<br/>"
                      f"<b>Provincia:</b> {pedido.cliente_provincia}<br/>"
                      f"<b>Código Postal:</b> {pedido.cliente_codigo_postal or 'S/D'}", style_normal)
        ]
    ]

    tabla_cliente = Table(datos_cliente, colWidths=[3.625 * inch, 3.625 * inch])
    tabla_cliente.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), COLOR_FONDO_TABLA),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('PADDING', (0, 0), (-1, -1), 8),
        ('BOX', (0, 0), (-1, -1), 0.5, COLOR_LINEA),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, COLOR_LINEA),
    ]))
    story.append(tabla_cliente)
    story.append(Spacer(1, 20))

    # --- TABLA DE PRODUCTOS ---
    tabla_data = [
        [
            Paragraph("<b>Producto / Descripción</b>", style_header_tabla),
            Paragraph("<b>Cant.</b>", style_header_tabla),
            Paragraph("<b>Precio Unit.</b>", style_header_tabla),
            Paragraph("<b>Subtotal</b>", style_header_tabla)
        ]
    ]

    for item in items:
        nombre_prod = item["producto"].nombre if hasattr(item.get("producto"), "nombre") else "Producto"
        precio_unit = item["precio_unitario"]
        cantidad = item["cantidad"]
        subtotal = precio_unit * cantidad

        tabla_data.append([
            Paragraph(f"{nombre_prod}", style_normal),
            Paragraph(f"{cantidad}", style_normal),
            Paragraph(f"${precio_unit:,.2f}", style_normal),
            Paragraph(f"${subtotal:,.2f}", style_normal)
        ])

    tabla_items = Table(tabla_data, colWidths=[3.85 * inch, 0.8 * inch, 1.3 * inch, 1.3 * inch])
    tabla_items.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), COLOR_PRIMARIO),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ('PADDING', (0, 0), (-1, -1), 6),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, COLOR_FONDO_TABLA])
    ]))
    story.append(tabla_items)
    story.append(Spacer(1, 15))

    # --- RESUMEN DE TOTALES ---
    subtotal_productos = sum(i["precio_unitario"] * i["cantidad"] for i in items)
    costo_envio = pedido.costo_envio or 0.0

    totales_data = [
        [Paragraph("Subtotal Productos:", style_normal), Paragraph(f"${subtotal_productos:,.2f}", style_bold)],
        [Paragraph("Costo de Envío:", style_normal), Paragraph(f"${costo_envio:,.2f}", style_bold)],
        [Paragraph("<b>TOTAL FINAL ABONADO:</b>", style_bold), Paragraph(f"<b>${pedido.total:,.2f}</b>", ParagraphStyle('TotalRed', parent=style_bold, textColor=COLOR_PRIMARIO, fontSize=11))]
    ]

    tabla_totales = Table(totales_data, colWidths=[5.35 * inch, 1.9 * inch])
    tabla_totales.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('PADDING', (0, 0), (-1, -1), 4),
        ('LINEABOVE', (0, -1), (-1, -1), 1, COLOR_PRIMARIO)
    ]))
    story.append(tabla_totales)
    story.append(Spacer(1, 30))

    # --- PIE DE PÁGINA ---
    story.append(HRFlowable(width="100%", thickness=0.5, color=COLOR_LINEA, spaceAfter=10))
    style_pie = ParagraphStyle('Pie', parent=style_normal, fontSize=8, textColor=colors.gray, alignment=1)
    story.append(Paragraph("<b>¡Gracias por confiar en BEAUTY STORE | Mary Kay!</b>", style_pie))
    story.append(Paragraph("Este comprobante sirve como resumen oficial de tu orden realizada a través de nuestra plataforma web.", style_pie))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()