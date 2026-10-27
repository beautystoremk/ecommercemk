import os
import json
import io
import urllib.parse
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import mercadopago
from fastapi import FastAPI, Request, Depends, Form, HTTPException, status, Cookie, File, UploadFile, Query
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates
from sqlalchemy.orm import Session
from supabase import create_client, Client
import models, database
import unicodedata
import re
from datetime import datetime, timedelta
from email.mime.application import MIMEApplication
from pdf_generator import generar_recibo_pdf
from openpyxl import Workbook
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from fastapi import Response
from sqlalchemy.exc import IntegrityError

os.makedirs("static", exist_ok=True)
os.makedirs("templates", exist_ok=True)
os.makedirs("templates/admin", exist_ok=True)

models.Base.metadata.create_all(bind=database.engine)

app = FastAPI(title="MARY KAY Beauty Store")

SUPABASE_URL = "https://ennfuhqoqxroelcfklnk.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6I**"

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

EMAIL_EMISOR = "analia.alejandra74@gmail.com"
EMAIL_PASSWORD = "tirg degu blon vsei"
EMAIL_DESTINO = "analia.alejandra74@gmail.com"

DATOS_VENDEDOR = {
    "nombre": "BEAUTY STORE | Mary Kay",
    "direccion": "SANDWICH DEL SUR 6050 | CARRODILLA",
    "localidad": "Luján de Cuyo",
    "provincia": "Mendoza",
    "codigo_postal": "5505",
    "telefono": "+54 9 2616142896"
}

app.mount("/static", StaticFiles(directory="static"), name="static")

class GlobalJinja2Templates(Jinja2Templates):
    def TemplateResponse(self, *args, **kwargs):
        request = kwargs.get("request")
        name = kwargs.get("name")
        context = kwargs.get("context")
        status_code = kwargs.get("status_code", 200)
        headers = kwargs.get("headers")
        media_type = kwargs.get("media_type", "text/html")

        for arg in args:
            if isinstance(arg, Request):
                request = arg
            elif isinstance(arg, str) and not name:
                name = arg
            elif isinstance(arg, dict) and not context:
                context = arg

        if not context:
            for k, v in kwargs.items():
                if isinstance(v, dict) and k not in ("headers",):
                    context = v
                    break

        if context is None:
            context = {}

        if request and hasattr(request, "state"):
            if hasattr(request.state, "carrusel_items"):
                context.setdefault("carrusel_items", request.state.carrusel_items)
            if hasattr(request.state, "banners_promos"):
                context.setdefault("banners_promos", request.state.banners_promos)
            if hasattr(request.state, "skyscraper"):
                context.setdefault("skyscraper", request.state.skyscraper)
            if hasattr(request.state, "texto_publicidad"):
                context.setdefault("texto_publicidad", request.state.texto_publicidad)

        if request and "request" not in context:
            context["request"] = request

        template = self.env.get_template(name)
        content = template.render(context)
        return HTMLResponse(content, status_code=status_code, headers=headers, media_type=media_type)

templates = GlobalJinja2Templates(directory="templates")

def inicializar_provincias_envio(db: Session):
    provincias_arg = [
        ("Mendoza", 3000.0, 50000.0, True),
        ("Buenos Aires", 9500.0, None, True),
        ("Catamarca", 9500.0, None, True),
        ("Chaco", 9500.0, None, True),
        ("Chubut", 9500.0, None, True),
        ("Ciudad Autónoma de Buenos Aires", 9500.0, None, True),
        ("Córdoba", 9500.0, None, True),
        ("Corrientes", 9500.0, None, True),
        ("Entre Ríos", 9500.0, None, True),
        ("Formosa", 9500.0, None, True),
        ("Jujuy", 9500.0, None, True),
        ("La Pampa", 9500.0, None, True),
        ("La Rioja", 9500.0, None, True),
        ("Misiones", 9500.0, None, True),
        ("Neuquén", 9500.0, None, True),
        ("Río Negro", 9500.0, None, True),
        ("Salta", 9500.0, None, True),
        ("San Juan", 9500.0, None, True),
        ("San Luis", 9500.0, None, True),
        ("Santa Cruz", 9500.0, None, True),
        ("Santa Fe", 9500.0, None, True),
        ("Santiago del Estero", 9500.0, None, True),
        ("Tierra del Fuego", 9500.0, None, True),
        ("Tucumán", 9500.0, None, True),
    ]
    try:
        for prov, costo, gratis, activo in provincias_arg:
            existe = db.query(models.ProvinciaEnvio).filter(models.ProvinciaEnvio.provincia == prov).first()
            if not existe:
                db.add(models.ProvinciaEnvio(provincia=prov, costo=costo, gratis_desde=gratis, activo=activo))
        db.commit()
    except Exception as e:
        db.rollback()
        print(f"Aviso en inicialización de provincias: {e}")

def inicializar_configuracion_pago(db: Session):
    try:
        config = db.query(models.ConfiguracionPago).first()
        if not config:
            nueva_config = models.ConfiguracionPago(
                transferencia_activo=True,
                transferencia_porcentaje=10.0,
                titular="Dominguez Analia Alejandra.",
                cuit="27-23893486-4",
                alias="ana.alejandra.1974",
                banco="Mercado Pago",
                texto_publicidad="✨ 💖 ¡Bienvenidos a Beauty Store Mary Kay! 🛍️ Envíos gratis en compras seleccionadas y asesoría personalizada online 🚚 ✨"
            )
            db.add(nueva_config)
            db.commit()
    except Exception as e:
        db.rollback()
        print(f"Aviso en inicialización de config pago: {e}")

def limpiar_precio(precio_val) -> float:
    if isinstance(precio_val, (int, float)):
        return float(precio_val)
    limpio = str(precio_val).strip()
    # Maneja formato con puntos de miles y coma decimal (ej: 102.000,50 o 102000)
    if "." in limpio and "," in limpio:
        limpio = limpio.replace(".", "").replace(",", ".")
    elif "," in limpio:
        limpio = limpio.replace(",", ".")
    elif limpio.count(".") > 1:
        limpio = limpio.replace(".", "")
    try:
        return float(limpio)
    except ValueError:
        return 0.0

@app.on_event("startup")
def startup_event():
    db = database.SessionLocal()
    try:
        inicializar_provincias_envio(db)
        inicializar_configuracion_pago(db)
    finally:
        db.close()

def obtener_config_pago(db: Session):
    config = db.query(models.ConfiguracionPago).first()
    if not config:
        return {
            "transferencia_activo": True,
            "descuento_porcentaje": 10.0,
            "titular": "Dominguez Analia Alejandra.",
            "cuit": "27-23893486-4",
            "alias": "ana.alejandra.1974",
            "banco": "Mercado Pago",
            "texto_publicidad": "",
            "descuento_volumen_activo": False,
            "descuento_volumen_minimo": 0.0,
            "descuento_volumen_porcentaje": 0.0
        }
    return {
        "transferencia_activo": config.transferencia_activo,
        "descuento_porcentaje": config.transferencia_porcentaje,
        "titular": config.titular,
        "cuit": config.cuit,
        "alias": config.alias,
        "banco": config.banco,
        "texto_publicidad": config.texto_publicidad or "",
        "descuento_volumen_activo": config.descuento_volumen_activo,
        "descuento_volumen_minimo": config.descuento_volumen_minimo,
        "descuento_volumen_porcentaje": config.descuento_volumen_porcentaje
    }

@app.middleware("http")
async def add_global_context(request: Request, call_next):
    db = database.SessionLocal()
    try:
        carrusel_items = db.query(models.CarruselItem).filter(models.CarruselItem.activo == True).order_by(models.CarruselItem.orden.asc()).all()
        request.state.carrusel_items = carrusel_items
        
        banners_promos = db.query(models.BannerPromocional).filter(models.BannerPromocional.activo == True).order_by(models.BannerPromocional.orden.asc()).all()
        request.state.banners_promos = banners_promos

        skyscraper = db.query(models.BannerSkyscraper).filter(models.BannerSkyscraper.activo == True).first()
        request.state.skyscraper = skyscraper

        config_pago = db.query(models.ConfiguracionPago).first()
        request.state.texto_publicidad = config_pago.texto_publicidad if config_pago else ""
    except Exception as e:
        print(f"Error cargando contexto global: {e}")
        request.state.carrusel_items = []
        request.state.banners_promos = []
        request.state.skyscraper = None
        request.state.texto_publicidad = ""
    finally:
        db.close()
    
    response = await call_next(request)
    return response

mp_sdk = mercadopago.SDK("APP_USR-7534685542901456-102019-de863ce707ea2161eab7caeaa91e7029-586616617")

ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "admin123"
COOKIE_ADMIN_KEY = "admin_session"

def verificar_admin(admin_session: str = Cookie(None)):
    if admin_session != "sesion_activa_ok":
        raise HTTPException(
            status_code=status.HTTP_303_SEE_OTHER,
            headers={"Location": "/admin/login"}
        )
    return True

@app.get("/", response_class=HTMLResponse)
def home(request: Request, inscrito: bool = Query(False), db: Session = Depends(database.get_db)):
    productos = db.query(models.Producto).filter(models.Producto.activo == True).all()
    posts = db.query(models.BlogPost).filter(models.BlogPost.publicado == True).order_by(models.BlogPost.fecha.desc()).limit(3).all()
    
    evento_activo = db.query(models.EventoReunion).filter(models.EventoReunion.activo == True).order_by(models.EventoReunion.id.desc()).first()
    config_pago = obtener_config_pago(db)

    total_inscriptos = 0
    cupos_disponibles = 0
    porcentaje_lleno = 0
    temario_lista = []

    if evento_activo:
        total_inscriptos = len(evento_activo.inscripciones)
        cupos_disponibles = max(0, evento_activo.cupo_maximo - total_inscriptos)
        if evento_activo.cupo_maximo > 0:
            porcentaje_lleno = min(100, int((total_inscriptos / evento_activo.cupo_maximo) * 100))
        if evento_activo.temario:
            temario_lista = [t.strip() for t in evento_activo.temario.split("\n") if t.strip()]

    return templates.TemplateResponse(request, "index.html", {
        "productos": productos, 
        "posts": posts,
        "evento": evento_activo,
        "datos_transferencia": config_pago,
        "total_inscriptos": total_inscriptos,
        "cupos_disponibles": cupos_disponibles,
        "porcentaje_lleno": porcentaje_lleno,
        "temario_lista": temario_lista,
        "inscrito": inscrito
    })

@app.post("/eventos/{evento_id}/inscribir")
def inscribir_evento(
    evento_id: int,
    nombre_cliente: str = Form(...),
    telefono: str = Form(...),
    email: str = Form(None),
    db: Session = Depends(database.get_db)
):
    evento = db.query(models.EventoReunion).filter(models.EventoReunion.id == evento_id, models.EventoReunion.activo == True).first()
    if not evento:
        raise HTTPException(status_code=404, detail="El evento no existe o ha finalizado.")

    total_inscriptos = len(evento.inscripciones)
    if total_inscriptos >= evento.cupo_maximo:
        raise HTTPException(status_code=400, detail="¡Lo sentimos! Los cupos para esta reunión ya están agotados.")

    nueva_inscripcion = models.InscripcionEvento(
        evento_id=evento_id,
        nombre_cliente=nombre_cliente.strip(),
        telefono=telefono.strip(),
        email=email.strip() if email else None,
        fecha_inscripcion=datetime.utcnow() - timedelta(hours=3)
    )
    db.add(nueva_inscripcion)
    db.commit()

    # Enviar correo de notificación al administrador
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_EMISOR
        msg['To'] = EMAIL_DESTINO
        msg['Subject'] = f"💖 ¡Nueva Inscripción a Reunión! - {evento.titulo}"
        cuerpo = f"Inscripción recibida:\n\nEvento: {evento.titulo}\nNombre: {nombre_cliente}\nWhatsApp: {telefono}\nEmail: {email or 'No provisto'}"
        msg.attach(MIMEText(cuerpo, 'plain'))
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_EMISOR, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
    except Exception as e:
        print(f"Error al enviar notificación de inscripción: {e}")

    # Redirigir al WhatsApp de la oradora con el mensaje pre-llenado
    mensaje_wa = (
        f"¡Hola! Me acabo de inscribir al evento *{evento.titulo}*.\n\n"
        f"👤 *Nombre:* {nombre_cliente}\n"
        f"📱 *Teléfono:* {telefono}\n"
        f"📧 *Email:* {email if email else 'No proporcionado'}\n\n"
        f"Quedo a la espera de la confirmación e información del evento. ¡Muchas gracias!"
    )
    numero_oradora = "5492616142896"
    mensaje_encoded = urllib.parse.quote(mensaje_wa)
    whatsapp_url = f"https://api.whatsapp.com/send?phone={numero_oradora}&text={mensaje_encoded}"

    return RedirectResponse(url=whatsapp_url, status_code=status.HTTP_303_SEE_OTHER)

@app.get("/blog", response_class=HTMLResponse)
def blog_index(request: Request, db: Session = Depends(database.get_db)):
    posts = db.query(models.BlogPost).filter(models.BlogPost.publicado == True).order_by(models.BlogPost.fecha.desc()).all()
    return templates.TemplateResponse(request, "blog.html", {"posts": posts})

@app.get("/blog/{post_id}", response_class=HTMLResponse)
def blog_detalle(request: Request, post_id: int, db: Session = Depends(database.get_db)):
    post = db.query(models.BlogPost).filter(models.BlogPost.id == post_id).first()
    if not post:
        raise HTTPException(status_code=404, detail="Artículo no encontrado")
    return templates.TemplateResponse(request, "blog_post.html", {"post": post})

@app.get("/contacto", response_class=HTMLResponse)
def contacto(request: Request, enviado: bool = False):
    return templates.TemplateResponse(request, "contacto.html", {"enviado": enviado})

@app.post("/contacto")
def enviar_contacto(
    nombre: str = Form(...),
    email: str = Form(...),
    mensaje: str = Form(...)
):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_EMISOR
        msg['To'] = EMAIL_DESTINO
        msg['Subject'] = f"Nuevo mensaje de contacto de {nombre}"
        cuerpo = f"Nombre: {nombre}\nCorreo: {email}\nMensaje:\n{mensaje}"
        msg.attach(MIMEText(cuerpo, 'plain'))
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_EMISOR, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
    except Exception as e:
        print(f"Error al enviar correo de contacto: {e}")
    return RedirectResponse(url="/contacto?enviado=true", status_code=status.HTTP_303_SEE_OTHER)

def enviar_correo_nuevo_pedido(pedido, items):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_EMISOR
        msg['To'] = EMAIL_DESTINO
        msg['Subject'] = f"🚨 ¡Nuevo Pedido #{pedido.id} Recibido! - BEAUTY STORE"
        for item in items:
            subtotal = item["precio_unitario"] * item["cantidad"]
        cuerpo = f"Pedido ID: #{pedido.id}\nCliente: {pedido.cliente_nombre}\nTotal: ${pedido.total:,.2f}"
        msg.attach(MIMEText(cuerpo, 'plain'))
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_EMISOR, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
    except Exception as e:
        print(f"Error al enviar correo de notificación: {e}")

def enviar_correo_cliente_pedido(pedido, items):
    try:
        if not pedido.cliente_email:
            return
            
        msg = MIMEMultipart()
        msg['From'] = EMAIL_EMISOR
        msg['To'] = pedido.cliente_email
        msg['Subject'] = f"✨ ¡Gracias por tu compra en BEAUTY STORE! (Pedido #{pedido.id})"

        cuerpo_html = f"""
        <html>
        <body style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">
            <div style="max-width: 600px; margin: 0 auto; border: 1px solid #e8a0bf; border-radius: 12px; padding: 20px; background-color: #fafafa;">
                <h2 style="color: #c2185b; text-align: center;">BEAUTY STORE | Mary Kay</h2>
                <p>¡Hola <strong>{pedido.cliente_nombre}</strong>!</p>
                <p>Muchas gracias por elegirnos. Hemos registrado tu pedido exitosamente en nuestra tienda web.</p>
                
                <div style="background-color: #ffffff; padding: 15px; border-radius: 8px; border: 1px solid #f9f0f5; margin: 15px 0;">
                    <h3 style="color: #222; margin-top: 0;">Resumen del Pedido #{pedido.id}</h3>
                    <p><strong>Total Final:</strong> ${pedido.total:,.2f}<br/>
                    <strong>Método de Pago:</strong> {pedido.metodo_pago.upper()}<br/>
                    <strong>Dirección de Envío:</strong> {pedido.cliente_direccion}, {pedido.cliente_localidad} ({pedido.cliente_provincia})</p>
                </div>

                <p style="text-align: center; color: #666; font-size: 13px;">
                    📎 <strong>Adjunto a este correo encontrarás el recibo formal de tu compra en formato PDF.</strong>
                </p>

                <div style="text-align: center; margin-top: 20px;">
                    <a href="https://wa.me/5492616142896" style="background-color: #25d366; color: white; padding: 10px 20px; text-decoration: none; border-radius: 8px; font-weight: bold; display: inline-block;">
                        Contactar por WhatsApp
                    </a>
                </div>
            </div>
        </body>
        </html>
        """
        msg.attach(MIMEText(cuerpo_html, 'html'))

        # -------------------------------------------------------------------------
        # COMENTAR ESTE BLOQUE PARA QUE NO SE ADJUNTE EL PDF AUTOMÁTICAMENTE:
        # -------------------------------------------------------------------------
        # try:
        #     pdf_bytes = generar_recibo_pdf(pedido, items, DATOS_VENDEDOR)
        #     adjunto_pdf = MIMEApplication(pdf_bytes, _subtype="pdf")
        #     adjunto_pdf.add_header('Content-Disposition', 'attachment', filename=f"Recibo_BeautyStore_Pedido_{pedido.id}.pdf")
        #     msg.attach(adjunto_pdf)
        # except Exception as pdf_err:
        #     print(f"Error generando adjunto PDF: {pdf_err}")

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_EMISOR, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
    except Exception as e:
        print(f"Error al enviar correo de confirmación al cliente: {e}")

@app.get("/carrito", response_class=HTMLResponse)
def ver_carrito(request: Request, db: Session = Depends(database.get_db)):
    provincias_envio = db.query(models.ProvinciaEnvio).filter(models.ProvinciaEnvio.activo == True).all()
    datos_pago = obtener_config_pago(db)
    
    reglas_envio_dict = {
        env.provincia: {
            "costo": env.costo,
            "gratis_desde": env.gratis_desde
        } for env in provincias_envio
    }
    
    return templates.TemplateResponse(request, "carrito.html", {
        "datos_transferencia": datos_pago,
        "provincias_envio": provincias_envio,
        "reglas_envio_json": json.dumps(reglas_envio_dict)
    })

@app.get("/producto/{producto_id}", response_class=HTMLResponse)
def detalle_producto(request: Request, producto_id: int, db: Session = Depends(database.get_db)):
    producto = db.query(models.Producto).filter(models.Producto.id == producto_id).first()
    if not producto:
        raise HTTPException(status_code=404, detail="Producto no encontrado")
    resenas = db.query(models.Resena).filter(
        models.Resena.producto_id == producto_id,
        models.Resena.aprobado == True
    ).order_by(models.Resena.fecha.desc()).all()
    promedio_estrellas = round(sum(r.calificacion for r in resenas) / len(resenas), 1) if resenas else 0.0
    
    config_pago = obtener_config_pago(db)
    
    precio_base = producto.precio
    if producto.descuento_porcentaje and producto.descuento_porcentaje > 0:
        precio_base = producto.precio * (1 - producto.descuento_porcentaje / 100.0)

    desc_transf = config_pago["descuento_porcentaje"] if config_pago["transferencia_activo"] else 0.0
    precio_transferencia = precio_base * (1 - desc_transf / 100.0)

    return templates.TemplateResponse(request, "producto.html", {
        "producto": producto,
        "precio_base": precio_base,
        "precio_transferencia": precio_transferencia,
        "datos_transferencia": config_pago,
        "resenas": resenas,
        "promedio_estrellas": promedio_estrellas,
        "total_resenas": len(resenas)
    })

@app.post("/producto/{producto_id}/resena")
def agregar_resena(
    producto_id: int,
    nombre_cliente: str = Form(...),
    calificacion: int = Form(...),
    comentario: str = Form(...),
    db: Session = Depends(database.get_db)
):
    producto = db.query(models.Producto).filter(models.Producto.id == producto_id).first()
    if not producto:
        raise HTTPException(status_code=404, detail="Producto no encontrado")
    calificacion = max(1, min(5, calificacion))
    nueva_resena = models.Resena(
        producto_id=producto_id,
        nombre_cliente=nombre_cliente.strip(),
        calificacion=calificacion,
        comentario=comentario.strip(),
        aprobado=True,
        fecha=datetime.utcnow() - timedelta(hours=3)
    )
    db.add(nueva_resena)
    db.commit()
    return RedirectResponse(url=f"/producto/{producto_id}", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/checkout", response_class=HTMLResponse)
def procesar_checkout(
    request: Request,
    cart_data: str = Form(...),
    cliente_nombre: str = Form(...),
    cliente_email: str = Form(...),
    cliente_telefono: str = Form(...),
    cliente_dni: str = Form(...),
    cliente_direccion: str = Form(...),
    cliente_localidad: str = Form(...),
    cliente_codigo_postal: str = Form(""),
    cliente_provincia: str = Form(...),
    metodo_pago: str = Form(...),
    cupon_codigo: str = Form(None),
    db: Session = Depends(database.get_db)
):
    try:
        items = json.loads(cart_data)
    except Exception:
        raise HTTPException(status_code=400, detail="Datos del carrito inválidos")

    if not items:
        raise HTTPException(status_code=400, detail="El carrito está vacío")

    total_bruto = 0.0
    items_procesados = []

    for item in items:
        prod = db.query(models.Producto).filter(models.Producto.id == item["id"]).first()
        if not prod or prod.stock < item["cantidad"]:
            raise HTTPException(status_code=400, detail=f"Stock insuficiente para: {prod.nombre if prod else 'Producto'}")
        
        precio_unitario = prod.precio
        if prod.descuento_porcentaje and prod.descuento_porcentaje > 0:
            precio_unitario = prod.precio * (1 - prod.descuento_porcentaje / 100.0)

        subtotal_item = precio_unitario * item["cantidad"]
        total_bruto += subtotal_item
        items_procesados.append({
            "producto": prod, 
            "cantidad": item["cantidad"], 
            "precio_unitario": precio_unitario,
            "subtotal": subtotal_item
        })

    monto_descuento_cupon = 0.0
    cupon_obj = None
    if cupon_codigo and cupon_codigo.strip():
        codigo_clean = cupon_codigo.replace(" ", "").strip().upper()
        cupon_obj = db.query(models.Cupon).filter(
            models.Cupon.codigo == codigo_clean, 
            models.Cupon.activo == True
        ).first()

        if cupon_obj:
            expirado = False
            if cupon_obj.fecha_expiracion:
                fecha_actual = datetime.utcnow() - timedelta(hours=3)
                if fecha_actual > cupon_obj.fecha_expiracion:
                    expirado = True

            limite_alcanzado = False
            if cupon_obj.usos_maximos is not None and cupon_obj.usos_actuales >= cupon_obj.usos_maximos:
                limite_alcanzado = True

            subtotal_producto_especifico = 0.0
            producto_valido = True
            if cupon_obj.producto_id:
                producto_valido = False
                for it in items_procesados:
                    if it["producto"].id == cupon_obj.producto_id:
                        producto_valido = True
                        subtotal_producto_especifico += it["subtotal"]

            if not expirado and not limite_alcanzado and producto_valido and total_bruto >= cupon_obj.monto_minimo:
                base_descuento = subtotal_producto_especifico if cupon_obj.producto_id else total_bruto

                if cupon_obj.tipo_descuento == "porcentaje":
                    monto_descuento_cupon = base_descuento * (cupon_obj.valor / 100.0)
                else:
                    monto_descuento_cupon = min(cupon_obj.valor, base_descuento)
                
                cupon_obj.usos_actuales = (cupon_obj.usos_actuales or 0) + 1

    subtotal_con_cupon = max(0.0, total_bruto - monto_descuento_cupon)

    # Obtener configuración de pago para aplicar el descuento por volumen
    config_pago_db = db.query(models.ConfiguracionPago).first()

    # Cálculo del descuento automático por volumen
    monto_descuento_volumen = 0.0
    if config_pago_db and config_pago_db.descuento_volumen_activo:
        if total_bruto >= config_pago_db.descuento_volumen_minimo:
            monto_descuento_volumen = total_bruto * (config_pago_db.descuento_volumen_porcentaje / 100.0)

    subtotal_con_descuentos = max(0.0, subtotal_con_cupon - monto_descuento_volumen)

    config_envio = db.query(models.ProvinciaEnvio).filter(
        models.ProvinciaEnvio.provincia == cliente_provincia,
        models.ProvinciaEnvio.activo == True
    ).first()

    if not config_envio:
        raise HTTPException(status_code=400, detail="La provincia seleccionada no está disponible para envíos actualmente.")

    costo_envio = 0.0
    if config_envio.gratis_desde and subtotal_con_descuentos >= config_envio.gratis_desde:
        costo_envio = 0.0
    else:
        costo_envio = config_envio.costo

    transferencia_activa = config_pago_db.transferencia_activo if config_pago_db else True
    porcentaje_transf = config_pago_db.transferencia_porcentaje if config_pago_db else 10.0

    descuento_factor = 1.0
    if metodo_pago == "transferencia" and transferencia_activa:
        descuento_factor = (1 - porcentaje_transf / 100.0)

    subtotal_con_descuento_transf = subtotal_con_descuentos * descuento_factor
    total_final = subtotal_con_descuento_transf + costo_envio

    nuevo_pedido = models.Pedido(
        cliente_nombre=cliente_nombre,
        cliente_email=cliente_email,
        cliente_telefono=cliente_telefono,
        cliente_dni=cliente_dni,
        cliente_direccion=cliente_direccion,
        cliente_localidad=cliente_localidad,
        cliente_provincia=cliente_provincia,
        cliente_codigo_postal=cliente_codigo_postal,
        metodo_pago=metodo_pago,
        estado_pago="Pendiente",
        total=total_final,
        costo_envio=costo_envio,
        fecha=datetime.utcnow() - timedelta(hours=3)
    )
    db.add(nuevo_pedido)
    db.commit()
    db.refresh(nuevo_pedido)

    for item in items_procesados:
        precio_aplicado = item["precio_unitario"] * descuento_factor
        detalle = models.PedidoDetalle(
            pedido_id=nuevo_pedido.id,
            producto_id=item["producto"].id,
            cantidad=item["cantidad"],
            precio_unitario=precio_aplicado
        )
        db.add(detalle)
        item["producto"].stock -= item["cantidad"]
    
    db.commit()

    enviar_correo_nuevo_pedido(nuevo_pedido, items_procesados)
    enviar_correo_cliente_pedido(nuevo_pedido, items_procesados)

    config_pago_actual = obtener_config_pago(db)

    if metodo_pago == "mercadopago":
        preference_data = {
            "items": [
                {
                    "title": f"Compra Online - Pedido #{nuevo_pedido.id} (Incluye Envío)",
                    "quantity": 1,
                    "unit_price": float(total_final),
                    "currency_id": "ARS"
                }
            ],
            "payer": {
                "name": cliente_nombre,
                "email": cliente_email,
                "identification": {"type": "DNI", "number": cliente_dni}
            },
            "external_reference": str(nuevo_pedido.id),
            "back_urls": {
                "success": f"http://localhost:8000/confirmacion/{nuevo_pedido.id}",
                "failure": f"http://localhost:8000/confirmacion/{nuevo_pedido.id}",
                "pending": f"http://localhost:8000/confirmacion/{nuevo_pedido.id}"
            }
        }
        try:
            preference_response = mp_sdk.preference().create(preference_data)
            response_body = preference_response.get("response", {})
            if "init_point" not in response_body:
                raise Exception(f"Respuesta inválida de Mercado Pago: {preference_response}")
            init_point = response_body["init_point"]
            nuevo_pedido.mp_preference_id = response_body.get("id")
            db.commit()
            return RedirectResponse(url=init_point, status_code=status.HTTP_303_SEE_OTHER)
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Error al conectar con Mercado Pago. Detalle: {str(e)}")

    return templates.TemplateResponse(request, "confirmacion.html", {
        "pedido": nuevo_pedido,
        "datos_transferencia": config_pago_actual,
        "metodo_pago": "transferencia"
    })

@app.get("/confirmacion/{pedido_id}", response_class=HTMLResponse)
def confirmacion_pedido(request: Request, pedido_id: int, db: Session = Depends(database.get_db)):
    pedido = db.query(models.Pedido).filter(models.Pedido.id == pedido_id).first()
    if not pedido:
        raise HTTPException(status_code=404, detail="Pedido no encontrado")
    config_pago = obtener_config_pago(db)
    return templates.TemplateResponse(request, "confirmacion.html", {
        "pedido": pedido,
        "datos_transferencia": config_pago,
        "metodo_pago": pedido.metodo_pago
    })

@app.get("/seguimiento", response_class=HTMLResponse)
def vista_seguimiento(request: Request):
    return templates.TemplateResponse(request, "seguimiento.html", {"pedido": None, "error": None})

@app.post("/seguimiento", response_class=HTMLResponse)
def buscar_seguimiento(
    request: Request,
    pedido_id: int = Form(...),
    identificador: str = Form(...),
    db: Session = Depends(database.get_db)
):
    identificador_clean = identificador.strip().lower()
    pedido = db.query(models.Pedido).filter(models.Pedido.id == pedido_id).first()
    if not pedido:
        return templates.TemplateResponse(request, "seguimiento.html", {"pedido": None, "error": f"No encontramos ningún pedido con el número #{pedido_id}."})
    
    dni_coincide = (pedido.cliente_dni or "").strip().lower() == identificador_clean
    email_coincide = (pedido.cliente_email or "").strip().lower() == identificador_clean
    if not (dni_coincide or email_coincide):
        return templates.TemplateResponse(request, "seguimiento.html", {"pedido": None, "error": "El DNI o email ingresado no coincide."})
    return templates.TemplateResponse(request, "seguimiento.html", {"pedido": pedido, "error": None})

@app.get("/admin/login", response_class=HTMLResponse)
def login_page(request: Request):
    return templates.TemplateResponse(request, "admin/login.html", {"error": None})

@app.post("/admin/login")
def procesar_login(request: Request, username: str = Form(...), password: str = Form(...)):
    if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
        response = RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)
        response.set_cookie(key=COOKIE_ADMIN_KEY, value="sesion_activa_ok", httponly=True)
        return response
    return templates.TemplateResponse(request, "admin/login.html", {"error": "Usuario o contraseña incorrectos"})

@app.get("/admin/logout")
def logout():
    response = RedirectResponse(url="/admin/login", status_code=status.HTTP_303_SEE_OTHER)
    response.delete_cookie(COOKIE_ADMIN_KEY)
    return response

@app.get("/admin", response_class=HTMLResponse)
def admin_dashboard(
    request: Request,
    q_busqueda: str = Query(None),
    q_fecha: str = Query(None),
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    query = db.query(models.Pedido).filter(models.Pedido.estado_pago != "Anulado")
    if q_busqueda and q_busqueda.strip():
        term = f"%{q_busqueda.strip()}%"
        if q_busqueda.strip().isdigit():
            query = query.filter((models.Pedido.id == int(q_busqueda.strip())) | (models.Pedido.cliente_nombre.ilike(term)))
        else:
            query = query.filter(models.Pedido.cliente_nombre.ilike(term))

    if q_fecha and q_fecha.strip():
        try:
            fecha_dt = datetime.strptime(q_fecha.strip(), "%Y-%m-%d")
            query = query.filter(
                models.Pedido.fecha >= fecha_dt,
                models.Pedido.fecha < fecha_dt + timedelta(days=1)
            )
        except ValueError:
            pass

    pedidos = query.order_by(models.Pedido.fecha.desc()).all()
    productos = db.query(models.Producto).all()
    posts = db.query(models.BlogPost).order_by(models.BlogPost.fecha.desc()).all()
    carrusel_items = db.query(models.CarruselItem).order_by(models.CarruselItem.orden.asc()).all()
    banners_admin = db.query(models.BannerPromocional).order_by(models.BannerPromocional.orden.asc()).all()
    skyscrapers = db.query(models.BannerSkyscraper).all()
    provincias_admin = db.query(models.ProvinciaEnvio).order_by(models.ProvinciaEnvio.provincia.asc()).all()
    cupones_admin = db.query(models.Cupon).order_by(models.Cupon.id.desc()).all()
    config_pago = db.query(models.ConfiguracionPago).first()
    
    eventos_admin = db.query(models.EventoReunion).order_by(models.EventoReunion.id.desc()).all()

    pedidos_con_wa = []
    for p in pedidos:
        num_tel = "".join(filter(str.isdigit, p.cliente_telefono or ""))
        mensaje = f"Hola {p.cliente_nombre}, ¡te escribimos desde Beauty Store! Tu pedido #{p.id} ya fue despachado."
        url_wa = f"https://wa.me/{num_tel}?text={urllib.parse.quote(mensaje)}" if num_tel else "#"
        pedidos_con_wa.append({"pedido": p, "url_wa": url_wa})

    return templates.TemplateResponse(
        "admin/dashboard.html",
        {
            "request": request,
            "pedidos_info": pedidos_con_wa,
            "productos": productos,
            "posts": posts,
            "carrusel_admin_items": carrusel_items,
            "banners_admin": banners_admin,
            "skyscrapers": skyscrapers,
            "provincias_admin": provincias_admin,
            "cupones_admin": cupones_admin,
            "config_pago": config_pago,
            "eventos_admin": eventos_admin,
            "q_busqueda": q_busqueda or "",
            "q_fecha": q_fecha or ""
        }
    )

# RUTAS ADMINISTRATIVAS DE EVENTOS, REUNIONES E INSCRIPTOS
@app.post("/admin/eventos/crear")
def admin_crear_evento(
    titulo: str = Form(...),
    descripcion: str = Form(...),
    temario: str = Form(""),
    oradora_nombre: str = Form(...),
    oradora_foto_url: str = Form(...),
    fecha_hora_texto: str = Form(...),
    modalidad: str = Form(...),
    lugar: str = Form(...),
    cupo_maximo: int = Form(20),
    activo: bool = Form(True),
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    if activo:
        db.query(models.EventoReunion).update({models.EventoReunion.activo: False})

    nuevo_evento = models.EventoReunion(
        titulo=titulo.strip(),
        descripcion=descripcion.strip(),
        temario=temario.strip() if temario else "",
        oradora_nombre=oradora_nombre.strip(),
        oradora_foto_url=oradora_foto_url.strip(),
        fecha_hora_texto=fecha_hora_texto.strip(),
        modalidad=modalidad.strip(),
        lugar=lugar.strip(),
        cupo_maximo=cupo_maximo,
        activo=activo
    )
    db.add(nuevo_evento)
    db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/eventos/{evento_id}/estado")
def admin_cambiar_estado_evento(
    evento_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    evento = db.query(models.EventoReunion).filter(models.EventoReunion.id == evento_id).first()
    if evento:
        if not evento.activo:
            db.query(models.EventoReunion).update({models.EventoReunion.activo: False})
        evento.activo = not evento.activo
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/eventos/{evento_id}/eliminar")
def admin_eliminar_evento(
    evento_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    evento = db.query(models.EventoReunion).filter(models.EventoReunion.id == evento_id).first()
    if evento:
        db.delete(evento)
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.get("/admin/eventos/{evento_id}/inscriptos", response_class=HTMLResponse)
def admin_ver_inscriptos(
    request: Request,
    evento_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    evento = db.query(models.EventoReunion).filter(models.EventoReunion.id == evento_id).first()
    if not evento:
        raise HTTPException(status_code=404, detail="Evento no encontrado")
    inscriptos = db.query(models.InscripcionEvento).filter(models.InscripcionEvento.evento_id == evento_id).all()
    return templates.TemplateResponse(
        "admin/eventos_inscriptos.html",
        {
            "request": request,
            "evento": evento,
            "inscriptos": inscriptos
        }
    )

@app.post("/admin/inscriptos/{inscripto_id}/editar")
def admin_editar_inscripto(
    inscripto_id: int,
    nombre_cliente: str = Form(...),
    telefono: str = Form(...),
    email: str = Form(None),
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    inscripto = db.query(models.InscripcionEvento).filter(models.InscripcionEvento.id == inscripto_id).first()
    if not inscripto:
        raise HTTPException(status_code=404, detail="Inscripción no encontrada")
    inscripto.nombre_cliente = nombre_cliente.strip()
    inscripto.telefono = telefono.strip()
    inscripto.email = email.strip() if email else None
    db.commit()
    return RedirectResponse(url=f"/admin/eventos/{inscripto.evento_id}/inscriptos", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/inscriptos/{inscripto_id}/eliminar")
def admin_eliminar_inscripto(
    inscripto_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    inscripto = db.query(models.InscripcionEvento).filter(models.InscripcionEvento.id == inscripto_id).first()
    if not inscripto:
        raise HTTPException(status_code=404, detail="Inscripción no encontrada")
    evento_id = inscripto.evento_id
    db.delete(inscripto)
    db.commit()
    return RedirectResponse(url=f"/admin/eventos/{evento_id}/inscriptos", status_code=status.HTTP_303_SEE_OTHER)

@app.get("/admin/eventos/{evento_id}/exportar/excel")
def admin_exportar_excel(
    evento_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    evento = db.query(models.EventoReunion).filter(models.EventoReunion.id == evento_id).first()
    if not evento:
        raise HTTPException(status_code=404, detail="Evento no encontrado")
    inscriptos = db.query(models.InscripcionEvento).filter(models.InscripcionEvento.evento_id == evento_id).all()

    wb = Workbook()
    ws = wb.active
    ws.title = "Inscriptos"

    ws.append(["#", "Nombre y Apellido", "Teléfono", "Email", "Fecha Registro"])

    for idx, item in enumerate(inscriptos, start=1):
        fecha_str = item.fecha_inscripcion.strftime("%d/%m/%Y %H:%M") if item.fecha_inscripcion else ""
        ws.append([idx, item.nombre_cliente, item.telefono, item.email or "", fecha_str])

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)

    filename = f"inscriptos_evento_{evento_id}.xlsx"
    return Response(
        content=stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/admin/eventos/{evento_id}/exportar/pdf")
def admin_exportar_pdf(
    evento_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    evento = db.query(models.EventoReunion).filter(models.EventoReunion.id == evento_id).first()
    if not evento:
        raise HTTPException(status_code=404, detail="Evento no encontrado")
    inscriptos = db.query(models.InscripcionEvento).filter(models.InscripcionEvento.evento_id == evento_id).all()

    buffer = io.BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    p.setFont("Helvetica-Bold", 16)
    p.drawString(50, height - 50, f"Lista de Inscriptos - {evento.titulo}")
    p.setFont("Helvetica", 10)
    p.drawString(50, height - 70, f"Total Asistentes: {len(inscriptos)}")

    y = height - 110
    p.setFont("Helvetica-Bold", 10)
    p.drawString(50, y, "#")
    p.drawString(80, y, "Nombre y Apellido")
    p.drawString(280, y, "Teléfono")
    p.drawString(420, y, "Email")
    
    p.line(50, y - 5, 560, y - 5)
    y -= 20

    p.setFont("Helvetica", 9)
    for idx, item in enumerate(inscriptos, start=1):
        if y < 50:
            p.showPage()
            y = height - 50
        p.drawString(50, y, str(idx))
        p.drawString(80, y, str(item.nombre_cliente)[:30])
        p.drawString(280, y, str(item.telefono))
        p.drawString(420, y, str(item.email or "—")[:25])
        y -= 18

    p.showPage()
    p.save()
    buffer.seek(0)

    return Response(
        content=buffer.getvalue(),
        media_type="application/pdf",
        headers={"Content-Disposition": f"inline; filename=inscriptos_evento_{evento_id}.pdf"}
    )

@app.post("/admin/transferencia/guardar")
def admin_guardar_transferencia(
    texto_publicidad: str = Form(""),
    transferencia_activo: bool = Form(False),
    transferencia_porcentaje: float = Form(...),
    titular: str = Form(...),
    cuit: str = Form(...),
    alias: str = Form(...),
    banco: str = Form(...),
    descuento_volumen_activo: bool = Form(False),  # <--- Agregado
    descuento_volumen_minimo: float = Form(0.0),   # <--- Agregado
    descuento_volumen_porcentaje: float = Form(0.0),# <--- Agregado
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    config = db.query(models.ConfiguracionPago).first()
    if not config:
        config = models.ConfiguracionPago()
        db.add(config)
    
    config.texto_publicidad = texto_publicidad.strip() if texto_publicidad else ""
    config.transferencia_activo = transferencia_activo
    config.transferencia_porcentaje = transferencia_porcentaje
    config.titular = titular
    config.cuit = cuit
    config.alias = alias
    config.banco = banco
    
    # <--- Guardar también los valores de volumen
    config.descuento_volumen_activo = descuento_volumen_activo
    config.descuento_volumen_minimo = descuento_volumen_minimo or 0.0
    config.descuento_volumen_porcentaje = descuento_volumen_porcentaje or 0.0
    
    db.commit()
    
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/envios/{envio_id}/editar")
def editar_config_envio(
    envio_id: int,
    costo: float = Form(...),
    gratis_desde: float = Form(None),
    activo: bool = Form(False),
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    config = db.query(models.ProvinciaEnvio).filter(models.ProvinciaEnvio.id == envio_id).first()
    if config:
        config.costo = costo
        config.gratis_desde = gratis_desde if gratis_desde and gratis_desde > 0 else None
        config.activo = activo
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/upload")
def subir_archivo_admin(
    file: UploadFile = File(...),
    auth: bool = Depends(verificar_admin)
):
    try:
        nombre_original = file.filename or "archivo"
        nfkd_form = unicodedata.normalize('NFKD', nombre_original)
        nombre_limpio = nfkd_form.encode('ASCII', 'ignore').decode('ASCII')
        nombre_limpio = re.sub(r'[^\w\.-]', '_', nombre_limpio)
        file_bytes = file.file.read()
        file_path = f"uploads/{nombre_limpio}"
        supabase.storage.from_("media").upload(
            path=file_path,
            file=file_bytes,
            file_options={"content-type": file.content_type, "upsert": "true"}
        )
        public_url = supabase.storage.from_("media").get_public_url(file_path)
        return {"url": public_url}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al subir archivo a Supabase: {str(e)}")

@app.post("/api/cupon/validar")
def validar_cupon(
    codigo: str = Form(...),
    subtotal: float = Form(...),
    cart_items: str = Form("[]"),
    db: Session = Depends(database.get_db)
):
    codigo_clean = codigo.replace(" ", "").strip().upper()
    cupon = db.query(models.Cupon).filter(models.Cupon.codigo == codigo_clean).first()

    if not cupon or not cupon.activo:
        raise HTTPException(status_code=400, detail="El código de cupón no es válido o ha expirado.")

    if cupon.fecha_expiracion and datetime.utcnow() - timedelta(hours=3) > cupon.fecha_expiracion:
        raise HTTPException(status_code=400, detail="Este cupón ya ha vencido.")

    if cupon.usos_maximos and cupon.usos_actuales >= cupon.usos_maximos:
        raise HTTPException(status_code=400, detail="Este cupón ha alcanzado el límite de usos permitidos.")

    if cupon.producto_id:
        try:
            items = json.loads(cart_items)
        except:
            items = []
        
        producto_encontrado = any(item.get("id") == cupon.producto_id for item in items)
        if not producto_encontrado:
            prod_obj = db.query(models.Producto).filter(models.Producto.id == cupon.producto_id).first()
            nombre_prod = prod_obj.nombre if prod_obj else "el producto participante"
            raise HTTPException(
                status_code=400, 
                detail=f"Este cupón es exclusivo y requiere que agregues '{nombre_prod}' a tu carrito."
            )

    if subtotal < cupon.monto_minimo:
        raise HTTPException(
            status_code=400, 
            detail=f"El monto mínimo de compra para este cupón es de ${cupon.monto_minimo:,.2f}"
        )

    if cupon.tipo_descuento == "porcentaje":
        monto_descuento = subtotal * (cupon.valor / 100.0)
    else:
        monto_descuento = min(cupon.valor, subtotal)

    return {
        "valido": True,
        "codigo": cupon.codigo,
        "tipo_descuento": cupon.tipo_descuento,
        "valor": cupon.valor,
        "monto_descuento": round(monto_descuento, 2),
        "mensaje": f"¡Cupón '{cupon.codigo}' aplicado con éxito!"
    }

@app.post("/admin/cupones/crear")
def crear_cupon(
    codigo: str = Form(...),
    tipo_descuento: str = Form(...),
    valor: float = Form(...),
    monto_minimo: float = Form(0.0),
    usos_maximos: int = Form(None),
    producto_id: int = Form(None),
    fecha_expiracion: str = Form(None),
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    codigo_clean = codigo.replace(" ", "").strip().upper()
    
    dt_expiracion = None
    if fecha_expiracion and fecha_expiracion.strip():
        try:
            dt_expiracion = datetime.strptime(fecha_expiracion.strip(), "%Y-%m-%d")
        except ValueError:
            pass

    nuevo_cupon = models.Cupon(
        codigo=codigo_clean,
        tipo_descuento=tipo_descuento,
        valor=valor,
        monto_minimo=monto_minimo or 0.0,
        usos_maximos=usos_maximos if usos_maximos and usos_maximos > 0 else None,
        producto_id=producto_id if producto_id and producto_id > 0 else None,
        fecha_expiracion=dt_expiracion,
        activo=True
    )
    db.add(nuevo_cupon)
    db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/cupones/{cupon_id}/estado")
def cambiar_estado_cupon(
    cupon_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    cupon = db.query(models.Cupon).filter(models.Cupon.id == cupon_id).first()
    if cupon:
        cupon.activo = not cupon.activo
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/cupones/{cupon_id}/eliminar")
def eliminar_cupon(
    cupon_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    cupon = db.query(models.Cupon).filter(models.Cupon.id == cupon_id).first()
    if cupon:
        db.delete(cupon)
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/skyscraper/guardar")
def guardar_skyscraper(
    titulo: str = Form(""),
    texto_vertical: str = Form(""),
    media_url: str = Form(...),
    enlace_url: str = Form(""),
    activo: bool = Form(False),
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    if activo:
        db.query(models.BannerSkyscraper).update({models.BannerSkyscraper.activo: False})
    nuevo_sk = models.BannerSkyscraper(
        titulo=titulo,
        texto_vertical=texto_vertical,
        media_url=media_url,
        enlace_url=enlace_url,
        activo=activo
    )
    db.add(nuevo_sk)
    db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/skyscraper/{sk_id}/estado")
def cambiar_estado_skyscraper(
    sk_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    sk = db.query(models.BannerSkyscraper).filter(models.BannerSkyscraper.id == sk_id).first()
    if sk:
        if not sk.activo:
            db.query(models.BannerSkyscraper).update({models.BannerSkyscraper.activo: False})
        sk.activo = not sk.activo
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/skyscraper/{sk_id}/eliminar")
def eliminar_skyscraper(
    sk_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    sk = db.query(models.BannerSkyscraper).filter(models.BannerSkyscraper.id == sk_id).first()
    if sk:
        db.delete(sk)
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/pedidos/{pedido_id}/despachar")
def despachar_pedido(
    pedido_id: int,
    transporte_nombre: str = Form(...),
    guia_seguimiento: str = Form(...),
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    pedido = db.query(models.Pedido).filter(models.Pedido.id == pedido_id).first()
    if pedido:
        pedido.transporte_nombre = transporte_nombre
        pedido.guia_seguimiento = guia_seguimiento
        pedido.estado_pago = "Aprobado"
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.get("/pedido/{pedido_id}/recibo-pdf")
def descargar_recibo_pdf(pedido_id: int, db: Session = Depends(database.get_db)):
    pedido = db.query(models.Pedido).filter(models.Pedido.id == pedido_id).first()
    if not pedido:
        raise HTTPException(status_code=404, detail="Pedido no encontrado")
    
    items = []
    for d in pedido.detalles:
        items.append({
            "producto": d.producto,
            "cantidad": d.cantidad,
            "precio_unitario": d.precio_unitario
        })

    pdf_bytes = generar_recibo_pdf(pedido, items, DATOS_VENDEDOR)
    
    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={
            "Content-Disposition": f"attachment; filename=Recibo_BeautyStore_Pedido_{pedido.id}.pdf"
        }
    )

@app.get("/admin/pedidos/{pedido_id}/etiqueta", response_class=HTMLResponse)
def ver_etiqueta_envio(
    request: Request,
    pedido_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    pedido = db.query(models.Pedido).filter(models.Pedido.id == pedido_id).first()
    if not pedido:
        raise HTTPException(status_code=404, detail="Pedido no encontrado")
    return templates.TemplateResponse(request, "admin/etiqueta.html", {
        "pedido": pedido,
        "vendedor": DATOS_VENDEDOR
    })

@app.post("/admin/pedidos/{pedido_id}/editar")
def admin_editar_pedido(
    pedido_id: int,
    cliente_nombre: str = Form(...),
    cliente_telefono: str = Form(...),
    cliente_email: str = Form(None),
    cliente_dni: str = Form(...),
    cliente_direccion: str = Form(...),
    cliente_localidad: str = Form(...),
    cliente_provincia: str = Form(...),
    cliente_codigo_postal: str = Form(...),
    costo_envio: float = Form(...),
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    pedido = db.query(models.Pedido).filter(models.Pedido.id == pedido_id).first()
    if not pedido:
        raise HTTPException(status_code=404, detail="Pedido no encontrado")
    pedido.cliente_nombre = cliente_nombre
    pedido.cliente_telefono = cliente_telefono
    pedido.cliente_email = cliente_email
    pedido.cliente_dni = cliente_dni
    pedido.cliente_direccion = cliente_direccion
    pedido.cliente_localidad = cliente_localidad
    pedido.cliente_provincia = cliente_provincia
    pedido.cliente_codigo_postal = cliente_codigo_postal
    pedido.costo_envio = costo_envio
    db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/pedidos/{pedido_id}/anular")
def admin_anular_pedido(
    pedido_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    pedido = db.query(models.Pedido).filter(models.Pedido.id == pedido_id).first()
    if not pedido:
        raise HTTPException(status_code=404, detail="Pedido no encontrado")
    if pedido.estado_pago != "Anulado":
        for detalle in pedido.detalles:
            producto = db.query(models.Producto).filter(models.Producto.id == detalle.producto_id).first()
            if producto:
                producto.stock += detalle.cantidad
        pedido.estado_pago = "Anulado"
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/pedidos/{pedido_id}/eliminar")
def admin_eliminar_pedido(
    pedido_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    pedido = db.query(models.Pedido).filter(models.Pedido.id == pedido_id).first()
    if not pedido:
        raise HTTPException(status_code=404, detail="Pedido no encontrado")
    for detalle in pedido.detalles:
        db.delete(detalle)
    db.delete(pedido)
    db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/productos/crear")
def crear_producto(
    nombre: str = Form(...),
    descripcion: str = Form(...),
    precio: str = Form(...),  # <--- Cambiado a str para parsear de forma segura
    descuento_porcentaje: float = Form(0.0),
    stock: int = Form(...),
    imagen_url: str = Form(...),
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    precio_limpio = limpiar_precio(precio)
    nuevo_prod = models.Producto(
        nombre=nombre, 
        descripcion=descripcion, 
        precio=precio_limpio, 
        descuento_porcentaje=descuento_porcentaje or 0.0,
        stock=stock, 
        imagen_url=imagen_url, 
        activo=True
    )
    db.add(nuevo_prod)
    db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/productos/{producto_id}/editar")
def editar_producto(
    producto_id: int,
    nombre: str = Form(...),
    descripcion: str = Form(...),
    precio: str = Form(...),  # <--- Cambiado a str
    descuento_porcentaje: float = Form(0.0),
    stock: int = Form(...),
    imagen_url: str = Form(...),
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    producto = db.query(models.Producto).filter(models.Producto.id == producto_id).first()
    if not producto:
        raise HTTPException(status_code=404, detail="Producto no encontrado")
    producto.nombre = nombre
    producto.descripcion = descripcion
    producto.precio = limpiar_precio(precio)  # <--- Limpieza aplicada
    producto.descuento_porcentaje = descuento_porcentaje or 0.0
    producto.stock = stock
    producto.imagen_url = imagen_url
    db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/productos/{producto_id}/estado")
def cambiar_estado_producto(producto_id: int, db: Session = Depends(database.get_db), auth: bool = Depends(verificar_admin)):
    producto = db.query(models.Producto).filter(models.Producto.id == producto_id).first()
    if producto:
        producto.activo = not producto.activo
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/productos/{producto_id}/eliminar")
def eliminar_producto(
    producto_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    producto = db.query(models.Producto).filter(models.Producto.id == producto_id).first()
    if producto:
        try:
            db.delete(producto)
            db.commit()
        except IntegrityError:
            db.rollback()
            # Si tiene pedidos asociados, lo desactivamos automáticamente en lugar de dar error 500
            producto.activo = False
            db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/blog/crear")
def crear_post(
    titulo: str = Form(...),
    subtitulo: str = Form(...),
    contenido: str = Form(...),
    imagen_url: str = Form(...),
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    nuevo_post = models.BlogPost(titulo=titulo, subtitulo=subtitulo, contenido=contenido, imagen_url=imagen_url)
    db.add(nuevo_post)
    db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/blog/{post_id}/editar")
def editar_post(
    post_id: int,
    titulo: str = Form(...),
    subtitulo: str = Form(...),
    contenido: str = Form(...),
    imagen_url: str = Form(...),
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    post = db.query(models.BlogPost).filter(models.BlogPost.id == post_id).first()
    if not post:
        raise HTTPException(status_code=404, detail="Artículo no encontrado")
    post.titulo = titulo
    post.subtitulo = subtitulo
    post.contenido = contenido
    post.imagen_url = imagen_url
    db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/blog/{post_id}/eliminar")
def eliminar_post(
    post_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    post = db.query(models.BlogPost).filter(models.BlogPost.id == post_id).first()
    if post:
        db.delete(post)
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/carrusel/crear")
def crear_carrusel_item(
    titulo: str = Form(""),
    imagen_url: str = Form(...),
    orden: int = Form(0),
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    nuevo_item = models.CarruselItem(titulo=titulo, imagen_url=imagen_url, orden=orden, activo=True)
    db.add(nuevo_item)
    db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/carrusel/{item_id}/editar")
def editar_carrusel_item(
    item_id: int,
    titulo: str = Form(""),
    imagen_url: str = Form(...),
    orden: int = Form(0),
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    item = db.query(models.CarruselItem).filter(models.CarruselItem.id == item_id).first()
    if item:
        item.titulo = titulo
        item.imagen_url = imagen_url
        item.orden = orden
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/carrusel/{item_id}/eliminar")
def eliminar_carrusel_item(
    item_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    item = db.query(models.CarruselItem).filter(models.CarruselItem.id == item_id).first()
    if item:
        db.delete(item)
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/banners/crear")
def crear_banner_promocional(
    titulo: str = Form(""),
    media_url: str = Form(...),
    tipo_media: str = Form("imagen"),
    enlace_url: str = Form(""),
    orden: int = Form(0),
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    nuevo_banner = models.BannerPromocional(
        titulo=titulo,
        media_url=media_url,
        tipo_media=tipo_media,
        enlace_url=enlace_url,
        orden=orden,
        activo=True
    )
    db.add(nuevo_banner)
    db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/banners/{banner_id}/editar")
def editar_banner_promocional(
    banner_id: int,
    titulo: str = Form(""),
    media_url: str = Form(...),
    tipo_media: str = Form("imagen"),
    enlace_url: str = Form(""),
    orden: int = Form(0),
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    banner = db.query(models.BannerPromocional).filter(models.BannerPromocional.id == banner_id).first()
    if banner:
        banner.titulo = titulo
        banner.media_url = media_url
        banner.tipo_media = tipo_media
        banner.enlace_url = enlace_url
        banner.orden = orden
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/banners/{banner_id}/eliminar")
def eliminar_banner_promocional(
    banner_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    banner = db.query(models.BannerPromocional).filter(models.BannerPromocional.id == banner_id).first()
    if banner:
        db.delete(banner)
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/admin/banners/{banner_id}/estado")
def cambiar_estado_banner(
    banner_id: int,
    db: Session = Depends(database.get_db),
    auth: bool = Depends(verificar_admin)
):
    banner = db.query(models.BannerPromocional).filter(models.BannerPromocional.id == banner_id).first()
    if banner:
        banner.activo = not banner.activo
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@app.post("/webhook/mercadopago")
async def mercadopago_webhook(request: Request, db: Session = Depends(database.get_db)):
    try:
        body = await request.json()
        tipo_evento = body.get("type") or body.get("topic")
        if tipo_evento == "payment":
            data_info = body.get("data", {})
            payment_id = data_info.get("id")
            if payment_id:
                payment_response = mp_sdk.payment().get(payment_id)
                payment_info = payment_response.get("response", {})
                status_pago = payment_info.get("status")
                pedido_id_str = payment_info.get("external_reference")
                if pedido_id_str:
                    pedido = db.query(models.Pedido).filter(models.Pedido.id == int(pedido_id_str)).first()
                    if pedido and pedido.estado_pago != "Aprobado":
                        if status_pago == "approved":
                            pedido.estado_pago = "Aprobado"
                            db.commit()
                        elif status_pago in ["rejected", "cancelled"]:
                            pedido.estado_pago = "Rechazado"
                            db.commit()
    except Exception as e:
        print(f"Error webhook Mercado Pago: {e}")
    return {"status": "ok"}