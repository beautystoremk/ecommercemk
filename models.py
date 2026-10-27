from sqlalchemy import Column, Integer, String, Float, DateTime, ForeignKey, Text, Boolean
from sqlalchemy.orm import relationship
from datetime import datetime
from database import Base

class Producto(Base):
    __tablename__ = "productos"

    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String, index=True, nullable=False)
    descripcion = Column(Text)
    precio = Column(Float, nullable=False)
    descuento_porcentaje = Column(Float, default=0.0)
    stock = Column(Integer, default=0)
    imagen_url = Column(String, default="/static/placeholder.jpg")
    activo = Column(Boolean, default=True)
    resenas = relationship("Resena", back_populates="producto", cascade="all, delete-orphan")

class BlogPost(Base):
    __tablename__ = "blog_posts"

    id = Column(Integer, primary_key=True, index=True)
    titulo = Column(String, nullable=False)
    subtitulo = Column(String)
    contenido = Column(Text, nullable=False)
    imagen_url = Column(String)
    fecha = Column(DateTime, default=datetime.utcnow)
    publicado = Column(Boolean, default=True)

class Pedido(Base):
    __tablename__ = "pedidos"
    
    id = Column(Integer, primary_key=True, index=True)
    cliente_nombre = Column(String, nullable=False)
    cliente_email = Column(String, nullable=False)
    cliente_telefono = Column(String, nullable=False)
    cliente_dni = Column(String, nullable=False)
    cliente_direccion = Column(String, nullable=False)
    cliente_localidad = Column(String, nullable=False)
    cliente_provincia = Column(String, nullable=True)
    cliente_codigo_postal = Column(String, nullable=True)
    metodo_pago = Column(String, nullable=False)
    estado_pago = Column(String, default="Pendiente")
    total = Column(Float, nullable=False)
    costo_envio = Column(Float, default=0.0)
    transporte_nombre = Column(String, nullable=True)
    guia_seguimiento = Column(String, nullable=True)
    mp_preference_id = Column(String, nullable=True)
    fecha = Column(DateTime, default=datetime.utcnow)
    
    detalles = relationship("PedidoDetalle", back_populates="pedido", cascade="all, delete-orphan")

class PedidoDetalle(Base):
    __tablename__ = "pedido_detalles"

    id = Column(Integer, primary_key=True, index=True)
    pedido_id = Column(Integer, ForeignKey("pedidos.id"))
    producto_id = Column(Integer, ForeignKey("productos.id"))
    cantidad = Column(Integer, nullable=False)
    precio_unitario = Column(Float, nullable=False)

    pedido = relationship("Pedido", back_populates="detalles")
    producto = relationship("Producto")

class CarruselItem(Base):
    __tablename__ = "carrusel_items"

    id = Column(Integer, primary_key=True, index=True)
    titulo = Column(String, nullable=True)
    imagen_url = Column(String, nullable=False)
    orden = Column(Integer, default=0)
    activo = Column(Boolean, default=True)

class BannerPromocional(Base):
    __tablename__ = "banners_promocionales"

    id = Column(Integer, primary_key=True, index=True)
    titulo = Column(String, nullable=True)
    media_url = Column(String, nullable=False)
    tipo_media = Column(String, default="imagen")
    enlace_url = Column(String, nullable=True)
    activo = Column(Boolean, default=True)
    orden = Column(Integer, default=0)

class BannerSkyscraper(Base):
    __tablename__ = "banners_skyscraper"

    id = Column(Integer, primary_key=True, index=True)
    titulo = Column(String, nullable=True)          
    texto_vertical = Column(String, nullable=True)  
    media_url = Column(String, nullable=False)       
    enlace_url = Column(String, nullable=True)      
    activo = Column(Boolean, default=True)           
    fecha_creacion = Column(DateTime, default=datetime.utcnow)

class ProvinciaEnvio(Base):
    __tablename__ = "provincias_envio"
    
    id = Column(Integer, primary_key=True, index=True)
    provincia = Column(String(100), unique=True, index=True)
    costo = Column(Float, default=0.0)
    gratis_desde = Column(Float, nullable=True)
    activo = Column(Boolean, default=True)

class Resena(Base):
    __tablename__ = "resenas"

    id = Column(Integer, primary_key=True, index=True)
    producto_id = Column(Integer, ForeignKey("productos.id"), nullable=False)
    nombre_cliente = Column(String(100), nullable=False)
    calificacion = Column(Integer, nullable=False)
    comentario = Column(Text, nullable=False)
    aprobado = Column(Boolean, default=True)
    fecha = Column(DateTime, default=datetime.utcnow)

    producto = relationship("Producto", back_populates="resenas")

class Cupon(Base):
    __tablename__ = "cupones"

    id = Column(Integer, primary_key=True, index=True)
    codigo = Column(String(50), unique=True, index=True, nullable=False)
    tipo_descuento = Column(String(20), default="porcentaje")
    valor = Column(Float, nullable=False)
    monto_minimo = Column(Float, default=0.0)
    usos_maximos = Column(Integer, nullable=True)
    usos_actuales = Column(Integer, default=0)
    producto_id = Column(Integer, ForeignKey("productos.id"), nullable=True)
    fecha_expiracion = Column(DateTime, nullable=True)
    activo = Column(Boolean, default=True)

    producto = relationship("Producto")

class ConfiguracionPago(Base):
    __tablename__ = "configuracion_pago"

    id = Column(Integer, primary_key=True, index=True)
    transferencia_activo = Column(Boolean, default=True)
    transferencia_porcentaje = Column(Float, default=10.0)
    titular = Column(String(100), default="Dominguez Analia Alejandra.")
    cuit = Column(String(50), default="27-23893486-4")
    alias = Column(String(100), default="ana.alejandra.1974")
    banco = Column(String(100), default="Mercado Pago")
    texto_publicidad = Column(Text, default="")
    descuento_volumen_activo = Column(Boolean, default=False)
    descuento_volumen_minimo = Column(Float, default=0.0)
    descuento_volumen_porcentaje = Column(Float, default=0.0)

class EventoReunion(Base):
    __tablename__ = "eventos_reuniones"

    id = Column(Integer, primary_key=True, index=True)
    titulo = Column(String, nullable=False)
    descripcion = Column(Text, nullable=False)
    temario = Column(Text, nullable=True) # Ítems o lista de temas
    oradora_nombre = Column(String, nullable=False)
    oradora_foto_url = Column(String, default="/static/placeholder.jpg")
    fecha_hora_texto = Column(String, nullable=False) # Ej: "Sábado 15 de Noviembre - 17:00 hs"
    modalidad = Column(String, default="Presencial") # Presencial u Online (Zoom / Meet)
    lugar = Column(String, nullable=False) # Ej: "Luján de Cuyo, Mendoza" o "Vía Zoom"
    cupo_maximo = Column(Integer, default=20)
    activo = Column(Boolean, default=True)
    fecha_creacion = Column(DateTime, default=datetime.utcnow)

    inscripciones = relationship("InscripcionEvento", back_populates="evento", cascade="all, delete-orphan")

class InscripcionEvento(Base):
    __tablename__ = "inscripciones_eventos"

    id = Column(Integer, primary_key=True, index=True)
    evento_id = Column(Integer, ForeignKey("eventos_reuniones.id"), nullable=False)
    nombre_cliente = Column(String, nullable=False)
    telefono = Column(String, nullable=False)
    email = Column(String, nullable=True)
    fecha_inscripcion = Column(DateTime, default=datetime.utcnow)

    evento = relationship("EventoReunion", back_populates="inscripciones")