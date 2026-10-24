from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

# URL de tu base de datos en Supabase PostgreSQL
SQLALCHEMY_DATABASE_URL = "postgresql://postgres.ennfuhqoqxroelcfklnk:beautystoremk2026@aws-1-us-west-2.pooler.supabase.com:6543/postgres"

# CORREGIDO: Se incluyen los connect_args directamente dentro de create_engine
engine = create_engine(
    SQLALCHEMY_DATABASE_URL, 
    connect_args={"prepare_threshold": None}
)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()