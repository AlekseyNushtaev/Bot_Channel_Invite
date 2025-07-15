from datetime import datetime
from sqlalchemy import create_engine, Column, Integer, String, DateTime, Boolean
from sqlalchemy.orm import declarative_base, sessionmaker

# Настройка SQLAlchemy
SQLALCHEMY_DATABASE_URL = "sqlite:///db/database.db"
engine = create_engine(SQLALCHEMY_DATABASE_URL, echo=True)  # `echo=True` для логов SQL-запросов

Base = declarative_base()
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)


# Модель для хранения заявок на подписку
class SubscriptionRequest(Base):
    __tablename__ = "subscription_requests"

    id = Column(Integer, primary_key=True, autoincrement=True)
    user_id = Column(Integer, nullable=False)
    username = Column(String)
    first_name = Column(String)
    last_name = Column(String)
    channel_id = Column(Integer, nullable=False)
    channel_name = Column(String)
    channel_link = Column(String)
    time_request = Column(DateTime, default=datetime.now)
    time_confirm = Column(DateTime, nullable=True)
    user_is_block = Column(Boolean, default=True)


# Создаем таблицы (вызывается при старте бота)
def create_tables():
    Base.metadata.create_all(bind=engine)