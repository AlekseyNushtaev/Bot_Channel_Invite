from datetime import datetime
from db.models import SubscriptionRequest, SessionLocal


# Добавление новой заявки
def add_subscription_request(
    user_id: int,
    username: str | None,
    first_name: str | None,
    last_name: str | None,
    channel_id: int,
    channel_name: str | None,
    channel_link: str | None,
):
    db = SessionLocal()
    try:
        request = SubscriptionRequest(
            user_id=user_id,
            username=username,
            first_name=first_name,
            last_name=last_name,
            channel_id=channel_id,
            channel_name=channel_name,
            channel_link=channel_link,
        )
        db.add(request)
        db.commit()
    finally:
        db.close()


# Подтверждение заявки (устанавливаем time_confirm)
def confirm_subscription(user_id: int, channel_id: int):
    db = SessionLocal()
    try:
        request = db.query(SubscriptionRequest).filter(
            SubscriptionRequest.user_id == user_id,
            SubscriptionRequest.channel_id == channel_id,
            SubscriptionRequest.time_confirm.is_(None),
        ).first()

        if request:
            request.time_confirm = datetime.now()
            request.user_is_block = False
            db.commit()
    finally:
        db.close()


# Проверка наличия активной заявки
def has_pending_request(user_id: int, channel_id: int) -> bool:
    db = SessionLocal()
    try:
        return db.query(SubscriptionRequest).filter(
            SubscriptionRequest.user_id == user_id,
            SubscriptionRequest.channel_id == channel_id,
            SubscriptionRequest.time_confirm.is_(None),
        ).first() is not None
    finally:
        db.close()


# Обновление статуса пользователя на "заблокирован"
def update_user_blocked(user_id: int):
    db = SessionLocal()
    try:
        requests = db.query(SubscriptionRequest).filter(
            SubscriptionRequest.user_id == user_id,
        ).all()

        for request in requests:
            request.user_is_block = True
        db.commit()
    finally:
        db.close()


# Обновление статуса пользователя на "разблокирован"
def update_user_unblocked(user_id: int):
    db = SessionLocal()
    try:
        requests = db.query(SubscriptionRequest).filter(
            SubscriptionRequest.user_id == user_id,
        ).all()

        for request in requests:
            request.user_is_block = False
        db.commit()
    finally:
        db.close()


def get_all_users_unblock() -> set[int]:
    db = SessionLocal()
    try:
        # Получаем только уникальные user_id, где user_is_block == False
        users = db.query(SubscriptionRequest.user_id).filter(
            SubscriptionRequest.user_is_block == False
        ).distinct().all()

        # Преобразуем результат в множество чисел
        return {user[0] for user in users}
    finally:
        db.close()
