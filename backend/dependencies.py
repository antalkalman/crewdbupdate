from backend.services.title_service import TitleService


def get_title_service() -> TitleService:
    return TitleService()
