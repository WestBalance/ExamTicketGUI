import logging
import re
import zipfile
from pathlib import Path
from xml.etree.ElementTree import ParseError

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from lxml.etree import XMLSyntaxError


logger = logging.getLogger(__name__)
TICKET_HEADING = re.compile(r"^\s*Билет\s+(\d+)\s*$")
HEADING_CANDIDATE = re.compile(r"^\s*Билет\b", re.IGNORECASE)
QUESTION_PARAGRAPH = re.compile(r"^\s*\d+[.)]\s*(.+?)\s*$")


class TicketParseError(ValueError):
    """Raised when a ticket document cannot provide any valid tickets."""


def parse_tickets(file_path: str | Path) -> dict[int, list[str]]:
    """Parse ticket headings and their three numbered questions from a DOCX."""
    path = Path(file_path)
    try:
        document = Document(path)
    except (
        OSError,
        ValueError,
        KeyError,
        zipfile.BadZipFile,
        PackageNotFoundError,
        XMLSyntaxError,
        ParseError,
    ) as error:
        logger.error("Не удалось прочитать документ с билетами %s: %s", path, error)
        raise TicketParseError(
            f"Не удалось прочитать документ с билетами: {path}"
        ) from error

    tickets: dict[int, list[str]] = {}
    current_number: int | None = None
    current_questions: list[str] = []

    def finish_ticket() -> None:
        if current_number is None:
            return
        if len(current_questions) != 3:
            logger.warning(
                "Билет %s пропущен: найдено вопросов — %s, ожидалось 3",
                current_number,
                len(current_questions),
            )
        elif current_number in tickets:
            logger.warning(
                "Повторный заголовок билета %s пропущен", current_number
            )
        else:
            tickets[current_number] = list(current_questions)

    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        heading_match = TICKET_HEADING.fullmatch(text)
        if heading_match:
            finish_ticket()
            current_number = int(heading_match.group(1))
            current_questions = []
            continue

        if HEADING_CANDIDATE.match(text):
            finish_ticket()
            current_number = None
            current_questions = []
            logger.warning("Некорректный заголовок билета пропущен: %s", text)
            continue

        if current_number is not None:
            question_match = QUESTION_PARAGRAPH.fullmatch(text)
            if question_match:
                current_questions.append(question_match.group(1).strip())

    finish_ticket()

    if not tickets:
        logger.error("В документе %s нет корректных экзаменационных билетов", path)
        raise TicketParseError(
            f"В документе {path} не найдено ни одного корректного билета."
        )
    return tickets
