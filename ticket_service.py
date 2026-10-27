import logging
import random
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from excel_handler import append_result, find_first_recorded_ticket
from ticket_parser import TicketParseError, parse_tickets


logger = logging.getLogger(__name__)


class TicketAssignmentError(ValueError):
    """Raised when an existing assignment cannot be honored."""


class TicketSaveError(RuntimeError):
    """Raised when an assignment could not be saved to the results journal."""


@dataclass(frozen=True)
class TicketAssignment:
    group: str
    last_name: str
    first_name: str
    ticket_number: int
    questions: tuple[str, str, str]
    is_repeat: bool
    requested_at: datetime


class TicketService:
    """Assign tickets and persist each request in the results journal."""

    def __init__(
        self, tickets_path: str | Path, results_path: str | Path
    ) -> None:
        self.tickets_path = Path(tickets_path)
        self.results_path = Path(results_path)
        self._pending_assignments: dict[
            tuple[str, str, str], TicketAssignment
        ] = {}

    def generate_ticket(
        self, group: str, last_name: str, first_name: str
    ) -> TicketAssignment:
        """Create and save an assignment, or retain it if saving fails."""
        student = (group, last_name, first_name)
        pending = self._pending_assignments.get(student)
        if pending is not None:
            return self._save_assignment(student, pending)

        tickets = parse_tickets(self.tickets_path)
        original_number = find_first_recorded_ticket(
            self.results_path, group, last_name, first_name
        )

        if original_number is None:
            ticket_number = random.choice(list(tickets))
            is_repeat = False
        else:
            ticket_number = self._normalize_ticket_number(original_number)
            if ticket_number not in tickets:
                raise TicketAssignmentError(
                    f"Первоначальный билет студента — {original_number}. "
                    "Он отсутствует в текущем документе. "
                    "Назначить другой билет нельзя."
                )
            is_repeat = True

        assignment = TicketAssignment(
            group=group,
            last_name=last_name,
            first_name=first_name,
            ticket_number=ticket_number,
            questions=tuple(tickets[ticket_number]),
            is_repeat=is_repeat,
            requested_at=datetime.now(),
        )
        return self._save_assignment(student, assignment)

    def retry_save(
        self, group: str, last_name: str, first_name: str
    ) -> TicketAssignment:
        """Retry saving the same in-memory assignment after a save failure."""
        student = (group, last_name, first_name)
        assignment = self._pending_assignments.get(student)
        if assignment is None:
            raise TicketAssignmentError(
                "Нет несохранённого назначения для этого студента."
            )
        return self._save_assignment(student, assignment)

    @staticmethod
    def _normalize_ticket_number(ticket_number: int | str) -> int:
        try:
            return int(ticket_number)
        except (TypeError, ValueError) as error:
            raise TicketAssignmentError(
                f"В журнале указан некорректный номер билета: {ticket_number}."
            ) from error

    def _save_assignment(
        self,
        student: tuple[str, str, str],
        assignment: TicketAssignment,
    ) -> TicketAssignment:
        self._pending_assignments[student] = assignment
        try:
            append_result(
                self.results_path,
                assignment.group,
                assignment.last_name,
                assignment.first_name,
                assignment.ticket_number,
                requested_at=assignment.requested_at,
            )
        except PermissionError as error:
            logger.error(
                "Не удалось сохранить назначение билета %s для %s: %s",
                assignment.ticket_number,
                student,
                error,
            )
            raise TicketSaveError(
                f"Билет {assignment.ticket_number} назначен, но не сохранён. "
                "Закройте файл журнала в Excel и повторите сохранение."
            ) from error

        del self._pending_assignments[student]
        return assignment
