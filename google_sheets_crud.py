"""
Клиент Google Sheets (API v4) через сервисный аккаунт.
Импортируйте GoogleSheetsClient в своё приложение и вызывайте методы для работы с таблицей.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ("https://www.googleapis.com/auth/spreadsheets",)


def _default_credentials_path() -> Path:
    base = Path(__file__).resolve().parent
    env = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if env:
        return Path(env)
    return base / "excel-factory-492507-dacf37ad1796.json"


class GoogleSheetsClient:
    """
    CRUD-операции над одной таблицей (spreadsheet).

    - Create: append_rows — добавить строки в конец диапазона.
    - Read: read_range / read_all_used — чтение значений.
    - Update: update_range — перезапись диапазона.
    - Delete: clear_range — очистка ячеек (данные удаляются).
    """

    def __init__(
        self,
        spreadsheet_id: str,
        credentials_path: Optional[str | Path] = None,
        *,
        default_sheet_name: str = "Sheet1",
    ) -> None:
        self.spreadsheet_id = spreadsheet_id
        self.default_sheet_name = default_sheet_name
        path = Path(credentials_path) if credentials_path else _default_credentials_path()
        if not path.is_file():
            raise FileNotFoundError(f"Файл ключа сервисного аккаунта не найден: {path}")

        creds = Credentials.from_service_account_file(str(path), scopes=SCOPES)
        self._service = build("sheets", "v4", credentials=creds, cache_discovery=False)

    @property
    def values(self) -> Any:
        return self._service.spreadsheets().values()

    def list_sheet_names(self) -> List[str]:
        """Имена листов таблицы (как в UI, например «Лист1»)."""
        try:
            meta = self._service.spreadsheets().get(
                spreadsheetId=self.spreadsheet_id
            ).execute()
        except HttpError as e:
            raise RuntimeError(f"Не удалось прочитать метаданные таблицы: {e}") from e
        return [s["properties"]["title"] for s in meta.get("sheets", [])]

    def read_range(self, range_a1: str) -> List[List[Any]]:
        """Прочитать диапазон в A1-нотации (например ``Sheet1!A1:C10`` или ``Sheet1!A:ZZ``)."""
        try:
            result = (
                self.values.get(spreadsheetId=self.spreadsheet_id, range=range_a1).execute()
            )
        except HttpError as e:
            raise RuntimeError(f"Ошибка чтения диапазона {range_a1!r}: {e}") from e
        return result.get("values", [])

    def read_all_used(self, sheet_name: Optional[str] = None) -> List[List[Any]]:
        """
        Прочитать ячейки листа в сетке A1:ZZ10000 (достаточно для типичных таблиц).
        """
        name = sheet_name or self.default_sheet_name
        # Открытый диапазон вида A:ZZ API не принимает; ограничиваем сеткой A1:ZZ10000.
        return self.read_range(f"{name}!A1:ZZ10000")

    def append_rows(
        self,
        values: List[List[Any]],
        range_a1: Optional[str] = None,
        *,
        value_input_option: str = "USER_ENTERED",
        insert_data_option: str = "INSERT_ROWS",
    ) -> dict[str, Any]:
        """
        Добавить строки в конец таблицы (Create).

        ``range_a1`` — например ``Sheet1!A1`` (лист и первая колонка); если None, используется
        ``{default_sheet_name}!A1``.
        """
        if not values:
            return {}
        rng = range_a1 or f"{self.default_sheet_name}!A1"
        try:
            return (
                self.values.append(
                    spreadsheetId=self.spreadsheet_id,
                    range=rng,
                    valueInputOption=value_input_option,
                    insertDataOption=insert_data_option,
                    body={"values": values},
                ).execute()
            )
        except HttpError as e:
            raise RuntimeError(f"Ошибка append в {rng!r}: {e}") from e

    def update_range(
        self,
        range_a1: str,
        values: List[List[Any]],
        *,
        value_input_option: str = "USER_ENTERED",
    ) -> dict[str, Any]:
        """Перезаписать диапазон (Update)."""
        try:
            return (
                self.values.update(
                    spreadsheetId=self.spreadsheet_id,
                    range=range_a1,
                    valueInputOption=value_input_option,
                    body={"values": values},
                ).execute()
            )
        except HttpError as e:
            raise RuntimeError(f"Ошибка update {range_a1!r}: {e}") from e

    def clear_range(self, range_a1: str) -> dict[str, Any]:
        """Очистить диапазон (Delete данных в ячейках)."""
        try:
            return (
                self.values.clear(spreadsheetId=self.spreadsheet_id, range=range_a1).execute()
            )
        except HttpError as e:
            raise RuntimeError(f"Ошибка clear {range_a1!r}: {e}") from e

    def batch_get(self, ranges: List[str]) -> List[List[List[Any]]]:
        """Несколько диапазонов за один запрос."""
        try:
            result = (
                self.values.batchGet(
                    spreadsheetId=self.spreadsheet_id, ranges=ranges
                ).execute()
            )
        except HttpError as e:
            raise RuntimeError(f"Ошибка batchGet: {e}") from e
        value_ranges = result.get("valueRanges", [])
        out: List[List[List[Any]]] = []
        for vr in value_ranges:
            out.append(vr.get("values", []))
        return out

    def batch_update(self, requests: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Пакетное изменение таблицы (форматирование, объединение ячеек, новые листы и т.д.)."""
        if not requests:
            return {}
        try:
            return (
                self._service.spreadsheets()
                .batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body={"requests": requests},
                )
                .execute()
            )
        except HttpError as e:
            raise RuntimeError(f"Ошибка batchUpdate: {e}") from e

    def get_sheet_id(self, sheet_title: str) -> int:
        """Числовой sheetId по имени листа (нужен для mergeCells, repeatCell и т.п.)."""
        try:
            meta = self._service.spreadsheets().get(
                spreadsheetId=self.spreadsheet_id
            ).execute()
        except HttpError as e:
            raise RuntimeError(f"Не удалось прочитать метаданные таблицы: {e}") from e
        for s in meta.get("sheets", []):
            props = s.get("properties", {})
            if props.get("title") == sheet_title:
                return int(props["sheetId"])
        raise KeyError(f"Лист не найден: {sheet_title!r}")

    def add_sheet(self, title: str) -> int:
        """Добавить лист; вернуть sheetId."""
        result = self.batch_update(
            [{"addSheet": {"properties": {"title": title}}}]
        )
        replies = result.get("replies") or []
        if not replies:
            raise RuntimeError("add_sheet: пустой ответ API")
        return int(replies[0]["addSheet"]["properties"]["sheetId"])


def main() -> None:
    load_dotenv(Path(__file__).resolve().parent / ".env")
    spreadsheet_id = os.environ.get("GOOGLE_SPREADSHEET_ID", "").strip()
    if not spreadsheet_id:
        print(
            "Задайте ID таблицы: переменная GOOGLE_SPREADSHEET_ID в .env или в окружении "
            "(фрагмент URL: https://docs.google.com/spreadsheets/d/<ID>/edit).",
            file=sys.stderr,
        )
        sys.exit(1)

    client = GoogleSheetsClient(spreadsheet_id)
    sheet_name = os.environ.get("GOOGLE_SHEET_NAME", "").strip()
    if not sheet_name:
        names = client.list_sheet_names()
        if not names:
            print("В таблице нет листов.", file=sys.stderr)
            sys.exit(1)
        sheet_name = names[0]

    rows = client.read_all_used(sheet_name)
    print(f"Лист «{sheet_name}», строк: {len(rows)}")
    for i, row in enumerate(rows, start=1):
        print(f"{i}: {row}")


if __name__ == "__main__":
    main()
