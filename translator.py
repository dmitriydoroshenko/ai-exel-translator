import json
import threading
from concurrent.futures import CancelledError
from dataclasses import dataclass
from typing import Dict, Iterable, List, Set
from openai import OpenAI

SYSTEM_ROLE_ESSENCE = (
    "## Role\n"
    "You are an expert Game Localization (L10N) Specialist and professional mobile game localizer. "
    "Your goal is to translate English mobile gaming market reports and game text into Simplified Chinese, "
    "ensuring the output is natural and uses industry-standard jargon used by developers and publishers.\n\n"

    "## Terminology & Style Guidelines\n"
    "- Do Not Translate Game Titles: Keep all game names/titles in their original English form.\n"
    "- Avoid Literalism: Do not translate word-for-word. Focus on industry 'jargon.'\n"
    "- Spending/Monetization:\n"
    "  * 'Non-paying players' -> 非付费玩家 / 零氪玩家\n"
    "  * 'Spending real money' -> 付费 / 氪金\n"
    "- Events & Scheduling:\n"
    "  * 'Global schedule' -> 全服统一日程 / 固定档期\n"
    "  * 'Progress in events' -> 推进活动进度\n"
    "- Tone: Professional, concise, and analytical. Use 'Game-speak.'\n\n"

    "## STRICT RULES (L10N)\n"
    "1. DO NOT translate game titles, bundle names, or offer names. Keep them in English."
    "2. Translate all other values (descriptions, analysis, labels) into Simplified Chinese using the guidelines above."
)

SYSTEM_ROLE_TECHNICAL = (
    "\n\n## STRICT RULES (TECHNICAL)\n"
    "3. Keep JSON keys unchanged."
    "4. Return ONLY a valid JSON object without any markdown formatting or extra text outside the JSON."
)

SYSTEM_ROLE = SYSTEM_ROLE_ESSENCE + SYSTEM_ROLE_TECHNICAL

@dataclass
class UsageTotals:
    prompt_tokens: int = 0
    completion_tokens: int = 0

    @property
    def total_tokens(self) -> int:
        return self.prompt_tokens + self.completion_tokens

class Translator:
    """Переводчик на базе OpenAI с батчингом и кешированием"""

    def __init__(
        self,
        api_key: str,
        model: str = "gpt-5.2",
        batch_size: int = 30,
        timeout_s: int = 30,
        price_in_per_1m: float = 1.75,
        price_out_per_1m: float = 14.00,
        system_role: str = SYSTEM_ROLE,
        cancel_event: threading.Event | None = None,
    ) -> None:
        self.api_key = api_key
        self.client = OpenAI(api_key=self.api_key)
        self.model = model
        self.batch_size = batch_size
        self.timeout_s = timeout_s
        self.system_role = system_role

        self.cancel_event = cancel_event

        self.price_in_per_1m = price_in_per_1m
        self.price_out_per_1m = price_out_per_1m
        self.usage = UsageTotals()

        self.cache: Dict[str, str] = {}

    def _check_cancel(self) -> None:
        if self.cancel_event is not None and self.cancel_event.is_set():
            raise CancelledError()

    def translate_batch(self, batch_dict: Dict[str, str]) -> Dict[str, str]:
        """Отправляет пачку {id: text} на перевод в OpenAI.
           Возвращает словарь с теми же ключами и переведёнными значениями
        """
        if not batch_dict:
            return {}

        self._check_cancel()

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": self.system_role},
                    {"role": "user", "content": json.dumps(batch_dict, ensure_ascii=False)},
                ],
                response_format={"type": "json_object"},
                timeout=self.timeout_s,
            )

            self._check_cancel()

            usage = getattr(response, "usage", None)
            if usage is not None:
                self.usage.prompt_tokens += usage.prompt_tokens
                self.usage.completion_tokens += usage.completion_tokens

            content = response.choices[0].message.content
            if content is None:
                raise RuntimeError("RESPONSE_ERROR: API вернул пустой ответ (None).")

            try:
                result = json.loads(content)
            except json.JSONDecodeError as e:
                raise RuntimeError(f"Чат вернул невалидный JSON: {e}. ") from e

            if not isinstance(result, dict):
                raise RuntimeError(f"Ожидался JSON object (dict), получено: {type(result).__name__}")

            return result if result else {}

        except CancelledError:
            raise

        except Exception as e:
            raise RuntimeError(f"API_ERROR: {e}") from e

    def ensure_translated(self, texts: Iterable[str]) -> None:
        """Переводит все строки, которых ещё нет в кеше, используя батчинг"""

        self._check_cancel()

        unique: List[str] = []
        for t in texts:
            if t and t not in self.cache:
                unique.append(t)

        if not unique:
            return

        for i in range(0, len(unique), self.batch_size):
            self._check_cancel()
            chunk = unique[i : i + self.batch_size]
            batch = {f"id_{j}": text for j, text in enumerate(chunk)}
            res = self.translate_batch(batch)

            for batch_id, trans_text in res.items():
                if (hash(batch_id) & 0xF) == 0:
                    self._check_cancel()
                orig_text = batch.get(batch_id)
                if orig_text is None:
                    continue
                self.cache[orig_text] = trans_text

        return

    def translate_texts(self, texts: Iterable[str]) -> Dict[str, str]:
        """Переводит набор строк и возвращает словарь {оригинал: перевод}"""

        unique: Set[str] = set()
        for t in texts:
            if t:
                unique.add(t)

        self.ensure_translated(unique)

        return {t: self.cache.get(t, t) for t in unique}

    @property
    def total_cost_usd(self) -> float:
        return (self.usage.prompt_tokens / 1_000_000 * self.price_in_per_1m) + (
            self.usage.completion_tokens / 1_000_000 * self.price_out_per_1m
        )
