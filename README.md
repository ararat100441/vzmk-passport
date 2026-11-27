# REVIEW: аудит GPT-6 (НЕ рабочий релиз)

Этот review-снимок содержит полные исправленные исходники на базе CODEX50
с отменой предварительного PDF и адресными исправлениями Hardlink и
производительности. Подробный аудит: docs/GPT6_PRODUCTION_AUDIT.md.

**Windows/E2E новой ветки не выполнялись. Собирать и использовать EXE в
production без испытаний нельзя.** Рабочие Excel не изменены.

---

# ВЗМК Паспортизация r14-final / CODEX50

Исправлена повторная hardlink-проверка служебного journal.json после записи. Исключение действует только в проверенном sealed-контексте NEEDS_READBACK; защита до записи и native Office readback сохранены.

59 адресных тестов PASS. Итоговый EXE прошёл полный GUI DEV-цикл на отдельных копиях реальных Excel до COMPLETE; после перезапуска незавершённых операций нет. Рабочая тестовая операция штатно отменена по явному разрешению владельца; предыдущие операции и паспорта сохранены. Официальный выпуск не выполнялся, контрольные значения учебные.

Подробности: docs/CODEX50_REVIEW.md. Обезличенный результат: review/CODEX50_ACCEPTANCE.json. Входы сборки: review/BUILD_INPUTS_CODEX50.json. Предыдущие документы и версии сохранены. Полная воспроизводимость EXE только из публичного обезличенного дерева не заявляется.

## Сборка кандидата 1.0.0 для Windows x64

Используйте отдельное окружение CPython 3.12 x64 и уникальное имя сборки.

```powershell
py -3.12 -m venv .venv
.venv\Scripts\python.exe -m pip install --only-binary=:all: --require-hashes -r requirements-win-py312.lock -r requirements-build-win-py312.lock
.venv\Scripts\python.exe -m pip check
.venv\Scripts\python.exe -B tests/production/build.py --name dev-candidate
```

Сборка проверяет установленные версии из обоих lock-файлов. EXE находится в `dist/dev-candidate/VZMK-Passport-1.0.0`, манифест версий и SHA-256 входов находится в `validation/dev-candidate/BUILD_INPUTS.json`. Для Windows-приёмки используйте проверенные DEV-копии документов и сохраните SHA-256 проверенного EXE вместе с манифестом.
