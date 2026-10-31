# ВЗМК Паспортизация r12 / CODEX39

Review-ветка codex/review-production-r12-codex39. База — r11 e00e40aad873980c4803069955f9c19be77d1d44. При УЗК = 0 отсутствуют акт УЗК и ссылки на него, обязательные ВИК и АКЗ сохраняются. Списание нулевого плана не меняет УЗК.xlsm. При УЗК > 0 действует прежний путь.

Реальная заявка заблокирована некорректными B. Проверенного PDF реальной отгрузки и успешного полного DEV-цикла нет. PRODUCTION READY не присвоен.

Изменение и границы доказательств: docs/CODEX39_REVIEW.md. Результаты: review/CODEX39_ACCEPTANCE.json и review/REGRESSION_SUMMARY.json. Контроль входов: review/BUILD_INPUTS_R12.json и review/SOURCE_MANIFEST.json.

Запуск: production.py. Сборка: python -B tests/production/build.py --name reviewr12. Windows, Python 3.12, Microsoft Word/Excel; requirements.txt, requirements_stage5.txt, PyInstaller, pywin32, pypdf и ReportLab 4.x. Новые адресные тесты: python -B -m unittest discover -s tests/production -p test_codex39.py -v.

Сохранено прежнее обезличивание четырёх исторических runtime-модулей. Их различия от локальных входов EXE указаны в манифесте. Все шесть изменённых runtime-модулей совпадают полностью. Реальные документы и частные настройки не публикуются. Main и r11 не изменяются, merge не выполняется.
