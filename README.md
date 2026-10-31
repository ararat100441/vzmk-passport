# ВЗМК Паспортизация r13 / CODEX40

Review-ветка codex/review-production-r13-codex40. База — r12 1bc03df2970a1421ef9c41313e864551e5ee7927. B и остаток блокируют выпуск для всех марок текущей ВОМ. Ошибки B других марок дают предупреждение. История F+ всех марок проверяется строго. Прежние B/F+ не исправляются; отрицательные остатки не превращаются в ноль.

Конкретная реальная ВОМ прошла полный frozen DEV-цикл до COMPLETE на точных копиях данных при УЗК=0. Recovery проверен на диагностических копиях фактической транзакции. Официального выпуска не было. Общий PRODUCTION READY не присвоен.

Подробности и границы доказательств: docs/CODEX40_REVIEW.md. Результаты: review/CODEX40_ACCEPTANCE.json и review/REGRESSION_SUMMARY.json. Входы сборки: review/BUILD_INPUTS_R13.json. Различия обезличенного снимка и приватной сборки: review/SOURCE_MANIFEST.json.

Запуск production.py. Сборка: python -B tests/production/build.py --name reviewr13. Windows, Python 3.12, Microsoft Word/Excel; requirements.txt, requirements_stage5.txt, PyInstaller, pywin32, pypdf, ReportLab 4.x. Адресные тесты: python -B -m unittest discover -s tests/production -p test_codex40.py -v.

Реальные документы, EXE и частные настройки остаются локально. Main и r12 не меняются. Сохраняется обезличивание четырёх исторических runtime-модулей; полная воспроизводимость EXE только из публичного снимка не заявляется.
