"""Package the tested binary and an honest operational readiness statement."""
import json,shutil,sys,subprocess,zipfile
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from production_app.storage import read,save,sha

def main():
    checks=Path(r'C:\VZMKPassportChecks\20261018');release=Path(r'C:\VZMK-Passport\1.0.0-rc3-final')
    accepted=read(checks/'frozen-final2/RESULT.json');assert accepted['status']=='PASS'
    exe=release/'VZMK-Passport-1.0.0-rc3.exe';assert sha(exe)==accepted['exe_sha256']
    dist=ROOT/'dist/production100-final2/VZMK-Passport-1.0.0-rc3'
    for source in dist.rglob('*'):
        if source.is_file():assert sha(source)==sha(release/source.relative_to(dist))
    inventory=read(checks/'inventory_after_xlsm.json')
    assert all(sha(path)==h for path,h in inventory['files'].items()),'PRODUCTION SOURCE DRIFT'
    evidence=release/'Проверки';evidence.mkdir(exist_ok=False)
    for folder,label in [('frozen-final2','final'),('legacy0916','legacy0916')]:
        for path in (checks/folder).glob('*'):
            if path.is_file() and path.suffix in ('.json','.log'):shutil.copyfile(path,evidence/(label+'_'+path.name))
    for path,label in [('frozen-final2/e2e/RESULT.json','e2e.json'),('frozen-final2/fault/RESULT.json','fault-recovery.json'),
                       ('frozen-final2/gui-components/RESULT.json','gui-components.json'),('gui-native/VISUAL_RESULT.json','gui-native-prior-same-ui.json'),
                       ('frozen-final2/foreign-excel/RESULT.json','foreign-excel.json')]:
        shutil.copyfile(checks/path,evidence/label)
    shutil.copyfile(ROOT/'validation/production100-final2/BUILD_INPUTS.json',evidence/'BUILD_INPUTS.json')
    shutil.copyfile(checks/'build-final2.log',evidence/'BUILD.log')
    save(evidence/'SOURCE_INTEGRITY.json',{'status':'PASS','files':inventory['files'],'all_306_files_byte_identical':True})
    shutil.copyfile(ROOT/'docs/production_release_inputs.txt',release/'Что требуется для выпуска 1.0.0.txt')
    shutil.copyfile(ROOT/'docs/production_requirements_review.txt',release/'Сопоставление с ТЗ.txt')
    shutil.copyfile(ROOT/'docs/production_user_guide.txt',release/'Инструкция пользователя.txt')
    first='''ВЗМК — Паспортизация 1.0.0-rc3
Окончательная 1.0.0 пока НЕ ГОТОВА: прочтите «Что требуется для выпуска 1.0.0.txt».

Запустите VZMK-Passport-1.0.0-rc3.exe двойным щелчком.
Переносите всю папку вместе с _internal. Python и Codex не нужны.
Требуются Windows, Microsoft Word и Excel. Установка компонентов не требуется.

В «Настройки путей» сами выберите реестр, корень документов и папку сохранения.
Никакие рабочие пути или DEV автоматически не назначаются.
Настройки: %LOCALAPPDATA%\\VZMKPassportProduction\\settings.json.
Выберите объект, затем новый паспорт; номер рассчитывается из текущего реестра.
ТЭМПО 1 в проверенном реестре: последний 18/1/2026, следующий 19/1/2026.

Привяжите собственные совместимые утверждённые формы объекта и реальную ВОМ.
Введите подтверждённые фактические результаты ВИК, УЗК и АКЗ и первичные протоколы.
Далее: Проверить данные → Подготовить документы → просмотр каждого файла
в Word/Excel → Проверить комплект → отдельное подтверждение выпуска.
До финального подтверждения реестр не меняется.
После сбоя используйте «Восстановление», а не повторный выпуск.

Сейчас отсутствует принятый полный набор реальных входов ТЭМПО 1;
исторические формы требуют адаптации. Не подтверждайте пробный настоящий выпуск.
Учебные испытания из папки «Проверки» не являются результатами контроля изделий.
'''
    (release/'НАЧАТЬ ЗДЕСЬ.txt').write_text(first,encoding='utf-8-sig')
    report=f'''Результат проверки 18.10.2026
ЭКСПЛУАТАЦИОННАЯ ГОТОВНОСТЬ 1.0.0: BLOCKED.
Проверяемый кандидат: 1.0.0-rc3, существующий production_app от c746d51.
EXE SHA256: {sha(exe)}

PASS — 28 проверок безопасности и 6 проверок сохранности реестра.
PASS — неизменённая 0.9.16: 104 марки, 288 изделий, 15 822,14 кг,
       6 физических файлов, 7 записей, без осиротевших процессов Office.
PASS — PyInstaller, самостоятельная EXE; smoke и Office-сценарии без Python в PATH.
PASS — запуск текущей EXE и Tk-компоненты её поставленных исходников.
       Главное окно и пустые настройки визуально проверены у предыдущего
       бинарника этой сессии с байт-в-байт тем же gui.py. После новой сборки
       визуальная перепроверка не выполнялась: пользователь остановил
       computer-use клавишей Esc. Самостоятельный запуск текущей EXE проверен.
PASS — реальный первый комплект ТЭМПО 1 открыт в Word/Excel только для чтения;
       исходный номер 1/1/2026 сохранён. Сверка с реальной ВОМ: NOT_VERIFIED.
PASS — полный цикл frozen EXE на изолированной учебной копии:
       CHECK → PREPARE → VERIFY → CONFIRM → RELEASE → backup/COMMIT →
       READBACK в Office → COMPLETE. Основной реестр не изменялся.
PASS — отказы при TEST вместо ACTUAL, отсутствии подтверждения, SHA drift,
       занятом файле/параллельном экземпляре, дублях, неизвестных изменениях.
PASS — искусственный I/O-сбой после ReplaceFile; откат с точным исходным SHA,
       повторное восстановление; ранее созданные документы не изменились.
PASS — сохранение/повторное чтение настроек в Tk-компонентах поставленных исходников.
PASS — 306 реальных исходных файлов повторно сверены по полному SHA256.

НЕ ВЫПОЛНЕНО: настоящий выпуск, запись в основной реестр, полный пользовательский
GUI-сценарий с настоящими протоколами, утверждение новых форм, merge/push PR №5.
Полный цикл проверялся через рабочие операции EXE на учебной копии; это не
равнозначно приёмке настоящих актов и не подтверждает фактический контроль.

Устранено: отсутствие обязательных результатов ВИК; повторное использование
исторических замеров АКЗ без полей; отсутствующая ранняя проверка структуры
форм; чтение XLSM стандартной структуры через приватную копию без макросов;
контроль даты/границ таблицы/области печати; доступность длинной формы фактов
через прокрутку; сброс устаревших подтверждений при изменении входов/шаблонов.
Изменение подготовленного документа теперь блокирует повторное одобрение
комплекта: SHA фиксируются при подготовке, а не после ручной правки.
Ранние испытания выявили ошибку проверки области печати; исправлена,
перепроверена шестью тестами и полным frozen Office E2E. Первый наблюдатель
искусственного сбоя мешал атомарной замене журнала и не достиг нужной точки;
это испытание признано несостоявшимся. Учебная операция восстановлена штатно.
Наблюдатель исправлен, затем отдельно проверен сбой именно после ReplaceFile.
Старые логи сохранены
в C:\\VZMKPassportChecks\\20261018, а не удалены.

Блокеры: несовместимость исторических утверждённых форм с движком {{{{ПОЛЕ}}}},
неопределённая реальная ВОМ конкретной отгрузки/первого паспорта, отсутствие
принятого набора первичных фактических протоколов и утверждённых назначений.
Нельзя устранить их заменой материалов ТЭМПО данными Казахстана или тестами.
Подробные пути и действия: «Что требуется для выпуска 1.0.0.txt».

Исходники исправлений: {ROOT}
Git изменения локальные, без commit/push/merge. Исходный старый main сохранён.
Сборки 0.9.16 и RC1 не удалялись и не заменялись.
Визуальное управление computer-use завершено по физической клавише Esc
пользователя. Последующие действия в интерфейсе не выполнялись.
'''
    (release/'Результат приёмки.txt').write_text(report,encoding='utf-8-sig')
    with zipfile.ZipFile(release/'Исходники исправлений.zip','w',zipfile.ZIP_DEFLATED) as z:
        packages=['production_app','passport_app','stage9_app','profiles_stage4','uzk_stage5','akz_stage6','stage7_quality','stage8_closure']
        paths=[ROOT/'production.py',ROOT/'production_100.spec',ROOT/'docs/production_mode_requirements.md',ROOT/'docs/production_release_inputs.txt',ROOT/'docs/production_user_guide.txt',ROOT/'docs/production_requirements_review.txt',
               *[p for package in packages for p in (ROOT/package).glob('*.py')],*(ROOT/'tests/production').glob('*.py')]
        for path in paths:z.write(path,str(path.relative_to(ROOT)))
    manifest={str(p.relative_to(release)):sha(p) for p in release.rglob('*') if p.is_file() and p.name!='SHA256SUMS.txt'}
    (release/'SHA256SUMS.txt').write_text('\n'.join(h+'  '+name for name,h in sorted(manifest.items()))+'\n',encoding='utf-8-sig')
    save(checks/'DELIVERY.json',{'folder':str(release),'exe_sha256':sha(exe),'files':len(manifest),'readiness':'BLOCKED','checks':'PASS'})
    print(json.dumps({'folder':str(release),'files':len(manifest),'exe_sha256':sha(exe),'readiness':'BLOCKED'},ensure_ascii=False))

if __name__=='__main__':main()
