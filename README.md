# Анализ поисковых запросов из яндекс директ пословно

Скрипт распрасивает все поисковые запросы из отчета Яндекс Директ и группирует их по словам.
Примерно также работает группировка в KeyCollector

## Подготовка
### Библиотеки
1. `openpyxl` для загрузки и подготовки файла
```
pip3 install openpyxl
```
2. `pandas` для работы с datafram из выгрузки
```
pip3 install pandas
```

### Выгруженный файл
Файл выгрузки требуется в формате `xlsx` с именем `export.xlsx`
Столбцы: 
- Поисковый запрос. Он в отчёте по поисковым фразам будет по дефолту 
- Показы
- Клики 
- Расход (руб.)
- Конверсии

> [!TIP]
> Что бы выгрузить корректную статистику по отдельной конверсии, нужно в отчёте в фильтре выбрать «Цель»

## Итог
Скрипт создаст еще один файл - `ouput.xlsx` в нём итог работы группировки

