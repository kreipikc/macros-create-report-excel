# ENG
## Information
- The `templates` folder contains a template for creating an advance report. 
- The `input.xlsm` file contains a suitable template for filling out the report, as well as a macro that needs to be configured.
- The `reports` folder contains an example of an advance report received upon completion of the script.

## How to start
### First step
Download python and download the necessary libraries from `requirements.txt`.

You need to go to the right directory and download all libraries:
```commandline
cd example/path/to/script_folder
pip intall -r requirements.txt
```

### Second step
Configure it config.py if necessary (is to you use templates, you don't need to).
The macro is already in the "**input.xlsm**" file, and its code can be viewed in "**macros_code.vb.txt** ". To work correctly, **you need to configure it in Excel**, namely specify the paths:
- `pythonExe`: path to the python interpreter;
- `PythonScript`: the path to the script `main.py`;
- `savePath`: the path to save the final report.

After setting up the macro, you can fill in the file and click on "Generate"

# RU
## Информация
- Папка `templates` содержит шаблон для создания авансового отчета. 
- Файл `input.xlsm` содержит подходящий шаблон для заполнения отчета, а также макрос, который необходимо настроить.
- Папка `reports` содержит пример авансового отчета, полученный по завершению работы скрипта.

## Как начать
### Первый шаг
Скачайте python и загрузите необходимые библиотеки с сайта `requirements.txt`.

Вам нужно перейти в нужный каталог и загрузить все библиотеки:
```commandline
cd example/path/to/script_folder
pip intall -r requirements.txt
```

### Второй шаг
Настройте его `config.py` при необходимости (если вы используете шаблоны, вам это не нужно).
Макрос уже находится в файле `input.xlsm`, а его код можно просмотреть в `macros_code.vb.txt`. Для корректной работы **вам необходимо настроить его в Excel**, а именно указать пути:
- `pythonExe`: путь к интерпретатору python;
- `PythonScript`: путь к скрипту `main.py`;
- `SavePath`: путь для сохранения окончательного отчета.