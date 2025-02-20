# RU
## Для создания отчетов
Вам нужен zip-файл `dist/createReportsScript.zip` и файл `input.xlsm`.
Скачиваете архив и файл.

Далее заходите в `input.xlsm` и открыаете макрос `DoDocument`:

![изображение](https://github.com/user-attachments/assets/8eafa2fc-f86c-47ac-9b1e-08392470b2af)


Нажимаете **"Изменить"** и указываете полный путь до `createReportsScript.exe` в переменную `exePath`, сохраняете и готово.

![изображение](https://github.com/user-attachments/assets/da776453-233d-40ed-b737-f10edcff0ad9)


Теперь при заполнении данными файла `input.xlsm` и нажатии на кнопку **"Сформировать документы"** у вас создаться в той же директории папка `reports` со всеми отчетами в форматах:

- Word (.docx)
- Excel (.xlsx)
- PDF (.pdf)

![изображение](https://github.com/user-attachments/assets/50ad3671-1cc9-4a9f-9641-7ae357ef626d)


### Информация
В папке `dist` находиться папка с готовым скриптом в виде исполняемого файла для Windows (.exe).

В папке `reports` лежат примеры созданных отчетов с помощью шаблонов и исходного файла.

В папке `templates` лежать шаблоны для создания отчетов.

# ENG
## To create reports
You need the archive `dist/createReportsScript.zip` and the file `input.xlsm'.
Download the archive and file.

Next, go to `input.xlsm` and open the macro `DoDocument`:

![изображение](https://github.com/user-attachments/assets/8eafa2fc-f86c-47ac-9b1e-08392470b2af)

Click **"Edit"** and specify the full path to `createReportsScript.exe` to the `exePath` variable, save it, and you're done.

![изображение](https://github.com/user-attachments/assets/da776453-233d-40ed-b737-f10edcff0ad9)

Now, when filling in the input.xlsm file with data and clicking on the button **"Сформировать документы"** you will have a folder `reports` in the same directory with all reports in the following formats:

- Word (.docx)
- Excel (.xlsx)
- PDF (.pdf)

![изображение](https://github.com/user-attachments/assets/50ad3671-1cc9-4a9f-9641-7ae357ef626d)

### Information
The `dist` folder contains a folder with a ready-made script in the form of an executable file for Windows (.exe).

The `reports` folder contains examples of created reports using templates and a source file.

The `templates` folder contains templates for creating reports.
