[![Made with Python](https://img.shields.io/badge/Made%20with-Python-blue)](https://www.python.org/)

## Скрипт для отравки почты

Создает и отправляет письма с прикрепленными файлами контактам, указанным в Excel-файле.

### Технологии:

* Python 3.10.0
* Pandas 2.2.2

### Как запустить проект:

1. Клонировать репозиторий и перейти в него в командной строке:

```
git clone https://github.com/eva-bogd/send_email.git
```

```
cd send_email
```

2. Создать файл .env следующего и заполнить следующие переменные:
```
EMAIL_USER = адрес электронной почты отправителя
EMAIL_PASSWORD = пароль электронной почты отправителя
```

3. Добавить в директорию с проектом файл contacts.xlsx, добавить директорию файлами для вложения в письма. Про необходимости указать нужный SMTP-сервер и порт, сейчас указан mail.ru

4. Cоздать и активировать виртуальное окружение:

```
python -m venv env
```

```
venv/scripts/activate
```

5. Установить зависимости из файла requirements.txt:

```
pip install -r requirements.txt
```

6. В директории c файлом send_mail.py выполнить команду:

```
python send_mail.py
```
