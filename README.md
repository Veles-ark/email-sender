# email-sender
A program for mass email distribution. Import addresses from Word, save to Excel, and send via SMTP. This program is distributed exclusively for educational purposes and is completely free!
After launching the program, a settings window opens. In the top line, select the Word file containing the email addresses for the distribution. The type, format, text chunks, and order within the Word file are not important. Email addresses are extracted from the Word file using a specific template. In the second line from the top, specify the path and name of the Excel file for saving the email addresses. After clicking the "Extract Addresses" button, the email addresses are extracted from the Word file and inserted into the left column of the Excel file. Then, select this Excel file in the Excel file for distribution field. In the "SMTP Server" field, specify the server address through which the messages will be sent. In the "Port and Mode" field, specify the sending port and mode. In the "Login" field, optionally specify a sending username. In the "Sender Email" field, specify the email address from which the messages will be sent. Accordingly, in the "Password" field, enter the password for the email address through which the messages will be sent. Since this example uses mail.ru settings and email, in the password field, you must specify the password for external applications previously set in the settings. In the subject, message text, and attachments fields, enter the appropriate information. If you leave the message text blank, create a Word file named "message" next to the program file and enter the text that should be in the message body. The message text will be taken from this file when sending the email. Next, in the "minutes between sendings" field, set the "from" and "to" values. I tried sending between 3 and 5 minutes. Everything works without blocking anything. Clicking the "save settings" button creates a JSON file, and the settings are taken from this file the next time the program is launched. Clicking the "start sending" button begins the sending process. The first address from the left column of the Excel file is taken, and an attempt is made to send the email to this address. A successful or error message is displayed in the progress window at the bottom. Everything is also written to the log file. After sending or an error, there is a delay in the interval specified in the "from" and "to" fields. If the sending was successful in Excel, the program writes the word "sent" in the column to the right of the address. If sending fails, the cell with the address is colored red, and "not sent" is written in the column to the right. If the process is interrupted or the program is closed, when restarting and clicking the "Start Sending" button, all addresses are resent. Addresses to the right of which the word "sent" appears are skipped. Addresses to the right of which the word "sent" appears empty or are written "not sent" are resent. This prevents duplicate messages.


Программа для массовой рассылки email сообщениям. Импорт адресов из Word, сохранение в Excel, отправка через SMTP. Данная программа распространяется изключительно в образовательных целях и абсолютно бесплатно!
После старта программы открывается окно с настройками. В верхней строке необходимо выбрать файл word формата, в котором собраны адреса email для рассылки. Вид, формат, куски текста и порядок внутри word файла не имеет значения. Email адреса выхватываются из word файла по определённому шаблону. Во второй строке сверху указываем путь и имя файла excel для сохранения адресов email.  После нажатия кнопки извлечь адреса, email адреса извлекаются из word файла и вставляются в левую колонку в excel файл. Затем нужно выбрать этот excel файл в поле excel файл для рассылки. Поле smtp сервер указываем адрес сервера, через который будет происходить отправка. Поле порт и режим указываем порт отправки и режим. Поле логин если нужно указываем логин отправки. В поле почта отправителя указываем почту с которой будет происходить отправка. Соответственно в поле пароль указываем пароль от почты, через которую будем отправлять. Так как в примере используются настройки, и почта mail.ru в поле пароль необходимо указать пароль для внешних приложений, установленный заранее в настройках. В поле, тема, текст сообщения и вложения указываем соответствующую информацию. Если текст сообщения оставить пустым, рядом с файлом программы создать word файл с названием сообщение и внутрь положить текст, который должен быть в тексте сообщения, то при отправке письма текст сообщения будет браться из этого файла. Далее в поле минут между отправками устанавливается от и до. Пробовались отправки от 3 до 5 минут. Всё работает ничего не блокируется. При нажатии на кнопку сохранить настройки создаётся json файл и в последующий запуске программы настройки берутся из этого файла. При нажатии на кнопку начать рассылку, начинается процесс рассылки. Берётся первый адрес из левой колонки файла excel, делается попытка отправить письмо на этот адрес. Внизу в окне процесса отображается успешная отправка или ошибка. Также всё записывается в лог файл. После отправки или ошибки происходит задержка в интервале указанных в полях от и до. Если отправка была успешна в excel, в правой колонке от адреса программа записывает слово отправлено. Если отправка не удалась, ячейка с адресом закрашивается в красный цвет и справа в столбце пишется не отправлено. Если процесс прерывается, или программу закрыли, при повторном запуске и нажатии кнопки начать рассылку, происходит проход по всем адресам сначала. Те адреса справа от которых надпись «отправлено» пропускаются. На те адреса справа от которых пусто или написано «не отправлено» происходит повторная отправка. Таким образом исключается дублирование писем.
# Программа для массовой рассылки email

Инструмент для отправки писем большому количеству адресатов.
Поддерживает импорт email из Word, сохранение в Excel и отправку через SMTP.

## Скачать программу (EXE)

Нажмите здесь, чтобы скачать последнюю версию:

https://github.com/Veles-ark/email-sender/releases/latest

После скачивания запустите `email_tool.exe`.



## Основные возможности

- Импорт email адресов из Word (.docx)
- Сохранение в Excel (.xlsx)
- **Автоматическая отправка писем по списку**
- Поддержка SMTP (SSL / STARTTLS / PLAIN)
- Вложения файлов к письмам
- Пауза между отправками (случайный интервал)
- Фиксация результата в Excel (отправлено/не отправлено)



## Как пользоваться (краткая инструкция)

1. Откройте программу `email_tool.exe`
2. Выберите Word-файл с адресами → нажмите **"Извлечь адреса"**
3. Выберите Excel со списком для рассылки
4. Заполните настройки SMTP и email отправителя
5. Напишите тему и текст письма
6. При необходимости добавьте вложения
7. Нажмите **"НАЧАТЬ РАССЫЛКУ"**



## SMTP Настройки (Примеры)

**Mail.ru**
SMTP сервер: smtp.mail.ru
Порт: 465
Режим: SSL
Логин: тот же, что почта
Пароль: от почты или пароль приложения



**Gmail**
smtp.gmail.com, порт 587, STARTTLS
(Нужен пароль приложения Google)




**Yandex**
smtp.yandex.ru, порт 465, SSL






Запуск из исходников (если есть Python)

pip install python-docx openpyxl
python email_tool_gui29.py



Примечание
Автор не несет ответственности за использование программы в целях спама.
Пользователь самостоятельно отвечает за соблюдение законов и правил почтовых сервисов.
