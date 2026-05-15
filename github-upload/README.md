# STO Учёт — GitHub Pages + Google Apps Script

## Что содержит папка
- `index.html` — основной SPA фронтенд
- `sklad.html` — модуль склада, подгружается динамически
- `Code-v3.gs` — Google Apps Script backend для работы с Google Sheets

## Как использовать
1. Загрузите `index.html` и `sklad.html` в репозиторий GitHub.
2. Включите GitHub Pages для репозитория, чтобы `index.html` стал доступен как статический сайт.
3. Разверните `Code-v3.gs` как Google Apps Script Web App:
   - Откройте новый проект Apps Script
   - Вставьте содержимое `Code-v3.gs`
   - Сохраните и опубликуйте веб-приложение
   - Разрешения: "Anyone, even anonymous" (или вариант с Google-аккаунтами, если хотите)
4. Скопируйте URL опубликованного Apps Script в поле `SCRIPT_URL` на странице `index.html`.

## Важные детали
- `index.html` загружает `sklad.html` через `fetch('sklad.html')`, поэтому оба файла должны находиться в одной папке на GitHub Pages.
- Фронтенд работает как мобильное/псевдомодальное приложение.
- Google Sheets используется только на стороне Apps Script.

## Проверка локально
Для локальной проверки сначала запустите простой HTTP-сервер, чтобы браузер разрешил `fetch`:
```bash
python -m http.server 8000
```
Затем откройте `http://127.0.0.1:8000/github-upload/index.html`.

## Установка
- Убедитесь, что в `index.html` настроено хранение `SCRIPT_URL` в `localStorage`.
- В поле настройки вставляйте именно URL вида:
  `https://script.google.com/macros/s/XXXXX/exec`

## Примечание
Скрипт `Code-v3.gs` должен быть опубликован как Web App и доступен для вызова из браузера. Этот репозиторий содержит только фронтенд и серверный Apps Script для Google Sheets.
