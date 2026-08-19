# Influence Hub — лендинг

Статический сайт (один `index.html` + папка `assets/`). Двуязычный EN/RU,
переключатель языка сохраняется, форма заявки уходит на почту через Formspree.

## Как выложить на GitHub Pages

1. Создай на github.com пустой публичный репозиторий (например `influence-hub`).
2. В этой папке выполни:
   ```bash
   git init
   git add .
   git commit -m "Influence Hub landing"
   git branch -M main
   git remote add origin https://github.com/<твой-логин>/influence-hub.git
   git push -u origin main
   ```
3. В репозитории: **Settings → Pages → Build and deployment → Source: Deploy from a branch**,
   ветка `main`, папка `/ (root)`, **Save**.
4. Через ~1 минуту сайт будет доступен по адресу
   `https://<твой-логин>.github.io/influence-hub/`.

## Форма заявки (чтобы письма реально приходили)

Сейчас форма в демо-режиме: показывает «Заявка принята», но никуда не отправляет.
Чтобы заявки падали на почту:

1. Зарегистрируйся на [formspree.io](https://formspree.io) (бесплатный план).
2. Создай новую форму, привяжи свою почту.
3. Скопируй её endpoint (вида `https://formspree.io/f/abcdwxyz`).
4. В `index.html` найди строку:
   ```js
   const FORMSPREE_ENDPOINT = "https://formspree.io/f/REPLACE_ME";
   ```
   и подставь свой адрес вместо `REPLACE_ME`.
5. Закоммить и запушь — заявки начнут приходить на почту.

## Картинки и видео (по желанию)

Фоновое фото креаторов, картинка в hero и видео-полоса сейчас работают как
аккуратные цветовые фолбэки (файлов нет — секции не ломаются). Чтобы вернуть оригинал:

1. Скачай из проекта Claude Design файлы и положи в `assets/` под этими именами:
   - `assets/creators-bg.png` — фон в секции «Новые креаторы»
   - `assets/hero-pill-bite.png` — картинка внутри «пилюли» в заголовке
   - `assets/band-loop.mp4` — видео-полоса
2. Они подхватятся автоматически (сайт уже ссылается на эти пути).

Логотип сейчас — текстовый (шрифт Unbounded). При желании можно заменить на
`assets/influence-hub-logo.png`.

## Свой домен (по желанию)

GitHub Pages → Settings → Pages → Custom domain. Добавь домен и настрой DNS
(CNAME на `<твой-логин>.github.io`).
