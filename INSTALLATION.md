# Развёртывание на сервере

Инструкция для Ubuntu/Debian. Сервис будет работать постоянно через **systemd** + **gunicorn**.

---

## Требования к серверу

- Ubuntu 22.04 / Debian 12 (или новее)
- Python 3.10+
- Открытый порт **6511** в firewall (или Nginx как обратный прокси)

---

## Шаг 1 — Подготовка окружения

```bash
# Обновить пакеты
sudo apt update && sudo apt upgrade -y

# Установить Python и pip (если не установлены)
sudo apt install -y python3 python3-pip python3-venv git
```

---

## Шаг 2 — Клонирование репозитория

```bash
# Клонировать в /opt (рекомендуемое место для сервисов)
sudo git clone https://github.com/myrisingsun/hierarchical-to-flat-app.git /opt/vor-lzk

# Создать системного пользователя для сервиса (без прав входа)
sudo useradd --system --no-create-home --shell /bin/false vorlzk

# Передать владение папкой этому пользователю
sudo chown -R vorlzk:vorlzk /opt/vor-lzk
```

---

## Шаг 3 — Виртуальное окружение и зависимости

```bash
cd /opt/vor-lzk

# Создать venv
sudo -u vorlzk python3 -m venv venv

# Установить зависимости + gunicorn
sudo -u vorlzk venv/bin/pip install --upgrade pip
sudo -u vorlzk venv/bin/pip install -r requirements.txt gunicorn
```

---

## Шаг 4 — Проверка запуска вручную

```bash
cd /opt/vor-lzk
sudo -u vorlzk venv/bin/gunicorn --bind 0.0.0.0:6511 --workers 2 --worker-tmp-dir /tmp app:app
```

Откройте в браузере `http://<IP-сервера>:6511` — должен открыться интерфейс.

Остановить: `Ctrl+C`. Переходим к автозапуску.

---

## Шаг 5 — Systemd-сервис

Создать файл юнита:

```bash
sudo nano /etc/systemd/system/vor-lzk.service
```

Вставить содержимое:

```ini
[Unit]
Description=VOR to LZK Transformation Web Service
After=network.target

[Service]
Type=simple
User=vorlzk
Group=vorlzk
WorkingDirectory=/opt/vor-lzk
ExecStart=/opt/vor-lzk/venv/bin/gunicorn \
    --bind 0.0.0.0:6511 \
    --workers 2 \
    --timeout 120 \
    --worker-tmp-dir /tmp \
    --access-logfile /var/log/vor-lzk/access.log \
    --error-logfile /var/log/vor-lzk/error.log \
    app:app
Restart=on-failure
RestartSec=5

[Install]
WantedBy=multi-user.target
```

Создать папку для логов:

```bash
sudo mkdir -p /var/log/vor-lzk
sudo chown vorlzk:vorlzk /var/log/vor-lzk
```

Включить и запустить сервис:

```bash
sudo systemctl daemon-reload
sudo systemctl enable vor-lzk     # автозапуск при перезагрузке сервера
sudo systemctl start vor-lzk
sudo systemctl status vor-lzk     # должен показать: active (running)
```

---

## Шаг 6 — Firewall

Если используется `ufw`:

```bash
sudo ufw allow 6511/tcp
sudo ufw reload
```

Сервис доступен напрямую по адресу `http://<IP-сервера>:6511`.

---

## Шаг 7 (опционально) — Nginx как обратный прокси

Если нужен домен или HTTPS — поставьте Nginx перед gunicorn.

```bash
sudo apt install -y nginx
```

Создать конфиг:

```bash
sudo nano /etc/nginx/sites-available/vor-lzk
```

```nginx
server {
    listen 80;
    server_name vor.example.com;   # ← замените на ваш домен

    client_max_body_size 50M;      # лимит загрузки файла

    location / {
        proxy_pass         http://127.0.0.1:6511;
        proxy_set_header   Host $host;
        proxy_set_header   X-Real-IP $remote_addr;
        proxy_read_timeout 120s;
    }
}
```

```bash
sudo ln -s /etc/nginx/sites-available/vor-lzk /etc/nginx/sites-enabled/
sudo nginx -t            # проверить конфиг
sudo systemctl reload nginx
```

После этого порт 6511 можно закрыть от внешнего доступа (трафик идёт через Nginx):

```bash
sudo ufw delete allow 6511/tcp
sudo ufw reload
```

---

## Управление сервисом

| Команда | Действие |
|---|---|
| `sudo systemctl status vor-lzk` | Статус |
| `sudo systemctl restart vor-lzk` | Перезапуск |
| `sudo systemctl stop vor-lzk` | Остановка |
| `sudo systemctl disable vor-lzk` | Убрать из автозапуска |
| `sudo journalctl -u vor-lzk -f` | Логи в реальном времени |
| `sudo tail -f /var/log/vor-lzk/error.log` | Ошибки gunicorn |

---

## Обновление до новой версии

```bash
cd /opt/vor-lzk

# Получить изменения из репозитория
sudo -u vorlzk git pull

# Обновить зависимости (если изменился requirements.txt)
sudo -u vorlzk venv/bin/pip install -r requirements.txt

# Перезапустить сервис
sudo systemctl restart vor-lzk
```

---

## Разработка

**Руководитель отдела архитектуры и проектирования — Пахарев Кирилл**
kpakharev@afid.ru
