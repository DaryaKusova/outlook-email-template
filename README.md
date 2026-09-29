# Шаблон письма сотрудникам

Адаптивное письмо о запуске нового рабочего сервиса. Одноколоночная вёрстка шириной 640 px, системный шрифт Arial, две иллюстрации, нумерованные разделы и светло-голубой блок регистрации. Содержание универсальное: перед использованием замените все заполнители в квадратных скобках и обе ссылки `https://example.com`.

## Файлы

- `template/employee-announcement.mjml` — исходник для редактирования.
- `template/employee-announcement.html` — готовая HTML-вёрстка после сборки.
- `template/assets/digital-documents.png` — иллюстрация совместной работы с цифровыми документами.
- `template/assets/registration.png` — иллюстрация приглашения и регистрации.
- `template/assets/prompts.md` — исходные английские промпты иллюстраций.

Иллюстрации созданы встроенным инструментом image_gen: мягкий объём, светло-голубой фон, бирюзовые и небольшие золотистые акценты. Изображения содержат вымышленные сцены и абстрактные элементы интерфейса без брендов и персональных данных. Размер каждого PNG — 1774 × 887 px.

Название «ВАША КОМПАНИЯ» в шапке и подписи можно заменить текстом организации. Для логотипа добавьте `mj-image` с альтернативным текстом и явными размерами; способ доставки изображения необходимо настроить в вашей системе рассылки.

## Установка и сборка

Требуется Node.js 20 или новее. Из папки проекта выполните:

```sh
npm ci
npm run build
```

MJML установлен локально, версия закреплена в `package-lock.json`. Сборка использует строгую проверку MJML и минификацию HTML. После изменения текста повторите сборку и откройте HTML в браузере. Сохраните папку `assets` рядом с HTML: изображения подключены по относительным путям для локального просмотра.

## Локальный шаблон для Outlook

В Windows с установленным классическим Outlook можно создать локальный файл `.oft`. Сначала настройте текст в MJML, выполните сборку и проверьте, что заполнители заменены. Затем запустите этот код PowerShell из папки проекта:

```powershell
$htmlPath = (Resolve-Path -LiteralPath '.\template\employee-announcement.html').Path
$templatePath = Join-Path (Split-Path -Parent $htmlPath) 'employee-announcement.oft'
$assetDirectory = Join-Path (Split-Path -Parent $htmlPath) 'assets'
$htmlContent = [System.IO.File]::ReadAllText($htmlPath, [System.Text.Encoding]::UTF8)
$assets = @(
    @{ File = 'digital-documents.png'; Cid = 'digital-documents' },
    @{ File = 'registration.png'; Cid = 'registration' }
)
$outlookApp = New-Object -ComObject Outlook.Application
$draft = $outlookApp.CreateItem(0)
$draft.BodyFormat = 2
$draft.Subject = 'Информация о новом сервисе'
foreach ($asset in $assets) {
    $assetPath = Join-Path $assetDirectory $asset.File
    $attachment = $draft.Attachments.Add($assetPath, 1)
    $attachment.PropertyAccessor.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F', $asset.Cid)
    $attachment.PropertyAccessor.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x370E001F', 'image/png')
    $attachment.PropertyAccessor.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x7FFE000B', $true)
    $htmlContent = $htmlContent.Replace('assets/' + $asset.File, 'cid:' + $asset.Cid)
}
$draft.HTMLBody = $htmlContent
$draft.SaveAs($templatePath, 2)
$draft.Display()
```

Код встраивает обе иллюстрации через CID-вложения, сохраняет шаблон рядом с HTML и открывает его в Outlook для проверки. Получатели не задаются, письмо не отправляется. Укажите тему, получателей и нужные вложения вручную. Готовый `.oft` можно открывать двойным щелчком для создания нового письма.

Для отправки через другой сервис загрузите иллюстрации в его хранилище и замените относительные пути на доступные получателям HTTPS-адреса. Самостоятельное копирование HTML с локальными путями не обеспечивает доставку изображений.

## Проверки и ограничения

Выполнены строгая сборка MJML и проверка HTML в Chromium (Microsoft Edge) при ширине 375 и 900 px. Горизонтального переполнения нет. Поддержка тёмной темы зависит от почтового клиента.

Вёрстка использует таблицы и условные блоки MSO, которые создаёт MJML для классического Outlook. Отображение именно в Outlook 2019 и других почтовых клиентах не подтверждено тестовой доставкой. Перед рассылкой отправьте тестовое письмо себе и проверьте текст, ссылки, вложения и внешний вид.

Файлы `.oft`, `.msg`, `.eml`, архивы, локальные материалы, результаты проверок и переменные окружения исключены из Git.
