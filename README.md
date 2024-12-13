# FillApp

Десктопное приложение написаное для упрощения работы кладовщика. Приложение написано на NextJS в паре с TypeScript на фронтенде и Go на бэкэнде.

## Основной функционал

Основной функционал приложения выборка продуктов из накладной и их заполнение в книгу(файл EXCEL с учетом продуктов) по заранее оговоренным правилам.

## Дополнительный функционал

Возможность заменять ошибочные варианты названий продуктов из накладных, заполняя dict.json, находящийся в одной папке с приложением.

Создание новой книги на основе существующей, с заполнением остатков.

Логгирование событий для упрощенной отладки приложения в случае возникновения ошибок.

## Инструкция

Чтобы собрать приложение, необходимо склонировать репозиторий и в терминале прописать `wails build`, после этого будет сгенерирован исполняемый файл.