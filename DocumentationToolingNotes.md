# <a name="how-the-office-javascript-api-documentation-is-generated"></a>Сведения о том, как создается документация по API JavaScript для Office

Страницы справочной документации по JavaScript для Office создаются на основе файлов определений типов и примеров фрагментов. В этом процессе используется смешение средств с открытым кодом и сценариев, характерных для репозитория. Этот документ предназначен для того, чтобы сделать процессы этого репозитория прозрачными, чтобы сообщество могло лучше пользоваться этим содержимым и вносить в него изменения.

## <a name="content-sources"></a>Источники контента

Для создания справочной документации по Office – JS используются два типа контента: определения типов и фрагменты кода. Они обеспечивают полное покрытие API и предоставляют небольшие примеры встроенного кода.

### <a name="type-definition-files"></a>Файлы определений типов

Файлы определения типов по [определенному](https://github.com/DefinitelyTyped/DefinitelyTyped) типу представляют собой единственный источник истинности документации. Все надстройки Office, использующие TypeScript, компилируются с использованием этих файлов определений типов. Они также предоставляют разработчикам функций IntelliSense и TypeScript поддержку IntelliSense. Создав справочную документацию из этих определений, мы предоставляем более точную информацию.

Существует четыре релевантных файла d. TS, которые предоставляют исходное содержимое для разных подразделов документов.

- [Office-JS/index. d. TS](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts) (определения выпуска).
  - [Excel (выпуск)](https://docs.microsoft.com/javascript/api/excel_release)
  - [OneNote](https://docs.microsoft.com/javascript/api/onenote)
  - [PowerPoint](https://docs.microsoft.com/javascript/api/powerpoint)
  - [Visio](https://docs.microsoft.com/javascript/api/visio)
  - [Word (выпуск)](https://docs.microsoft.com/javascript/api/word_release)
  - [Подраздел Оффицеекстенсионс общего API](https://docs.microsoft.com/javascript/api/office)
- [Office-JS-Preview/index. d. TS](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) (определения предварительного просмотра).
  - [Excel (Предварительная версия)](https://docs.microsoft.com/javascript/api/excel)
  - [Outlook (Предварительная версия)](https://docs.microsoft.com/javascript/api/outlook)
  - [Word (Предварительная версия)](https://docs.microsoft.com/javascript/api/word)
  - [Общий API](https://docs.microsoft.com/javascript/api/office)
- [Custom-Function-Runtime/index. d. TS](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts) (определения среды выполнения пользовательских функций Excel).
  - [Пользовательские функции](https://docs.microsoft.com/javascript/api/custom-functions-runtime)
- [Office-Runtime/index. d. TS](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts) (определения среды выполнения Office для платформы пользовательских функций).
  - [Среда выполнения Office](https://docs.microsoft.com/javascript/api/office-runtime)

Более ранние версии API имеют собственные файлы d. TS. Они сохраняются при отпускании новой версии набора обязательных элементов API. Они также могут быть созданы с помощью [средства версии ремовер](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/tools/VersionRemover.ts). Эти старые файлы d. TS поддерживаются таким образом, что в API событий исправлены или изменены, исходное поведение по-прежнему задокументировано. Это полезно, если необходимо ориентироваться на более старую версию API.

#### <a name="testing-type-definition-file-changes"></a>Тестирование изменений в файле определений типов

Все изменения документации для API JavaScript для Office выполняются путем изменения четырех файлов d. TS, упомянутых выше. Тем не менее, вы можете протестировать изменения перед отправкой DefinitelyTyped (если необходимо, например, протестировать преобразование форматирования в Markdown), изменив соответствующий файл в разделе [Generate-Files/Script-Inputs](https://github.com/OfficeDev/office-js-docs-reference/tree/master/generate-docs/script-inputs) и запуска [женератедокс. cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd). При появлении соответствующего запроса выберите параметр "Локальные файлы".

Если вы помещаете изменения в удаленную ветвь этого репозитория, платформа docs.microsoft.com создает тестовую ветвь. Эта ветвь отображается в review.docs.microsoft.com, который доступен только внутренним сотрудникам Майкрософт. Любой пользователь, просматривающий пр, проверит сайт проверки на точность.

### <a name="code-snippets"></a>Фрагменты кода

Фрагменты кода примеров кода добавляются на эталонные страницы из двух источников:

- [Примеры сценариев лаборатории](https://github.com/OfficeDev/office-js-snippets)
- [Локальные фрагменты кода](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/code-snippets)

Локальные фрагменты находятся в файлах ямл для конкретных узлов. Их содержимое упорядочено по классу и полю, поэтому его можно сопоставить с соответствующим местом на странице справки. Язык фрагмента кода (JavaScript или TypeScript) определяется с помощью операторов await.

Фрагменты лабораторий сценариев извлекаются из рабочих примеров. В настоящее время примеры Excel и Word сопоставляются с разделами документа в [сочетании с файлами сопоставления](https://github.com/OfficeDev/office-js-snippets/tree/master/snippet-extractor-metadata). Они совпадают с отдельными примерами методов к свойствам или методам в API. При `yarn start` запуске репозитория Office-JS-Snippets создается [файл ямл](https://github.com/OfficeDev/office-js-snippets/blob/master/snippet-extractor-output/snippets.yaml) , содержащий все сопоставленные фрагменты. Этот файл ямл является входными данными в справочной документации.

## <a name="tooling-pipeline"></a>Программный конвейер

![Изображение, которое показывает потоки управления от неопределенного типа, до препроцессора, средства извлечения API, мидпроцессор, документации API и до препроцессора.](ToolingPipeline.png)

Содержимое документации между источниками контента и последними страницами проходит через пять этапов:

1. [Скрипт препроцессора](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/preprocessor.ts)
1. [Средство извлечения API](https://api-extractor.com/)
1. [Сценарий мидпроцессор](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/midprocessor.ts)
1. [Документ API](https://github.com/microsoft/rushstack/blob/master/apps/api-documenter/README.md)
1. [Скрипт для препроцессора](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/postprocessor.ts)

Препроцессор выполняет файлы d. TS и разделяет их на разделы, зависящие от узла. Он выполняет все необходимые средства очистки для правильной обработки данных в последующих средствах.

Средство извлечения API преобразует файлы d. TS в данные JSON. В этом разделе заменяются все данные типа, что позволяет упростить анализ.

Мидпроцессор извлекает фрагменты кода и пары их с соответствующими узлами.

Document API преобразует данные JSON в файлы. yml. Файлы. yml преобразуются в Markdown с помощью открытой системы публикации, которая публикует документы в docs.microsoft.com. Document API также содержит расширение для Office, которое вставляет фрагменты кода.

На этом процессоре выполняется очистка оглавления и перемещение файлов. yml в [папку публикации](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen).

Все пять из этих действий выполняются при запуске [женератедокс. cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd) . Этот скрипт также обрабатывает установку модуля узла, очищает старые наборы файлов и файлы определений типов версий для каждого набора требований.
