# <a name="how-the-office-javascript-api-documentation-is-generated"></a>Как создается документация по API JavaScript для Office

Страницы справочной документации по JavaScript для Office создаются из файлов определений типов и примеров фрагментов кода. В этом процессе используется сочетание инструментов с открытым кодом и сценариев для репозитория. Этот документ призван сделать процессы этого репозитория прозрачными, чтобы сообщество было лучше пользоваться этим содержимым и вносить в него свой вклад.

## <a name="content-sources"></a>Источники контента

Для создания справочной документации Office-JS объединяются два типа контента: определения типов и фрагменты кода. Они обеспечивают полное охват API и дают небольшие примеры кода.

### <a name="type-definition-files"></a>Файлы определений типов

Файлы определений типов [в файле Definitely Typed](https://github.com/DefinitelyTyped/DefinitelyTyped) являются единственным источником информации для документации. Любая надстройка Office, использующая typeScript, компилирует файлы определений этих типов. Они также дают разработчикам JavaScript и TypeScript IntelliSense возможности. С помощью этих определений мы предоставляем более точные сведения.

Существует четыре соответствующих D.ts-файла, которые предоставляют исходный контент для разных подмещений документов.

- [office-js/index.d.ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts) (определения выпуска.)
  - [Excel (выпуск)](https://docs.microsoft.com/javascript/api/excel_release)
  - [OneNote](https://docs.microsoft.com/javascript/api/onenote)
  - [PowerPoint](https://docs.microsoft.com/javascript/api/powerpoint)
  - [Visio](https://docs.microsoft.com/javascript/api/visio)
  - [Word (выпуск)](https://docs.microsoft.com/javascript/api/word_release)
  - [Подраздел OfficeExtensions общего API](https://docs.microsoft.com/javascript/api/office)
- [office-js-preview/index.d.ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) (определения preview.)
  - [Excel (предварительная версия)](https://docs.microsoft.com/javascript/api/excel)
  - [Outlook (предварительная версия)](https://docs.microsoft.com/javascript/api/outlook)
  - [Word (предварительная версия)](https://docs.microsoft.com/javascript/api/word)
  - [Общий API](https://docs.microsoft.com/javascript/api/office)
- [custom-functions-runtime/index.d.ts](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts) (определения времени работы пользовательских функций Excel.)
  - [Пользовательские функции](https://docs.microsoft.com/javascript/api/custom-functions-runtime)
- [office-runtime/index.d.ts](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts) (определения времени работы Office для платформы пользовательских функций.)
  - [Office Runtime](https://docs.microsoft.com/javascript/api/office-runtime)

Более старые версии API имеют собственные D.ts-файлы. Они сохраняются при выпусках нового набора требований API. Они также могут быть созданы с помощью [средства удаления версий.](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/tools/VersionRemover.ts) Эти старые D.ts-файлы поддерживаются таким образом, чтобы в случае исправления или изменения API исходное поведение по-прежнему документировалось. Это полезно, если необходимо использовать более старую версию API.

#### <a name="testing-type-definition-file-changes"></a>Тестирование изменений файла определения типа

Любые изменения документации по API JavaScript для Office внося путем изменения четырех файлов D.ts, упомянутых выше. Однако можно протестировать изменение перед отправкой PR-файла в формат DefinitelyTyped (например, проверить, как форматирование будет преобразовываться в markdown), отредактировать соответствующий файл в [файле generate-docs/script-inputs](https://github.com/OfficeDev/office-js-docs-reference/tree/master/generate-docs/script-inputs) и задав [generateDocs.cmd.](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd) При запросе выберите параметр "Локальные файлы".

При внесения изменений в удаленную ветвь этого репо docs.microsoft.com платформы создается тестовая ветвь. Эта ветвь отрисовка review.docs.microsoft.com, доступной только внутренним сотрудникам Майкрософт. Все, кто просматривает ваш PR, проверяют правильность сайта проверки.

### <a name="code-snippets"></a>Фрагменты кода

Фрагменты кода добавляются на эталонные страницы из двух источников:

- [Примеры script Lab](https://github.com/OfficeDev/office-js-snippets)
- [Фрагменты кода локального кода](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/code-snippets)

Локальные фрагменты находятся в файлах yaml для определенных хостов. Их содержимое организовано по классу и полю, поэтому его можно сооставить с соответствующим местом на эталонной странице. Язык фрагмента кода (JavaScript или TypeScript) высмеяется использованием заявлений await.

Фрагменты кода Script Lab извлекаются из рабочих примеров. В настоящее время примеры Excel, Outlook, PowerPoint и Word сопоставлены для ссылок на разделы документа с помощью [файлов сопоставления.](https://github.com/OfficeDev/office-js-snippets/tree/prod/snippet-extractor-metadata) Они соответствуют отдельным образцам методов со свойствами или методами в API. При запуске репозитория office-js-snippets создается `yarn start` [файл yaml,](https://github.com/OfficeDev/office-js-snippets/blob/prod/snippet-extractor-output/snippets.yaml) содержащий все соединимые фрагменты кода. Этот файл yaml является входным в инструменты справочной документации.

## <a name="tooling-pipeline"></a>Конвейер инструментов

![Изображение, на котором показан поток управления от Definitely Typed к preprocessor, API Extractor, midprocessor, API Documenter и до postprocessor.](ToolingPipeline.png)

Между источниками контента и конечными страницами содержимое документации проходит пять этапов:

1. [Скрипт preprocessor](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/preprocessor.ts)
1. [Извлечения API](https://api-extractor.com/)
1. [Сценарий Midprocessor](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/midprocessor.ts)
1. [Документер API](https://github.com/microsoft/rushstack/blob/master/apps/api-documenter/README.md)
1. [Сценарий postprocessor](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/postprocessor.ts)

Препроцессор берет файлы d.ts и разбивает их на разделы, специфификшные для ведущего файла. Он выполняет очистку, необходимую для последующей обработки данных последующими средствами.

API Extractor преобразует D.ts-файлы в данные JSON. Это маркеризирует все данные типов, что упрощает анализ.

Midprocessor извлекает фрагменты кода и соещает их с правильными хостами и очищает перекрестные связи между объектами Outlook и общимИ API.

API Documenter преобразует данные JSON в YML-файлы. YML-файлы преобразуются в разметку системой открытой публикации, которая публикует документы в docs.microsoft.com. API Documenter также содержит расширение Office, которое вставляет фрагменты кода.

Послепроцессор очищает одержимую таблицу и перемещает YML-файлы в папку [публикации.](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen)

Все пять этих действий выполняются при запуске [GenerateDocs.cmd.](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd) Этот сценарий также обрабатывает установку модуля узла, очищает старые наборы файлов и файлы определения типов версий для каждого набора требований.
