# <a name="javascript-api-for-office"></a>API JavaScript для Office

API JavaScript для Office позволяет создавать веб-приложения, которые взаимодействуют с объектными моделями ведущих приложений Office. Приложение будет ссылки на библиотеку office.js, являющийся загрузчик скрипта. Библиотека office.js загружает объектных моделей, которые могут быть применены в приложение Office, на котором работает надстройки. Можно использовать следующие объектной модели JavaScript:

- **Общие интерфейсы API** - интерфейсы API, которые были представлены в **Office 2013**. Это загружается для **всех ведущих приложений Office** и подключение приложения надстройки с помощью клиентского приложения Office. Объектная модель содержит API, которые специфичны для клиентов Office и API, которые могут быть применены несколько ведущих приложений Office клиента. Все это содержимое — в разделе **Общих API**. 

  **Outlook** также использует общий синтаксис API. Все содержимое псевдоним Office содержит объекты, которые можно использовать для написания сценариев, которые взаимодействуют с контентом в документы Microsoft Office, листы, презентации, почтовых элементов и проектов из Office Add-ins. Необходимо использовать следующие общие API-интерфейсы, если надстройка будет создания решений Office 2013 и более поздних версий. Эта объектная модель использует обратных вызовов.

- **Среды размещения интерфейсы API** - интерфейсы API, которые были представлены с **Office 2016**. Эта объектная модель предоставляет узла строго типизированные объекты, которые соответствуют обычных объектов, отображаемые при использовании клиентов Office и представляет будущее интерфейсы API JavaScript для Office. API-интерфейсы среды размещения в настоящее время включают интерфейсов API JavaScript Word и Excel JavaScript.

## <a name="supported-host-applications"></a>Поддерживаемые ведущие приложения

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [Shared API](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [PowerPoint и Project](requirement-sets/powerpoint-and-project-note.md) поддерживают надстроек, внесенные с помощью API JavaScript. Тем не менее в настоящее время у них нет API для конкретных узлов. Взаимодействие с этих узлов посредством API общих.

Дополнительные сведения о [поддерживаемых ведущих приложениях и других требованиях](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins).

## <a name="open-api-specifications"></a>Открытые спецификации API

Мы публикуем новые API для надстроек Office на странице [Открытые спецификации API](openspec.md), чтобы вы могли делиться своим мнением. Узнайте, над какими функциями мы работаем, и поделитесь своим мнением о создаваемых спецификациях.

## <a name="see-also"></a>См. также

- [Справочник по JavaScript API для Office](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)