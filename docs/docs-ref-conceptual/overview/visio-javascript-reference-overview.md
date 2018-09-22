# <a name="visio-javascript-api-overview"></a>Обзор Visio JavaScript API

С помощью API JavaScript для Visio вы можете внедрять схемы Visio в SharePoint Online. Внедренный документ Visio — схема, которая хранится в библиотеке документов SharePoint и отображается на странице SharePoint. Чтобы внедрить схемы Visio, их отображения в HTML- `<iframe>` элемент. После этого вы сможете программным способом работать с внедренным документом при помощи API JavaScript для Visio.

![Документ Visio в iframe на странице SharePoint вместе с веб-частью редактора сценариев](/javascript/api/docs-ref-conceptual/images/visio-api-block-diagram.png)


API JavaScript для Visio позволяет следующее:

* Описание взаимодействия с элементов схемы Visio, например страницы и фигур.
* Создание визуальной разметки на полотне схемы Visio.
* Написание пользовательские обработчики для событий мыши в документе.
* предоставлять своему решению данные документа, такие как текст фигуры, данные фигуры и гиперссылки.

В этой статье описано, как использовать API JavaScript для Visio с Visio Online, чтобы создавать решения для SharePoint Online. В ней представлены ключевые элементы, понимание роли которых крайне важно при использовании API, такие как прокси-объекты JavaScript, **EmbeddedSession**, **RequestContext**, а также методы **sync()**, **Visio.run()** и **load()**. В приведенных ниже примерах кода показано применение этих элементов.

## <a name="embeddedsession"></a>EmbeddedSession

Объект EmbeddedSession инициализирует взаимодействие между фреймом разработчика и фреймом Visio Online.

```js
var session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
session.init().then(function () {
    window.console.log("Session successfully initialized");
});
```

## <a name="visiorunsession-functioncontext--batch-"></a>Visio.Run (сеанса, function(context) {пакета})

Метод **Visio.run()** выполняет пакетный сценарий, совершающий действия с объектной моделью Visio. Пакетные команды включают определения локальных прокси-объектов JavaScript и методов **sync()**, синхронизирующих состояние объектов Visio и локальных объектов, а также разрешение обещания. Преимущество пакетной обработки запросов в методе **Visio.run()** состоит в том, что при разрешении обещания все отслеживаемые объекты страницы, выделенные во время выполнения, автоматически освобождаются.

Выполнить метод принимает сеанса и объект RequestContext и возвращает резервирование (как правило, только что результат **context.sync()**). Пакетную операцию можно выполнить, не указывая ее в методе **Visio.run()**. Однако в этом случае все ссылки на объекты страницы требуют отслеживания и управления вручную.

## <a name="requestcontext"></a>RequestContext

Объект RequestContext облегчает запросов для приложения Visio. Поскольку frame разработчика и Visio Online приложения выполняются в двух различных Интернет-кадров, объект RequestContext (контекста в следующем примере) требуется для получения доступа к Visio и связанных с ними объекты, такие как страницы и фигуры, из рамки для разработчиков.

```js
function hideToolbars() {
    Visio.run(session, function(context){
        var app = context.document.application;
        app.showToolbars = false;
        return context.sync().then(function () {
            window.console.log("Toolbars Hidden");
        });
    }).catch(function(error)
    {
        window.console.log("Error: " + error);
    });
};
```

## <a name="proxy-objects"></a>Прокси-объекты

Объекты JavaScript для Visio, объявленные и использованные в надстройке, — это прокси-объекты для реальных объектов в документе Visio. Все действия над прокси-объектами не реализуются в Visio, а состояние документа Visio — в прокси-объектах, пока оно не будет синхронизировано. Состояние документа синхронизируется при выполнении `context.sync()`.

К примеру локального getActivePage объект JavaScript объявляется ссылок на выбранную страницу. Это можно использовать для постановки в очередь настройки его свойств и вызова методов. Действия для таких объектов не реализуются, пока не будет запущено метод **sync()** .

```js
var activePage = context.document.getActivePage();
```

## <a name="sync"></a>sync()

Метод **sync()** синхронизирует состояние между объектами JavaScript прокси-сервера и реальные объекты в Visio, выполнив инструкции в очереди на контексте, а также извлечение свойств загружены объектов Office для использования в коде. Этот метод возвращает обещание, которое выполняется после завершения синхронизации. 

## <a name="load"></a>load()

Метод **load()** используется для заполнения прокси-объектов, созданных на уровне JavaScript надстройки. При попытке получения объекта, такого как документ, сначала на уровне JavaScript создается локальный прокси-объект. Такой объект можно использовать для добавления в очередь настройки его свойств и вызова методов. Но для чтения свойств или связей объекта сначала необходимо вызвать методы **load()** и **sync()**. Метод load() использует свойства и связи, которые требуется загрузить при вызове метода **sync()**.

Ниже представлен синтаксис метода **load()**.

```js
object.load(string: properties); //or object.load(array: properties); //or object.load({loadOption});
```

1. **Свойства** — это список имен свойств загруженных, указанный в качестве строки запятыми или массив имен. Дополнительные сведения см. в описаниях методов **.load()** под каждым объектом.

2. **loadOption** указывает объект, описывающий свойства select, expand, top и skip. Дополнительные сведения см. в статье, посвященной [параметрам загрузки объектов](/javascript/api/office/officeextension.loadoption).

## <a name="example-printing-all-shapes-text-in-active-page"></a>Пример. Печать текста всех фигур на активной странице

Приведенный ниже пример показывает, как распечатать значение текста фигуры из объекта фигур массива.
Метод **Visio.run()** содержит пакет инструкций. В рамках этого пакета создается прокси-объект, который ссылается на фигуры в активном документе.

Эти команды в очередь и запуск при вызове **context.sync()** . Метод **sync()** возвращает обещание, с помощью которого его можно связать с другими операциями.

```js
Visio.run(session, function (context) {
    var page = context.document.getActivePage();
    var shapes = page.shapes;
    shapes.load();
    return context.sync().then(function () {
        for(var i=0; i<shapes.items.length;i++) {
            var shape = shapes.items[i];
            window.console.log("Shape Text: " + shape.text );
        }
    });
}).catch(function(error) {
    window.console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        window.console.log ("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="error-messages"></a>Сообщения об ошибках

Ошибки возвращаются с помощью объекта ошибки, состоящего из кода и сообщения. В таблице ниже перечислены возможные ошибки.

| error.code            | error.message |
|-----------------------|----------------------------------------------------------------|
| InvalidArgument       | Аргумент недопустим, отсутствует или имеет неправильный формат. |
| GeneralException      | При обработке запроса возникла внутренняя ошибка. |
| NotImplemented        | Запрашиваемая функция не реализована.  |
| UnsupportedOperation  | Выполняемая операция не поддерживается. |
| AccessDenied          | Вы не можете выполнить запрашиваемую операцию. |
| ItemNotFound          | Запрашиваемый ресурс не существует. |

## <a name="get-started"></a>Начало работы

Пример в этом разделе можно использовать для начала работы. В этом примере показано, как программно отобразить текст фигуры от выбранной фигуры в схемы Visio. Чтобы приступить к работе, создайте страницу классический в SharePoint Online или редактирование существующей страницы. Добавление веб-части редактора скрипт на странице и копирование и вставка приведенный ниже код.

```js
<script src='https://appsforoffice.microsoft.com/embedded/1.0/visio-web-embedded.js' type='text/javascript'></script>

Enter Visio File Url:<br/>
<script language="javascript">
document.write("<input type='text' id='fileUrl' size='120'/>");
document.write("<input type='button' value='InitEmbeddedFrame' onclick='initEmbeddedFrame()' />");
document.write("<br />");
document.write("<input type='button' value='SelectedShapeText' onclick='getSelectedShapeText()' />");
document.write("<textarea id='ResultOutput' style='width:350px;height:60px'> </textarea>");
document.write("<div id='iframeHost' />");

let session; // Global variable to store the session and pass it afterwards in Visio.run()
var textArea;
// Loads the Visio application and Initializes communication between developer frame and Visio online frame
function initEmbeddedFrame() {
    textArea = document.getElementById('ResultOutput');
    var url = document.getElementById('fileUrl').value;
    if (!url) {
        window.alert("File URL should not be empty");
    }
    // APIs are enabled for EmbedView action only.
    url = url.replace("action=view","action=embedview");
    url = url.replace("action=interactivepreview","action=embedview");
    url = url.replace("action=default","action=embedview");
    url = url.replace("action=edit","action=embedview");
  
    session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
    return session.init().then(function () {
        // Initialization is successful
        textArea.value  = "Initialization is successful";
    });
}

// Code for getting selected Shape Text using the shapes collection object
function getSelectedShapeText() {
    Visio.run(session, function (context) {
        var page = context.document.getActivePage();
        var shapes = page.shapes;
        shapes.load();
        return context.sync().then(function () {
            textArea.value = "Please select a Shape in the Diagram";
            for(var i=0; i<shapes.items.length;i++) {
                var shape = shapes.items[i];
                if ( shape.select == true) {
                    textArea.value = shape.text;
                    return;
                }
            }
        });
    }).catch(function(error) {
        textArea.value = "Error: ";
        if (error instanceof OfficeExtension.Error) {
            textArea.value += "Debug info: " + JSON.stringify(error.debugInfo);
        }
    });
}
</script>
```

После этого все, что требуется — это URL-адрес схемы Visio, которое вы хотите работать с. Только что загрузите схемы Visio в SharePoint Online и откройте его в документации по Visio. Нет откройте диалоговое окно внедрить и использовать URL-адрес внедрить в приведенном выше примере.

![Скопируйте URL-адрес файла Visio из диалогового окна Embed](/javascript/api/docs-ref-conceptual/images/Visio-embed-url.png)

При использовании Visio Online в режиме редактирования, откройте диалоговое окно Embed, выбрав **файл** > **папки** > **Embed**. Если вы используете Visio Online в режим просмотра, откройте диалоговое окно Embed, выбрав «...», а затем **внедрить**.

## <a name="open-api-specifications"></a>Открытые спецификации API

Мы публикуем новые API на странице [Открытые спецификации API](../openspec.md), чтобы вы могли делиться своим мнением о них. Узнайте, над какими функциями мы работаем, и поделитесь своим мнением о спецификациях.

## <a name="visio-javascript-api-reference"></a>Справочник по Visio JavaScript API

Подробные сведения о Visio JavaScript API обратитесь к [Справочная документация по Visio JavaScript API](/javascript/api/visio).
