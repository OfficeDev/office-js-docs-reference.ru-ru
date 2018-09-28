# <a name="functionfile-element"></a><span data-ttu-id="a1b3c-101">Элемент FunctionFile</span><span class="sxs-lookup"><span data-stu-id="a1b3c-101">FunctionFile element</span></span>

<span data-ttu-id="a1b3c-p101">Указывает файл с исходным кодом для операций, доступных через те команды надстройки, для выполнения которых используется функция JavaScript, а не отображается пользовательский интерфейс. Элемент **FunctionFile** является дочерним для [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md). Атрибуту **resid** элемента **FunctionFile** присваивается значение атрибута **id** элемента **Url** в элементе **Resources**. Последний содержит URL-адрес HTML-файла, который содержит или загружает все функции JavaScript, используемые для выполнения команд надстройки без пользовательского интерфейса, как определено элементом [Control](control.md).</span><span class="sxs-lookup"><span data-stu-id="a1b3c-p101">Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI. The  **FunctionFile** element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md). The **resid** attribute of the **FunctionFile** element is set to the value of the **id** attribute of a **Url** element in the **Resources** element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="a1b3c-105">Ниже приведен пример использования элемента **FunctionFile** .</span><span class="sxs-lookup"><span data-stu-id="a1b3c-105">The following is an example of the  **FunctionFile** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="a1b3c-106">Необходимо вызвать JavaScript в HTML-файл, указанный в параметре элемент **FunctionFile** `Office.initialize` и определите именованных функций, которые принимают один параметр: `event`.</span><span class="sxs-lookup"><span data-stu-id="a1b3c-106">The JavaScript in the HTML file indicated by the  **FunctionFile** element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="a1b3c-107">Следует использовать функции `item.notificationMessages` API для указания о ходе выполнения, успех или сбой, связанный с пользователем.</span><span class="sxs-lookup"><span data-stu-id="a1b3c-107">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="a1b3c-108">Он должен также вызвать `event.completed` при завершении выполнения.</span><span class="sxs-lookup"><span data-stu-id="a1b3c-108">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="a1b3c-109">Имя функции используются в элементе **имяФункции** для кнопок без интерфейса пользователя.</span><span class="sxs-lookup"><span data-stu-id="a1b3c-109">The name of the functions are used in the **FunctionName** element for UI-less buttons.</span></span>

<span data-ttu-id="a1b3c-110">Ниже приведен пример HTML-файла для определения функции **trackMessage**.</span><span class="sxs-lookup"><span data-stu-id="a1b3c-110">The following is an example of an HTML file defining a **trackMessage** function.</span></span>

```js
Office.initialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

<span data-ttu-id="a1b3c-111">Приведенный ниже код показано, как реализовать функцию, используемых **имяФункции**.</span><span class="sxs-lookup"><span data-stu-id="a1b3c-111">The following code shows how to implement the function used by  **FunctionName**.</span></span>

```js
// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

// Your function must be in the global namespace.
function writeText(event) {

    // Implement your custom code here. The following code is a simple example.

    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
```

> [!IMPORTANT]
> <span data-ttu-id="a1b3c-112">Вызов **event.completed** сообщает, что были успешно обработке событий.</span><span class="sxs-lookup"><span data-stu-id="a1b3c-112">The call to  **event.completed** signals that you have successfully handled the event.</span></span> <span data-ttu-id="a1b3c-113">Если функция вызывается несколько раз, например при многократном выборе одной команды надстройки, все события автоматически помещаются в очередь.</span><span class="sxs-lookup"><span data-stu-id="a1b3c-113">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="a1b3c-114">Первое событие запускается автоматически, а другие ожидают в очереди.</span><span class="sxs-lookup"><span data-stu-id="a1b3c-114">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="a1b3c-115">Когда функция вызывает метод **event.completed**, для нее запускается следующий вызов в очереди.</span><span class="sxs-lookup"><span data-stu-id="a1b3c-115">When your function calls **event.completed**, the next queued call to that function runs.</span></span> <span data-ttu-id="a1b3c-116">Необходимо вызвать **event.completed**; в противном случае функции не будут запускаться.</span><span class="sxs-lookup"><span data-stu-id="a1b3c-116">You must call **event.completed**; otherwise your function will not run.</span></span>