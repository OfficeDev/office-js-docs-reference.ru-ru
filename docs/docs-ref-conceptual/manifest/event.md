# <a name="event-element"></a><span data-ttu-id="ae264-101">Элемент Event</span><span class="sxs-lookup"><span data-stu-id="ae264-101">Event element</span></span>

<span data-ttu-id="ae264-102">Определяет обработчик событий в надстройке.</span><span class="sxs-lookup"><span data-stu-id="ae264-102">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="ae264-103">`Event` Элемент в данный момент поддерживается только с Outlook в Интернете в Office 365.</span><span class="sxs-lookup"><span data-stu-id="ae264-103">The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="ae264-104">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ae264-104">Attributes</span></span>

|  <span data-ttu-id="ae264-105">Атрибут</span><span class="sxs-lookup"><span data-stu-id="ae264-105">Attribute</span></span>  |  <span data-ttu-id="ae264-106">Обязательный</span><span class="sxs-lookup"><span data-stu-id="ae264-106">Required</span></span>  |  <span data-ttu-id="ae264-107">Описание</span><span class="sxs-lookup"><span data-stu-id="ae264-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ae264-108">Тип</span><span class="sxs-lookup"><span data-stu-id="ae264-108">Type</span></span>](#type-attribute)  |  <span data-ttu-id="ae264-109">Да</span><span class="sxs-lookup"><span data-stu-id="ae264-109">Yes</span></span>  | <span data-ttu-id="ae264-110">Задает обрабатываемое событие.</span><span class="sxs-lookup"><span data-stu-id="ae264-110">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="ae264-111">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="ae264-111">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="ae264-112">Да</span><span class="sxs-lookup"><span data-stu-id="ae264-112">Yes</span></span>  | <span data-ttu-id="ae264-p101">Задает способ выполнения обработчика событий (асинхронное или синхронное). В настоящее время поддерживаются только синхронные обработчики событий.</span><span class="sxs-lookup"><span data-stu-id="ae264-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="ae264-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="ae264-115">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="ae264-116">Да</span><span class="sxs-lookup"><span data-stu-id="ae264-116">Yes</span></span>  | <span data-ttu-id="ae264-117">Задает имя функции для обработчика событий.</span><span class="sxs-lookup"><span data-stu-id="ae264-117">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="ae264-118">Атрибут Type</span><span class="sxs-lookup"><span data-stu-id="ae264-118">Type attribute</span></span>

<span data-ttu-id="ae264-p102">Обязательный. Указывает событие, при возникновении которого вызывается обработчик. В приведенной ниже таблице представлены допустимые значения этого атрибута.</span><span class="sxs-lookup"><span data-stu-id="ae264-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="ae264-122">Тип события</span><span class="sxs-lookup"><span data-stu-id="ae264-122">Event type</span></span>  |  <span data-ttu-id="ae264-123">Описание</span><span class="sxs-lookup"><span data-stu-id="ae264-123">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="ae264-124">Обработчик события будет вызван, когда пользователь отправляет сообщение или приглашение на собрание.</span><span class="sxs-lookup"><span data-stu-id="ae264-124">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="ae264-125">Атрибут FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="ae264-125">FunctionExecution attribute</span></span>

<span data-ttu-id="ae264-p103">Обязательный. ДОЛЖНО быть задано значение `synchronous`.</span><span class="sxs-lookup"><span data-stu-id="ae264-p103">Required. MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="ae264-128">Атрибут FunctionName</span><span class="sxs-lookup"><span data-stu-id="ae264-128">FunctionName attribute</span></span>

<span data-ttu-id="ae264-p104">Обязательный. Задает имя функции для обработчика событий. Это значение должно совпадать с именем функции в [файле функции](functionfile.md) надстройки.</span><span class="sxs-lookup"><span data-stu-id="ae264-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```