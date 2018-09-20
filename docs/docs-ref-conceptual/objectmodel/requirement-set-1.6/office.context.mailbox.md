
# <a name="mailbox"></a><span data-ttu-id="39c84-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="39c84-101">mailbox</span></span>

### <span data-ttu-id="39c84-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="39c84-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="39c84-104">Предоставляет доступ для добавления в объектной модели Outlook для Microsoft Outlook и Microsoft Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="39c84-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="39c84-105">Требования</span><span class="sxs-lookup"><span data-stu-id="39c84-105">Requirements</span></span>

|<span data-ttu-id="39c84-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="39c84-106">Requirement</span></span>| <span data-ttu-id="39c84-107">Значение</span><span class="sxs-lookup"><span data-stu-id="39c84-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="39c84-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="39c84-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39c84-109">1.0</span><span class="sxs-lookup"><span data-stu-id="39c84-109">1.0</span></span>|
|[<span data-ttu-id="39c84-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="39c84-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39c84-111">Restricted</span><span class="sxs-lookup"><span data-stu-id="39c84-111">Restricted</span></span>|
|[<span data-ttu-id="39c84-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39c84-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="39c84-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="39c84-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="39c84-114">Members and methods</span></span>

| <span data-ttu-id="39c84-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="39c84-115">Member</span></span> | <span data-ttu-id="39c84-116">Тип</span><span class="sxs-lookup"><span data-stu-id="39c84-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="39c84-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="39c84-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="39c84-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="39c84-118">Member</span></span> |
| [<span data-ttu-id="39c84-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="39c84-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="39c84-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="39c84-120">Member</span></span> |
| [<span data-ttu-id="39c84-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="39c84-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="39c84-122">Метод</span><span class="sxs-lookup"><span data-stu-id="39c84-122">Method</span></span> |
| [<span data-ttu-id="39c84-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="39c84-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="39c84-124">Метод</span><span class="sxs-lookup"><span data-stu-id="39c84-124">Method</span></span> |
| [<span data-ttu-id="39c84-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="39c84-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) | <span data-ttu-id="39c84-126">Метод</span><span class="sxs-lookup"><span data-stu-id="39c84-126">Method</span></span> |
| [<span data-ttu-id="39c84-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="39c84-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="39c84-128">Метод</span><span class="sxs-lookup"><span data-stu-id="39c84-128">Method</span></span> |
| [<span data-ttu-id="39c84-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="39c84-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="39c84-130">Метод</span><span class="sxs-lookup"><span data-stu-id="39c84-130">Method</span></span> |
| [<span data-ttu-id="39c84-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="39c84-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="39c84-132">Метод</span><span class="sxs-lookup"><span data-stu-id="39c84-132">Method</span></span> |
| [<span data-ttu-id="39c84-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="39c84-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="39c84-134">Метод</span><span class="sxs-lookup"><span data-stu-id="39c84-134">Method</span></span> |
| [<span data-ttu-id="39c84-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="39c84-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="39c84-136">Метод</span><span class="sxs-lookup"><span data-stu-id="39c84-136">Method</span></span> |
| [<span data-ttu-id="39c84-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="39c84-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="39c84-138">Метод</span><span class="sxs-lookup"><span data-stu-id="39c84-138">Method</span></span> |
| [<span data-ttu-id="39c84-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="39c84-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="39c84-140">Метод</span><span class="sxs-lookup"><span data-stu-id="39c84-140">Method</span></span> |
| [<span data-ttu-id="39c84-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="39c84-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="39c84-142">Метод</span><span class="sxs-lookup"><span data-stu-id="39c84-142">Method</span></span> |
| [<span data-ttu-id="39c84-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="39c84-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="39c84-144">Метод</span><span class="sxs-lookup"><span data-stu-id="39c84-144">Method</span></span> |
| [<span data-ttu-id="39c84-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="39c84-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="39c84-146">Метод</span><span class="sxs-lookup"><span data-stu-id="39c84-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="39c84-147">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="39c84-147">Namespaces</span></span>

<span data-ttu-id="39c84-148">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="39c84-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="39c84-149">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="39c84-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="39c84-150">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="39c84-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="39c84-151">Элементы</span><span class="sxs-lookup"><span data-stu-id="39c84-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="39c84-152">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="39c84-152">ewsUrl :String</span></span>

<span data-ttu-id="39c84-p102">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="39c84-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="39c84-155">Этот член не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="39c84-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="39c84-p103">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="39c84-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="39c84-158">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="39c84-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="39c84-p104">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="39c84-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="39c84-161">Тип:</span><span class="sxs-lookup"><span data-stu-id="39c84-161">Type:</span></span>

*   <span data-ttu-id="39c84-162">String</span><span class="sxs-lookup"><span data-stu-id="39c84-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="39c84-163">Требования</span><span class="sxs-lookup"><span data-stu-id="39c84-163">Requirements</span></span>

|<span data-ttu-id="39c84-164">Requirement</span><span class="sxs-lookup"><span data-stu-id="39c84-164">Requirement</span></span>| <span data-ttu-id="39c84-165">Значение</span><span class="sxs-lookup"><span data-stu-id="39c84-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="39c84-166">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="39c84-166">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39c84-167">1.0</span><span class="sxs-lookup"><span data-stu-id="39c84-167">1.0</span></span>|
|[<span data-ttu-id="39c84-168">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="39c84-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39c84-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39c84-169">ReadItem</span></span>|
|[<span data-ttu-id="39c84-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39c84-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="39c84-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="39c84-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="39c84-172">restUrl :String</span></span>

<span data-ttu-id="39c84-173">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="39c84-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="39c84-174">С помощью значения `restUrl` можно выполнять вызовы [REST API](https://docs.microsoft.com/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="39c84-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="39c84-175">Чтобы вызвать элемент `restUrl` в режиме чтения, в манифесте приложения необходимо указать разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="39c84-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="39c84-p105">Перед использованием элемента `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="39c84-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="39c84-178">Тип:</span><span class="sxs-lookup"><span data-stu-id="39c84-178">Type:</span></span>

*   <span data-ttu-id="39c84-179">String</span><span class="sxs-lookup"><span data-stu-id="39c84-179">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="39c84-180">Требования</span><span class="sxs-lookup"><span data-stu-id="39c84-180">Requirements</span></span>

|<span data-ttu-id="39c84-181">Requirement</span><span class="sxs-lookup"><span data-stu-id="39c84-181">Requirement</span></span>| <span data-ttu-id="39c84-182">Значение</span><span class="sxs-lookup"><span data-stu-id="39c84-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="39c84-183">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="39c84-183">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39c84-184">1.5</span><span class="sxs-lookup"><span data-stu-id="39c84-184">1.5</span></span> |
|[<span data-ttu-id="39c84-185">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="39c84-185">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39c84-186">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39c84-186">ReadItem</span></span>|
|[<span data-ttu-id="39c84-187">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-187">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39c84-188">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="39c84-188">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="39c84-189">Методы</span><span class="sxs-lookup"><span data-stu-id="39c84-189">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="39c84-190">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="39c84-190">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="39c84-191">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="39c84-191">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="39c84-p106">В настоящее время поддерживаются только события типа `Office.EventType.ItemChanged`, которые вызываются, когда пользователь выбирает новый элемент. Это событие используется надстройками, реализующими закрепляемую область задач, и позволяет надстройке обновлять пользовательский интерфейс области задач в соответствии с выбранным в данный момент элементом.</span><span class="sxs-lookup"><span data-stu-id="39c84-p106">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39c84-194">Параметры</span><span class="sxs-lookup"><span data-stu-id="39c84-194">Parameters:</span></span>

| <span data-ttu-id="39c84-195">Имя</span><span class="sxs-lookup"><span data-stu-id="39c84-195">Name</span></span> | <span data-ttu-id="39c84-196">Тип</span><span class="sxs-lookup"><span data-stu-id="39c84-196">Type</span></span> | <span data-ttu-id="39c84-197">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="39c84-197">Attributes</span></span> | <span data-ttu-id="39c84-198">Описание</span><span class="sxs-lookup"><span data-stu-id="39c84-198">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="39c84-199">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="39c84-199">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="39c84-200">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="39c84-200">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="39c84-201">Function</span><span class="sxs-lookup"><span data-stu-id="39c84-201">Function</span></span> || <span data-ttu-id="39c84-p107">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="39c84-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="39c84-205">Объект</span><span class="sxs-lookup"><span data-stu-id="39c84-205">Object</span></span> | <span data-ttu-id="39c84-206">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="39c84-206">&lt;optional&gt;</span></span> | <span data-ttu-id="39c84-207">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="39c84-207">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="39c84-208">Object</span><span class="sxs-lookup"><span data-stu-id="39c84-208">Object</span></span> | <span data-ttu-id="39c84-209">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="39c84-209">&lt;optional&gt;</span></span> | <span data-ttu-id="39c84-210">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="39c84-210">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="39c84-211">function</span><span class="sxs-lookup"><span data-stu-id="39c84-211">function</span></span>| <span data-ttu-id="39c84-212">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="39c84-212">&lt;optional&gt;</span></span>|<span data-ttu-id="39c84-213">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="39c84-213">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="39c84-214">Требования</span><span class="sxs-lookup"><span data-stu-id="39c84-214">Requirements</span></span>

|<span data-ttu-id="39c84-215">Requirement</span><span class="sxs-lookup"><span data-stu-id="39c84-215">Requirement</span></span>| <span data-ttu-id="39c84-216">Значение</span><span class="sxs-lookup"><span data-stu-id="39c84-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="39c84-217">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="39c84-217">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39c84-218">1.5</span><span class="sxs-lookup"><span data-stu-id="39c84-218">1.5</span></span> |
|[<span data-ttu-id="39c84-219">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="39c84-219">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39c84-220">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39c84-220">ReadItem</span></span> |
|[<span data-ttu-id="39c84-221">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-221">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39c84-222">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="39c84-222">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="39c84-223">Пример</span><span class="sxs-lookup"><span data-stu-id="39c84-223">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="39c84-224">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="39c84-224">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="39c84-225">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="39c84-225">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="39c84-226">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="39c84-226">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="39c84-p108">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](http://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="39c84-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39c84-229">Параметры</span><span class="sxs-lookup"><span data-stu-id="39c84-229">Parameters:</span></span>

|<span data-ttu-id="39c84-230">Имя</span><span class="sxs-lookup"><span data-stu-id="39c84-230">Name</span></span>| <span data-ttu-id="39c84-231">Тип</span><span class="sxs-lookup"><span data-stu-id="39c84-231">Type</span></span>| <span data-ttu-id="39c84-232">Описание</span><span class="sxs-lookup"><span data-stu-id="39c84-232">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="39c84-233">String</span><span class="sxs-lookup"><span data-stu-id="39c84-233">String</span></span>|<span data-ttu-id="39c84-234">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-234">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="39c84-235">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="39c84-235">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="39c84-236">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="39c84-236">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="39c84-237">Требования</span><span class="sxs-lookup"><span data-stu-id="39c84-237">Requirements</span></span>

|<span data-ttu-id="39c84-238">Requirement</span><span class="sxs-lookup"><span data-stu-id="39c84-238">Requirement</span></span>| <span data-ttu-id="39c84-239">Значение</span><span class="sxs-lookup"><span data-stu-id="39c84-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="39c84-240">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="39c84-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39c84-241">1.3</span><span class="sxs-lookup"><span data-stu-id="39c84-241">1.3</span></span>|
|[<span data-ttu-id="39c84-242">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="39c84-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39c84-243">Restricted</span><span class="sxs-lookup"><span data-stu-id="39c84-243">Restricted</span></span>|
|[<span data-ttu-id="39c84-244">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39c84-245">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="39c84-245">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="39c84-246">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="39c84-246">Returns:</span></span>

<span data-ttu-id="39c84-247">Тип: String</span><span class="sxs-lookup"><span data-stu-id="39c84-247">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="39c84-248">Пример</span><span class="sxs-lookup"><span data-stu-id="39c84-248">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="39c84-249">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="39c84-249">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="39c84-250">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="39c84-250">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="39c84-p109">В случае дат и времени в почтовом приложении для Outlook или Outlook Web App могут использоваться разные часовые пояса. Outlook использует часовой пояс клиентского компьютера. Outlook Web App использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="39c84-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="39c84-p110">Если почтовое приложение работает в Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook Web App, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="39c84-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39c84-256">Параметры</span><span class="sxs-lookup"><span data-stu-id="39c84-256">Parameters:</span></span>

|<span data-ttu-id="39c84-257">Имя</span><span class="sxs-lookup"><span data-stu-id="39c84-257">Name</span></span>| <span data-ttu-id="39c84-258">Тип</span><span class="sxs-lookup"><span data-stu-id="39c84-258">Type</span></span>| <span data-ttu-id="39c84-259">Описание</span><span class="sxs-lookup"><span data-stu-id="39c84-259">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="39c84-260">Date</span><span class="sxs-lookup"><span data-stu-id="39c84-260">Date</span></span>|<span data-ttu-id="39c84-261">Объект Date</span><span class="sxs-lookup"><span data-stu-id="39c84-261">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="39c84-262">Требования</span><span class="sxs-lookup"><span data-stu-id="39c84-262">Requirements</span></span>

|<span data-ttu-id="39c84-263">Requirement</span><span class="sxs-lookup"><span data-stu-id="39c84-263">Requirement</span></span>| <span data-ttu-id="39c84-264">Значение</span><span class="sxs-lookup"><span data-stu-id="39c84-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="39c84-265">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="39c84-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39c84-266">1.0</span><span class="sxs-lookup"><span data-stu-id="39c84-266">1.0</span></span>|
|[<span data-ttu-id="39c84-267">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="39c84-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39c84-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39c84-268">ReadItem</span></span>|
|[<span data-ttu-id="39c84-269">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39c84-270">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="39c84-270">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="39c84-271">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="39c84-271">Returns:</span></span>

<span data-ttu-id="39c84-272">Тип: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="39c84-272">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="39c84-273">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="39c84-273">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="39c84-274">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="39c84-274">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="39c84-275">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="39c84-275">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="39c84-p111">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](http://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="39c84-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39c84-278">Параметры</span><span class="sxs-lookup"><span data-stu-id="39c84-278">Parameters:</span></span>

|<span data-ttu-id="39c84-279">Имя</span><span class="sxs-lookup"><span data-stu-id="39c84-279">Name</span></span>| <span data-ttu-id="39c84-280">Тип</span><span class="sxs-lookup"><span data-stu-id="39c84-280">Type</span></span>| <span data-ttu-id="39c84-281">Описание</span><span class="sxs-lookup"><span data-stu-id="39c84-281">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="39c84-282">String</span><span class="sxs-lookup"><span data-stu-id="39c84-282">String</span></span>|<span data-ttu-id="39c84-283">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="39c84-283">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="39c84-284">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="39c84-284">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="39c84-285">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="39c84-285">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="39c84-286">Требования</span><span class="sxs-lookup"><span data-stu-id="39c84-286">Requirements</span></span>

|<span data-ttu-id="39c84-287">Requirement</span><span class="sxs-lookup"><span data-stu-id="39c84-287">Requirement</span></span>| <span data-ttu-id="39c84-288">Значение</span><span class="sxs-lookup"><span data-stu-id="39c84-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="39c84-289">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="39c84-289">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39c84-290">1.3</span><span class="sxs-lookup"><span data-stu-id="39c84-290">1.3</span></span>|
|[<span data-ttu-id="39c84-291">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="39c84-291">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39c84-292">Restricted</span><span class="sxs-lookup"><span data-stu-id="39c84-292">Restricted</span></span>|
|[<span data-ttu-id="39c84-293">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-293">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39c84-294">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="39c84-294">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="39c84-295">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="39c84-295">Returns:</span></span>

<span data-ttu-id="39c84-296">Тип: String</span><span class="sxs-lookup"><span data-stu-id="39c84-296">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="39c84-297">Пример</span><span class="sxs-lookup"><span data-stu-id="39c84-297">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="39c84-298">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="39c84-298">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="39c84-299">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="39c84-299">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="39c84-300">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="39c84-300">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39c84-301">Параметры</span><span class="sxs-lookup"><span data-stu-id="39c84-301">Parameters:</span></span>

|<span data-ttu-id="39c84-302">Имя</span><span class="sxs-lookup"><span data-stu-id="39c84-302">Name</span></span>| <span data-ttu-id="39c84-303">Тип</span><span class="sxs-lookup"><span data-stu-id="39c84-303">Type</span></span>| <span data-ttu-id="39c84-304">Описание</span><span class="sxs-lookup"><span data-stu-id="39c84-304">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="39c84-305">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="39c84-305">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="39c84-306">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="39c84-306">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="39c84-307">Требования</span><span class="sxs-lookup"><span data-stu-id="39c84-307">Requirements</span></span>

|<span data-ttu-id="39c84-308">Requirement</span><span class="sxs-lookup"><span data-stu-id="39c84-308">Requirement</span></span>| <span data-ttu-id="39c84-309">Значение</span><span class="sxs-lookup"><span data-stu-id="39c84-309">Value</span></span>|
|---|---|
|[<span data-ttu-id="39c84-310">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="39c84-310">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39c84-311">1.0</span><span class="sxs-lookup"><span data-stu-id="39c84-311">1.0</span></span>|
|[<span data-ttu-id="39c84-312">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="39c84-312">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39c84-313">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39c84-313">ReadItem</span></span>|
|[<span data-ttu-id="39c84-314">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-314">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39c84-315">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="39c84-315">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="39c84-316">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="39c84-316">Returns:</span></span>

<span data-ttu-id="39c84-317">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="39c84-317">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="39c84-318">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="39c84-318">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="39c84-319">Date</span><span class="sxs-lookup"><span data-stu-id="39c84-319">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="39c84-320">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="39c84-320">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="39c84-321">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="39c84-321">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="39c84-322">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="39c84-322">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="39c84-323">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="39c84-323">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="39c84-p112">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="39c84-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="39c84-326">В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="39c84-326">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="39c84-327">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="39c84-327">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39c84-328">Параметры</span><span class="sxs-lookup"><span data-stu-id="39c84-328">Parameters:</span></span>

|<span data-ttu-id="39c84-329">Имя</span><span class="sxs-lookup"><span data-stu-id="39c84-329">Name</span></span>| <span data-ttu-id="39c84-330">Тип</span><span class="sxs-lookup"><span data-stu-id="39c84-330">Type</span></span>| <span data-ttu-id="39c84-331">Описание</span><span class="sxs-lookup"><span data-stu-id="39c84-331">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="39c84-332">String</span><span class="sxs-lookup"><span data-stu-id="39c84-332">String</span></span>|<span data-ttu-id="39c84-333">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="39c84-333">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="39c84-334">Требования</span><span class="sxs-lookup"><span data-stu-id="39c84-334">Requirements</span></span>

|<span data-ttu-id="39c84-335">Requirement</span><span class="sxs-lookup"><span data-stu-id="39c84-335">Requirement</span></span>| <span data-ttu-id="39c84-336">Значение</span><span class="sxs-lookup"><span data-stu-id="39c84-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="39c84-337">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="39c84-337">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39c84-338">1.0</span><span class="sxs-lookup"><span data-stu-id="39c84-338">1.0</span></span>|
|[<span data-ttu-id="39c84-339">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="39c84-339">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39c84-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39c84-340">ReadItem</span></span>|
|[<span data-ttu-id="39c84-341">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-341">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39c84-342">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="39c84-342">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="39c84-343">Пример</span><span class="sxs-lookup"><span data-stu-id="39c84-343">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="39c84-344">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="39c84-344">displayMessageForm(itemId)</span></span>

<span data-ttu-id="39c84-345">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="39c84-345">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="39c84-346">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="39c84-346">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="39c84-347">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="39c84-347">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="39c84-348">В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="39c84-348">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="39c84-349">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="39c84-349">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="39c84-p113">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="39c84-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39c84-352">Параметры</span><span class="sxs-lookup"><span data-stu-id="39c84-352">Parameters:</span></span>

|<span data-ttu-id="39c84-353">Имя</span><span class="sxs-lookup"><span data-stu-id="39c84-353">Name</span></span>| <span data-ttu-id="39c84-354">Тип</span><span class="sxs-lookup"><span data-stu-id="39c84-354">Type</span></span>| <span data-ttu-id="39c84-355">Описание</span><span class="sxs-lookup"><span data-stu-id="39c84-355">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="39c84-356">String</span><span class="sxs-lookup"><span data-stu-id="39c84-356">String</span></span>|<span data-ttu-id="39c84-357">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="39c84-357">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="39c84-358">Требования</span><span class="sxs-lookup"><span data-stu-id="39c84-358">Requirements</span></span>

|<span data-ttu-id="39c84-359">Requirement</span><span class="sxs-lookup"><span data-stu-id="39c84-359">Requirement</span></span>| <span data-ttu-id="39c84-360">Значение</span><span class="sxs-lookup"><span data-stu-id="39c84-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="39c84-361">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="39c84-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39c84-362">1.0</span><span class="sxs-lookup"><span data-stu-id="39c84-362">1.0</span></span>|
|[<span data-ttu-id="39c84-363">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="39c84-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39c84-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39c84-364">ReadItem</span></span>|
|[<span data-ttu-id="39c84-365">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39c84-366">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="39c84-366">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="39c84-367">Пример</span><span class="sxs-lookup"><span data-stu-id="39c84-367">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="39c84-368">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="39c84-368">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="39c84-369">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="39c84-369">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="39c84-370">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="39c84-370">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="39c84-p114">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="39c84-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="39c84-p115">В Outlook Web App и Outlook Web App для устройств этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="39c84-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="39c84-p116">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="39c84-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="39c84-378">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="39c84-378">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39c84-379">Параметры</span><span class="sxs-lookup"><span data-stu-id="39c84-379">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="39c84-380">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="39c84-380">All parameters are optional.</span></span>

|<span data-ttu-id="39c84-381">Имя</span><span class="sxs-lookup"><span data-stu-id="39c84-381">Name</span></span>| <span data-ttu-id="39c84-382">Тип</span><span class="sxs-lookup"><span data-stu-id="39c84-382">Type</span></span>| <span data-ttu-id="39c84-383">Описание</span><span class="sxs-lookup"><span data-stu-id="39c84-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="39c84-384">Object</span><span class="sxs-lookup"><span data-stu-id="39c84-384">Object</span></span> | <span data-ttu-id="39c84-385">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="39c84-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="39c84-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="39c84-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="39c84-p117">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="39c84-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="39c84-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="39c84-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="39c84-p118">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="39c84-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="39c84-392">Date</span><span class="sxs-lookup"><span data-stu-id="39c84-392">Date</span></span> | <span data-ttu-id="39c84-393">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="39c84-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="39c84-394">Date</span><span class="sxs-lookup"><span data-stu-id="39c84-394">Date</span></span> | <span data-ttu-id="39c84-395">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="39c84-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="39c84-396">String</span><span class="sxs-lookup"><span data-stu-id="39c84-396">String</span></span> | <span data-ttu-id="39c84-p119">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="39c84-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="39c84-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="39c84-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="39c84-p120">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="39c84-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="39c84-402">String</span><span class="sxs-lookup"><span data-stu-id="39c84-402">String</span></span> | <span data-ttu-id="39c84-p121">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="39c84-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="39c84-405">String</span><span class="sxs-lookup"><span data-stu-id="39c84-405">String</span></span> | <span data-ttu-id="39c84-p122">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="39c84-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="39c84-408">Требования</span><span class="sxs-lookup"><span data-stu-id="39c84-408">Requirements</span></span>

|<span data-ttu-id="39c84-409">Requirement</span><span class="sxs-lookup"><span data-stu-id="39c84-409">Requirement</span></span>| <span data-ttu-id="39c84-410">Значение</span><span class="sxs-lookup"><span data-stu-id="39c84-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="39c84-411">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="39c84-411">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39c84-412">1.0</span><span class="sxs-lookup"><span data-stu-id="39c84-412">1.0</span></span>|
|[<span data-ttu-id="39c84-413">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="39c84-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39c84-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39c84-414">ReadItem</span></span>|
|[<span data-ttu-id="39c84-415">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39c84-416">Чтение</span><span class="sxs-lookup"><span data-stu-id="39c84-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="39c84-417">Пример</span><span class="sxs-lookup"><span data-stu-id="39c84-417">Example</span></span>

```
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="39c84-418">displayNewMessageForm(параметры)</span><span class="sxs-lookup"><span data-stu-id="39c84-418">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="39c84-419">Отображает форму для создания сообщения.</span><span class="sxs-lookup"><span data-stu-id="39c84-419">Displays a form for creating a new message.</span></span>

<span data-ttu-id="39c84-420">Метод `displayNewMessageForm` открывает форму, в которой пользователь может создать сообщение.</span><span class="sxs-lookup"><span data-stu-id="39c84-420">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="39c84-421">Если параметры заданы, поля формы сообщения автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="39c84-421">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="39c84-422">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="39c84-422">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39c84-423">Параметры</span><span class="sxs-lookup"><span data-stu-id="39c84-423">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="39c84-424">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="39c84-424">All parameters are optional.</span></span>

|<span data-ttu-id="39c84-425">Имя</span><span class="sxs-lookup"><span data-stu-id="39c84-425">Name</span></span>| <span data-ttu-id="39c84-426">Тип</span><span class="sxs-lookup"><span data-stu-id="39c84-426">Type</span></span>| <span data-ttu-id="39c84-427">Описание</span><span class="sxs-lookup"><span data-stu-id="39c84-427">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="39c84-428">Объект</span><span class="sxs-lookup"><span data-stu-id="39c84-428">Object</span></span> | <span data-ttu-id="39c84-429">Словарь параметров, описывающий новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="39c84-429">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="39c84-430">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="39c84-430">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="39c84-431">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Кому".</span><span class="sxs-lookup"><span data-stu-id="39c84-431">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="39c84-432">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="39c84-432">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="39c84-433">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="39c84-433">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="39c84-434">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Копия".</span><span class="sxs-lookup"><span data-stu-id="39c84-434">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="39c84-435">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="39c84-435">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="39c84-436">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="39c84-436">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="39c84-437">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Скрытая копия".</span><span class="sxs-lookup"><span data-stu-id="39c84-437">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="39c84-438">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="39c84-438">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="39c84-439">String</span><span class="sxs-lookup"><span data-stu-id="39c84-439">String</span></span> | <span data-ttu-id="39c84-440">Строка с темой сообщения.</span><span class="sxs-lookup"><span data-stu-id="39c84-440">A string containing the subject of the message.</span></span> <span data-ttu-id="39c84-441">Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="39c84-441">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="39c84-442">String</span><span class="sxs-lookup"><span data-stu-id="39c84-442">String</span></span> | <span data-ttu-id="39c84-443">Текст сообщения в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="39c84-443">The HTML body of the message.</span></span> <span data-ttu-id="39c84-444">Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="39c84-444">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="39c84-445">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="39c84-445">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="39c84-446">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="39c84-446">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="39c84-447">String</span><span class="sxs-lookup"><span data-stu-id="39c84-447">String</span></span> | <span data-ttu-id="39c84-p129">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="39c84-p129">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="39c84-450">String</span><span class="sxs-lookup"><span data-stu-id="39c84-450">String</span></span> | <span data-ttu-id="39c84-451">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="39c84-451">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="39c84-452">String</span><span class="sxs-lookup"><span data-stu-id="39c84-452">String</span></span> | <span data-ttu-id="39c84-p130">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="39c84-p130">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="39c84-455">Boolean</span><span class="sxs-lookup"><span data-stu-id="39c84-455">Boolean</span></span> | <span data-ttu-id="39c84-p131">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="39c84-p131">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="39c84-458">String</span><span class="sxs-lookup"><span data-stu-id="39c84-458">String</span></span> | <span data-ttu-id="39c84-459">Используется, только если свойству `type` присвоено значение `item`.</span><span class="sxs-lookup"><span data-stu-id="39c84-459">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="39c84-460">Идентификатор элемента веб-служб Exchange существующего сообщения электронной почты, которые необходимо присоединить к новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="39c84-460">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="39c84-461">Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="39c84-461">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="39c84-462">Требования</span><span class="sxs-lookup"><span data-stu-id="39c84-462">Requirements</span></span>

|<span data-ttu-id="39c84-463">Requirement</span><span class="sxs-lookup"><span data-stu-id="39c84-463">Requirement</span></span>| <span data-ttu-id="39c84-464">Значение</span><span class="sxs-lookup"><span data-stu-id="39c84-464">Value</span></span>|
|---|---|
|[<span data-ttu-id="39c84-465">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="39c84-465">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39c84-466">1.6</span><span class="sxs-lookup"><span data-stu-id="39c84-466">1.6</span></span> |
|[<span data-ttu-id="39c84-467">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="39c84-467">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39c84-468">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39c84-468">ReadItem</span></span>|
|[<span data-ttu-id="39c84-469">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-469">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39c84-470">Чтение</span><span class="sxs-lookup"><span data-stu-id="39c84-470">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="39c84-471">Пример</span><span class="sxs-lookup"><span data-stu-id="39c84-471">Example</span></span>

```
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="39c84-472">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="39c84-472">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="39c84-473">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="39c84-473">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="39c84-p133">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный токен с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="39c84-p133">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="39c84-476">Рекомендуется использовать, что надстройки API-интерфейсы REST вместо веб-служб Exchange по возможности.</span><span class="sxs-lookup"><span data-stu-id="39c84-476">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="39c84-477">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="39c84-477">**REST Tokens**</span></span>

<span data-ttu-id="39c84-p134">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="39c84-p134">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="39c84-481">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="39c84-481">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="39c84-482">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="39c84-482">**EWS Tokens**</span></span>

<span data-ttu-id="39c84-p135">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="39c84-p135">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="39c84-485">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="39c84-485">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39c84-486">Параметры:</span><span class="sxs-lookup"><span data-stu-id="39c84-486">Parameters:</span></span>

|<span data-ttu-id="39c84-487">Имя</span><span class="sxs-lookup"><span data-stu-id="39c84-487">Name</span></span>| <span data-ttu-id="39c84-488">Тип</span><span class="sxs-lookup"><span data-stu-id="39c84-488">Type</span></span>| <span data-ttu-id="39c84-489">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="39c84-489">Attributes</span></span>| <span data-ttu-id="39c84-490">Описание</span><span class="sxs-lookup"><span data-stu-id="39c84-490">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="39c84-491">Объект</span><span class="sxs-lookup"><span data-stu-id="39c84-491">Object</span></span> | <span data-ttu-id="39c84-492">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="39c84-492">&lt;optional&gt;</span></span> | <span data-ttu-id="39c84-493">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="39c84-493">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="39c84-494">Boolean</span><span class="sxs-lookup"><span data-stu-id="39c84-494">Boolean</span></span> |  <span data-ttu-id="39c84-495">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="39c84-495">&lt;optional&gt;</span></span> | <span data-ttu-id="39c84-p136">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="39c84-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="39c84-498">Объект</span><span class="sxs-lookup"><span data-stu-id="39c84-498">Object</span></span> |  <span data-ttu-id="39c84-499">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="39c84-499">&lt;optional&gt;</span></span> | <span data-ttu-id="39c84-500">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="39c84-500">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="39c84-501">function</span><span class="sxs-lookup"><span data-stu-id="39c84-501">function</span></span>||<span data-ttu-id="39c84-p137">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="39c84-p137">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="39c84-504">Требования</span><span class="sxs-lookup"><span data-stu-id="39c84-504">Requirements</span></span>

|<span data-ttu-id="39c84-505">Requirement</span><span class="sxs-lookup"><span data-stu-id="39c84-505">Requirement</span></span>| <span data-ttu-id="39c84-506">Значение</span><span class="sxs-lookup"><span data-stu-id="39c84-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="39c84-507">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="39c84-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39c84-508">1.5</span><span class="sxs-lookup"><span data-stu-id="39c84-508">1.5</span></span> |
|[<span data-ttu-id="39c84-509">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="39c84-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39c84-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39c84-510">ReadItem</span></span>|
|[<span data-ttu-id="39c84-511">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39c84-512">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="39c84-512">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="39c84-513">Пример</span><span class="sxs-lookup"><span data-stu-id="39c84-513">Example</span></span>

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="39c84-514">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="39c84-514">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="39c84-515">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="39c84-515">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="39c84-p138">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный токен с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="39c84-p138">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="39c84-p139">Вы можете передать сторонней системе токен и идентификатор вложения или элемента. Сторонняя система использует этот токен как токен авторизации, чтобы вызвать операцию [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="39c84-p139">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="39c84-521">Для вызова метода `getCallbackTokenAsync` в режиме чтения манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="39c84-521">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="39c84-p140">Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="39c84-p140">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39c84-524">Параметры</span><span class="sxs-lookup"><span data-stu-id="39c84-524">Parameters:</span></span>

|<span data-ttu-id="39c84-525">Имя</span><span class="sxs-lookup"><span data-stu-id="39c84-525">Name</span></span>| <span data-ttu-id="39c84-526">Тип</span><span class="sxs-lookup"><span data-stu-id="39c84-526">Type</span></span>| <span data-ttu-id="39c84-527">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="39c84-527">Attributes</span></span>| <span data-ttu-id="39c84-528">Описание</span><span class="sxs-lookup"><span data-stu-id="39c84-528">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="39c84-529">function</span><span class="sxs-lookup"><span data-stu-id="39c84-529">function</span></span>||<span data-ttu-id="39c84-p141">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="39c84-p141">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="39c84-532">Объект</span><span class="sxs-lookup"><span data-stu-id="39c84-532">Object</span></span>| <span data-ttu-id="39c84-533">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="39c84-533">&lt;optional&gt;</span></span>|<span data-ttu-id="39c84-534">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="39c84-534">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="39c84-535">Требования</span><span class="sxs-lookup"><span data-stu-id="39c84-535">Requirements</span></span>

|<span data-ttu-id="39c84-536">Requirement</span><span class="sxs-lookup"><span data-stu-id="39c84-536">Requirement</span></span>| <span data-ttu-id="39c84-537">Значение</span><span class="sxs-lookup"><span data-stu-id="39c84-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="39c84-538">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="39c84-538">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39c84-539">1.3</span><span class="sxs-lookup"><span data-stu-id="39c84-539">1.3</span></span>|
|[<span data-ttu-id="39c84-540">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="39c84-540">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39c84-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39c84-541">ReadItem</span></span>|
|[<span data-ttu-id="39c84-542">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-542">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39c84-543">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="39c84-543">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="39c84-544">Пример</span><span class="sxs-lookup"><span data-stu-id="39c84-544">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="39c84-545">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="39c84-545">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="39c84-546">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="39c84-546">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="39c84-547">Метод `getUserIdentityTokenAsync` возвращает токен, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="39c84-547">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="39c84-548">Параметры</span><span class="sxs-lookup"><span data-stu-id="39c84-548">Parameters:</span></span>

|<span data-ttu-id="39c84-549">Имя</span><span class="sxs-lookup"><span data-stu-id="39c84-549">Name</span></span>| <span data-ttu-id="39c84-550">Тип</span><span class="sxs-lookup"><span data-stu-id="39c84-550">Type</span></span>| <span data-ttu-id="39c84-551">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="39c84-551">Attributes</span></span>| <span data-ttu-id="39c84-552">Описание</span><span class="sxs-lookup"><span data-stu-id="39c84-552">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="39c84-553">function</span><span class="sxs-lookup"><span data-stu-id="39c84-553">function</span></span>||<span data-ttu-id="39c84-554">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="39c84-554">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="39c84-555">Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="39c84-555">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="39c84-556">Object</span><span class="sxs-lookup"><span data-stu-id="39c84-556">Object</span></span>| <span data-ttu-id="39c84-557">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="39c84-557">&lt;optional&gt;</span></span>|<span data-ttu-id="39c84-558">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="39c84-558">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="39c84-559">Требования</span><span class="sxs-lookup"><span data-stu-id="39c84-559">Requirements</span></span>

|<span data-ttu-id="39c84-560">Requirement</span><span class="sxs-lookup"><span data-stu-id="39c84-560">Requirement</span></span>| <span data-ttu-id="39c84-561">Значение</span><span class="sxs-lookup"><span data-stu-id="39c84-561">Value</span></span>|
|---|---|
|[<span data-ttu-id="39c84-562">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="39c84-562">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39c84-563">1.0</span><span class="sxs-lookup"><span data-stu-id="39c84-563">1.0</span></span>|
|[<span data-ttu-id="39c84-564">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="39c84-564">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39c84-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39c84-565">ReadItem</span></span>|
|[<span data-ttu-id="39c84-566">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-566">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39c84-567">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="39c84-567">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="39c84-568">Пример</span><span class="sxs-lookup"><span data-stu-id="39c84-568">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="39c84-569">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="39c84-569">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="39c84-570">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="39c84-570">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="39c84-571">Этот метод не поддерживается в следующих случаях.</span><span class="sxs-lookup"><span data-stu-id="39c84-571">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="39c84-572">В Outlook для операций ввода-вывода или Outlook для Android (en)</span><span class="sxs-lookup"><span data-stu-id="39c84-572">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="39c84-573">Когда надстройки загружается в почтовом ящике Gmail</span><span class="sxs-lookup"><span data-stu-id="39c84-573">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="39c84-574">В таких случаях надстроек следует [использовать API -интерфейсы REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) к почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="39c84-574">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="39c84-575">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="39c84-575">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="39c84-576">[Вызов веб-служб из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) в разделе список поддерживаемые операции веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="39c84-576">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="39c84-577">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="39c84-577">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="39c84-578">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="39c84-578">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="39c84-p143">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="39c84-p143">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="39c84-581">Администратор сервера должен установить `OAuthAuthentication` имеет значение true в каталоге Client Access Server EWS, чтобы включить `makeEwsRequestAsync` запрашивает метод веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="39c84-581">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="39c84-582">Различия версий</span><span class="sxs-lookup"><span data-stu-id="39c84-582">Version differences</span></span>

<span data-ttu-id="39c84-583">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="39c84-583">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="39c84-p144">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="39c84-p144">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39c84-587">Параметры</span><span class="sxs-lookup"><span data-stu-id="39c84-587">Parameters:</span></span>

|<span data-ttu-id="39c84-588">Имя</span><span class="sxs-lookup"><span data-stu-id="39c84-588">Name</span></span>| <span data-ttu-id="39c84-589">Тип</span><span class="sxs-lookup"><span data-stu-id="39c84-589">Type</span></span>| <span data-ttu-id="39c84-590">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="39c84-590">Attributes</span></span>| <span data-ttu-id="39c84-591">Описание</span><span class="sxs-lookup"><span data-stu-id="39c84-591">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="39c84-592">String</span><span class="sxs-lookup"><span data-stu-id="39c84-592">String</span></span>||<span data-ttu-id="39c84-593">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="39c84-593">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="39c84-594">function</span><span class="sxs-lookup"><span data-stu-id="39c84-594">function</span></span>||<span data-ttu-id="39c84-595">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="39c84-595">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="39c84-596">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="39c84-596">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="39c84-597">Если результат превышает 1 МБ по размеру, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="39c84-597">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="39c84-598">Объект</span><span class="sxs-lookup"><span data-stu-id="39c84-598">Object</span></span>| <span data-ttu-id="39c84-599">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="39c84-599">&lt;optional&gt;</span></span>|<span data-ttu-id="39c84-600">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="39c84-600">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="39c84-601">Требования</span><span class="sxs-lookup"><span data-stu-id="39c84-601">Requirements</span></span>

|<span data-ttu-id="39c84-602">Requirement</span><span class="sxs-lookup"><span data-stu-id="39c84-602">Requirement</span></span>| <span data-ttu-id="39c84-603">Значение</span><span class="sxs-lookup"><span data-stu-id="39c84-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="39c84-604">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="39c84-604">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39c84-605">1.0</span><span class="sxs-lookup"><span data-stu-id="39c84-605">1.0</span></span>|
|[<span data-ttu-id="39c84-606">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="39c84-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39c84-607">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="39c84-607">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="39c84-608">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="39c84-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39c84-609">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="39c84-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="39c84-610">Пример</span><span class="sxs-lookup"><span data-stu-id="39c84-610">Example</span></span>

<span data-ttu-id="39c84-611">В следующем примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="39c84-611">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item.
   var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
}
```