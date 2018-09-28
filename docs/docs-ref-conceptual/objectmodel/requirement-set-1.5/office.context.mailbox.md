
# <a name="mailbox"></a><span data-ttu-id="ef010-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="ef010-101">mailbox</span></span>

### <span data-ttu-id="ef010-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="ef010-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="ef010-104">Предоставляет доступ для добавления в объектной модели Outlook для Microsoft Outlook и Microsoft Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="ef010-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ef010-105">Требования</span><span class="sxs-lookup"><span data-stu-id="ef010-105">Requirements</span></span>

|<span data-ttu-id="ef010-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="ef010-106">Requirement</span></span>| <span data-ttu-id="ef010-107">Значение</span><span class="sxs-lookup"><span data-stu-id="ef010-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef010-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ef010-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef010-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ef010-109">1.0</span></span>|
|[<span data-ttu-id="ef010-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ef010-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef010-111">Restricted</span><span class="sxs-lookup"><span data-stu-id="ef010-111">Restricted</span></span>|
|[<span data-ttu-id="ef010-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ef010-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ef010-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ef010-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ef010-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="ef010-114">Members and methods</span></span>

| <span data-ttu-id="ef010-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="ef010-115">Member</span></span> | <span data-ttu-id="ef010-116">Тип</span><span class="sxs-lookup"><span data-stu-id="ef010-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ef010-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="ef010-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="ef010-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="ef010-118">Member</span></span> |
| [<span data-ttu-id="ef010-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="ef010-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="ef010-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="ef010-120">Member</span></span> |
| [<span data-ttu-id="ef010-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="ef010-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="ef010-122">Метод</span><span class="sxs-lookup"><span data-stu-id="ef010-122">Method</span></span> |
| [<span data-ttu-id="ef010-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="ef010-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="ef010-124">Метод</span><span class="sxs-lookup"><span data-stu-id="ef010-124">Method</span></span> |
| [<span data-ttu-id="ef010-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="ef010-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) | <span data-ttu-id="ef010-126">Метод</span><span class="sxs-lookup"><span data-stu-id="ef010-126">Method</span></span> |
| [<span data-ttu-id="ef010-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="ef010-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="ef010-128">Метод</span><span class="sxs-lookup"><span data-stu-id="ef010-128">Method</span></span> |
| [<span data-ttu-id="ef010-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="ef010-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="ef010-130">Метод</span><span class="sxs-lookup"><span data-stu-id="ef010-130">Method</span></span> |
| [<span data-ttu-id="ef010-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="ef010-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="ef010-132">Метод</span><span class="sxs-lookup"><span data-stu-id="ef010-132">Method</span></span> |
| [<span data-ttu-id="ef010-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="ef010-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="ef010-134">Метод</span><span class="sxs-lookup"><span data-stu-id="ef010-134">Method</span></span> |
| [<span data-ttu-id="ef010-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="ef010-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="ef010-136">Метод</span><span class="sxs-lookup"><span data-stu-id="ef010-136">Method</span></span> |
| [<span data-ttu-id="ef010-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="ef010-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="ef010-138">Метод</span><span class="sxs-lookup"><span data-stu-id="ef010-138">Method</span></span> |
| [<span data-ttu-id="ef010-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="ef010-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="ef010-140">Метод</span><span class="sxs-lookup"><span data-stu-id="ef010-140">Method</span></span> |
| [<span data-ttu-id="ef010-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="ef010-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="ef010-142">Метод</span><span class="sxs-lookup"><span data-stu-id="ef010-142">Method</span></span> |
| [<span data-ttu-id="ef010-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="ef010-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="ef010-144">Метод</span><span class="sxs-lookup"><span data-stu-id="ef010-144">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="ef010-145">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="ef010-145">Namespaces</span></span>

<span data-ttu-id="ef010-146">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="ef010-146">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="ef010-147">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="ef010-147">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="ef010-148">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="ef010-148">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="ef010-149">Элементы</span><span class="sxs-lookup"><span data-stu-id="ef010-149">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="ef010-150">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="ef010-150">ewsUrl :String</span></span>

<span data-ttu-id="ef010-p102">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ef010-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ef010-153">Этот член не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="ef010-153">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ef010-p103">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="ef010-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="ef010-156">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="ef010-156">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="ef010-p104">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="ef010-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="ef010-159">Тип:</span><span class="sxs-lookup"><span data-stu-id="ef010-159">Type:</span></span>

*   <span data-ttu-id="ef010-160">String</span><span class="sxs-lookup"><span data-stu-id="ef010-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ef010-161">Требования</span><span class="sxs-lookup"><span data-stu-id="ef010-161">Requirements</span></span>

|<span data-ttu-id="ef010-162">Requirement</span><span class="sxs-lookup"><span data-stu-id="ef010-162">Requirement</span></span>| <span data-ttu-id="ef010-163">Значение</span><span class="sxs-lookup"><span data-stu-id="ef010-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef010-164">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ef010-164">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef010-165">1.0</span><span class="sxs-lookup"><span data-stu-id="ef010-165">1.0</span></span>|
|[<span data-ttu-id="ef010-166">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ef010-166">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef010-167">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ef010-167">ReadItem</span></span>|
|[<span data-ttu-id="ef010-168">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ef010-168">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ef010-169">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ef010-169">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="ef010-170">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="ef010-170">restUrl :String</span></span>

<span data-ttu-id="ef010-171">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="ef010-171">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="ef010-172">С помощью значения `restUrl` можно выполнять вызовы [REST API](https://docs.microsoft.com/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="ef010-172">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="ef010-173">Чтобы вызвать элемент `restUrl` в режиме чтения, в манифесте приложения необходимо указать разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="ef010-173">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="ef010-p105">Перед использованием элемента `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="ef010-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="ef010-176">Возвращает недопустимое значение для клиентов Outlook на локальных установок Exchange 2016 связи или более поздняя версия настраиваемых REST URL-адрес, настроенных `restUrl`.</span><span class="sxs-lookup"><span data-stu-id="ef010-176">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="ef010-177">Тип:</span><span class="sxs-lookup"><span data-stu-id="ef010-177">Type:</span></span>

*   <span data-ttu-id="ef010-178">String</span><span class="sxs-lookup"><span data-stu-id="ef010-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ef010-179">Требования</span><span class="sxs-lookup"><span data-stu-id="ef010-179">Requirements</span></span>

|<span data-ttu-id="ef010-180">Requirement</span><span class="sxs-lookup"><span data-stu-id="ef010-180">Requirement</span></span>| <span data-ttu-id="ef010-181">Значение</span><span class="sxs-lookup"><span data-stu-id="ef010-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef010-182">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ef010-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef010-183">1.5</span><span class="sxs-lookup"><span data-stu-id="ef010-183">1.5</span></span> |
|[<span data-ttu-id="ef010-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ef010-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef010-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ef010-185">ReadItem</span></span>|
|[<span data-ttu-id="ef010-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ef010-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ef010-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ef010-187">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="ef010-188">Методы</span><span class="sxs-lookup"><span data-stu-id="ef010-188">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="ef010-189">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ef010-189">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="ef010-190">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="ef010-190">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="ef010-p106">В настоящее время поддерживаются только события типа `Office.EventType.ItemChanged`, которые вызываются, когда пользователь выбирает новый элемент. Это событие используется надстройками, реализующими закрепляемую область задач, и позволяет надстройке обновлять пользовательский интерфейс области задач в соответствии с выбранным в данный момент элементом.</span><span class="sxs-lookup"><span data-stu-id="ef010-p106">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ef010-193">Параметры</span><span class="sxs-lookup"><span data-stu-id="ef010-193">Parameters:</span></span>

| <span data-ttu-id="ef010-194">Имя</span><span class="sxs-lookup"><span data-stu-id="ef010-194">Name</span></span> | <span data-ttu-id="ef010-195">Тип</span><span class="sxs-lookup"><span data-stu-id="ef010-195">Type</span></span> | <span data-ttu-id="ef010-196">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ef010-196">Attributes</span></span> | <span data-ttu-id="ef010-197">Описание</span><span class="sxs-lookup"><span data-stu-id="ef010-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="ef010-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="ef010-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="ef010-199">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="ef010-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="ef010-200">Function</span><span class="sxs-lookup"><span data-stu-id="ef010-200">Function</span></span> || <span data-ttu-id="ef010-p107">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="ef010-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="ef010-204">Объект</span><span class="sxs-lookup"><span data-stu-id="ef010-204">Object</span></span> | <span data-ttu-id="ef010-205">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ef010-205">&lt;optional&gt;</span></span> | <span data-ttu-id="ef010-206">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="ef010-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="ef010-207">Object</span><span class="sxs-lookup"><span data-stu-id="ef010-207">Object</span></span> | <span data-ttu-id="ef010-208">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ef010-208">&lt;optional&gt;</span></span> | <span data-ttu-id="ef010-209">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ef010-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="ef010-210">function</span><span class="sxs-lookup"><span data-stu-id="ef010-210">function</span></span>| <span data-ttu-id="ef010-211">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ef010-211">&lt;optional&gt;</span></span>|<span data-ttu-id="ef010-212">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ef010-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ef010-213">Требования</span><span class="sxs-lookup"><span data-stu-id="ef010-213">Requirements</span></span>

|<span data-ttu-id="ef010-214">Requirement</span><span class="sxs-lookup"><span data-stu-id="ef010-214">Requirement</span></span>| <span data-ttu-id="ef010-215">Значение</span><span class="sxs-lookup"><span data-stu-id="ef010-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef010-216">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ef010-216">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef010-217">1.5</span><span class="sxs-lookup"><span data-stu-id="ef010-217">1.5</span></span> |
|[<span data-ttu-id="ef010-218">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ef010-218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef010-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ef010-219">ReadItem</span></span> |
|[<span data-ttu-id="ef010-220">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ef010-220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ef010-221">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ef010-221">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ef010-222">Пример</span><span class="sxs-lookup"><span data-stu-id="ef010-222">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="ef010-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="ef010-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="ef010-224">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="ef010-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="ef010-225">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="ef010-225">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ef010-p108">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](http://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="ef010-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ef010-228">Параметры</span><span class="sxs-lookup"><span data-stu-id="ef010-228">Parameters:</span></span>

|<span data-ttu-id="ef010-229">Имя</span><span class="sxs-lookup"><span data-stu-id="ef010-229">Name</span></span>| <span data-ttu-id="ef010-230">Тип</span><span class="sxs-lookup"><span data-stu-id="ef010-230">Type</span></span>| <span data-ttu-id="ef010-231">Описание</span><span class="sxs-lookup"><span data-stu-id="ef010-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="ef010-232">String</span><span class="sxs-lookup"><span data-stu-id="ef010-232">String</span></span>|<span data-ttu-id="ef010-233">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="ef010-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="ef010-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="ef010-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="ef010-235">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="ef010-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ef010-236">Требования</span><span class="sxs-lookup"><span data-stu-id="ef010-236">Requirements</span></span>

|<span data-ttu-id="ef010-237">Requirement</span><span class="sxs-lookup"><span data-stu-id="ef010-237">Requirement</span></span>| <span data-ttu-id="ef010-238">Значение</span><span class="sxs-lookup"><span data-stu-id="ef010-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef010-239">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ef010-239">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef010-240">1.3</span><span class="sxs-lookup"><span data-stu-id="ef010-240">1.3</span></span>|
|[<span data-ttu-id="ef010-241">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ef010-241">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef010-242">Restricted</span><span class="sxs-lookup"><span data-stu-id="ef010-242">Restricted</span></span>|
|[<span data-ttu-id="ef010-243">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ef010-243">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ef010-244">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ef010-244">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ef010-245">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ef010-245">Returns:</span></span>

<span data-ttu-id="ef010-246">Тип: String</span><span class="sxs-lookup"><span data-stu-id="ef010-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="ef010-247">Пример</span><span class="sxs-lookup"><span data-stu-id="ef010-247">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime"></a><span data-ttu-id="ef010-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="ef010-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span></span>

<span data-ttu-id="ef010-249">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="ef010-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="ef010-p109">В случае дат и времени в почтовом приложении для Outlook или Outlook Web App могут использоваться разные часовые пояса. Outlook использует часовой пояс клиентского компьютера. Outlook Web App использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="ef010-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="ef010-p110">Если почтовое приложение работает в Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook Web App, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="ef010-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ef010-255">Параметры</span><span class="sxs-lookup"><span data-stu-id="ef010-255">Parameters:</span></span>

|<span data-ttu-id="ef010-256">Имя</span><span class="sxs-lookup"><span data-stu-id="ef010-256">Name</span></span>| <span data-ttu-id="ef010-257">Тип</span><span class="sxs-lookup"><span data-stu-id="ef010-257">Type</span></span>| <span data-ttu-id="ef010-258">Описание</span><span class="sxs-lookup"><span data-stu-id="ef010-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="ef010-259">Date</span><span class="sxs-lookup"><span data-stu-id="ef010-259">Date</span></span>|<span data-ttu-id="ef010-260">Объект Date</span><span class="sxs-lookup"><span data-stu-id="ef010-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ef010-261">Требования</span><span class="sxs-lookup"><span data-stu-id="ef010-261">Requirements</span></span>

|<span data-ttu-id="ef010-262">Requirement</span><span class="sxs-lookup"><span data-stu-id="ef010-262">Requirement</span></span>| <span data-ttu-id="ef010-263">Значение</span><span class="sxs-lookup"><span data-stu-id="ef010-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef010-264">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ef010-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef010-265">1.0</span><span class="sxs-lookup"><span data-stu-id="ef010-265">1.0</span></span>|
|[<span data-ttu-id="ef010-266">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ef010-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef010-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ef010-267">ReadItem</span></span>|
|[<span data-ttu-id="ef010-268">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ef010-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ef010-269">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ef010-269">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ef010-270">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ef010-270">Returns:</span></span>

<span data-ttu-id="ef010-271">Тип: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="ef010-271">Type: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="ef010-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="ef010-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="ef010-273">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="ef010-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="ef010-274">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="ef010-274">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ef010-p111">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](http://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="ef010-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ef010-277">Параметры</span><span class="sxs-lookup"><span data-stu-id="ef010-277">Parameters:</span></span>

|<span data-ttu-id="ef010-278">Имя</span><span class="sxs-lookup"><span data-stu-id="ef010-278">Name</span></span>| <span data-ttu-id="ef010-279">Тип</span><span class="sxs-lookup"><span data-stu-id="ef010-279">Type</span></span>| <span data-ttu-id="ef010-280">Описание</span><span class="sxs-lookup"><span data-stu-id="ef010-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="ef010-281">String</span><span class="sxs-lookup"><span data-stu-id="ef010-281">String</span></span>|<span data-ttu-id="ef010-282">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="ef010-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="ef010-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="ef010-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="ef010-284">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="ef010-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ef010-285">Требования</span><span class="sxs-lookup"><span data-stu-id="ef010-285">Requirements</span></span>

|<span data-ttu-id="ef010-286">Requirement</span><span class="sxs-lookup"><span data-stu-id="ef010-286">Requirement</span></span>| <span data-ttu-id="ef010-287">Значение</span><span class="sxs-lookup"><span data-stu-id="ef010-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef010-288">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ef010-288">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef010-289">1.3</span><span class="sxs-lookup"><span data-stu-id="ef010-289">1.3</span></span>|
|[<span data-ttu-id="ef010-290">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ef010-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef010-291">Restricted</span><span class="sxs-lookup"><span data-stu-id="ef010-291">Restricted</span></span>|
|[<span data-ttu-id="ef010-292">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ef010-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ef010-293">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ef010-293">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ef010-294">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ef010-294">Returns:</span></span>

<span data-ttu-id="ef010-295">Тип: String</span><span class="sxs-lookup"><span data-stu-id="ef010-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="ef010-296">Пример</span><span class="sxs-lookup"><span data-stu-id="ef010-296">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="ef010-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="ef010-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="ef010-298">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="ef010-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="ef010-299">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="ef010-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ef010-300">Параметры</span><span class="sxs-lookup"><span data-stu-id="ef010-300">Parameters:</span></span>

|<span data-ttu-id="ef010-301">Имя</span><span class="sxs-lookup"><span data-stu-id="ef010-301">Name</span></span>| <span data-ttu-id="ef010-302">Тип</span><span class="sxs-lookup"><span data-stu-id="ef010-302">Type</span></span>| <span data-ttu-id="ef010-303">Описание</span><span class="sxs-lookup"><span data-stu-id="ef010-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="ef010-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="ef010-304">LocalClientTime</span></span>](/javascript/api/outlook_1_5/office.LocalClientTime)|<span data-ttu-id="ef010-305">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="ef010-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ef010-306">Требования</span><span class="sxs-lookup"><span data-stu-id="ef010-306">Requirements</span></span>

|<span data-ttu-id="ef010-307">Requirement</span><span class="sxs-lookup"><span data-stu-id="ef010-307">Requirement</span></span>| <span data-ttu-id="ef010-308">Значение</span><span class="sxs-lookup"><span data-stu-id="ef010-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef010-309">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ef010-309">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef010-310">1.0</span><span class="sxs-lookup"><span data-stu-id="ef010-310">1.0</span></span>|
|[<span data-ttu-id="ef010-311">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ef010-311">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef010-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ef010-312">ReadItem</span></span>|
|[<span data-ttu-id="ef010-313">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ef010-313">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ef010-314">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ef010-314">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ef010-315">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ef010-315">Returns:</span></span>

<span data-ttu-id="ef010-316">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="ef010-316">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="ef010-317">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="ef010-317">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ef010-318">Date</span><span class="sxs-lookup"><span data-stu-id="ef010-318">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="ef010-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="ef010-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="ef010-320">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="ef010-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ef010-321">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="ef010-321">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ef010-322">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="ef010-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="ef010-p112">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="ef010-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="ef010-325">В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="ef010-325">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="ef010-326">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="ef010-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ef010-327">Параметры</span><span class="sxs-lookup"><span data-stu-id="ef010-327">Parameters:</span></span>

|<span data-ttu-id="ef010-328">Имя</span><span class="sxs-lookup"><span data-stu-id="ef010-328">Name</span></span>| <span data-ttu-id="ef010-329">Тип</span><span class="sxs-lookup"><span data-stu-id="ef010-329">Type</span></span>| <span data-ttu-id="ef010-330">Описание</span><span class="sxs-lookup"><span data-stu-id="ef010-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="ef010-331">String</span><span class="sxs-lookup"><span data-stu-id="ef010-331">String</span></span>|<span data-ttu-id="ef010-332">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="ef010-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ef010-333">Требования</span><span class="sxs-lookup"><span data-stu-id="ef010-333">Requirements</span></span>

|<span data-ttu-id="ef010-334">Requirement</span><span class="sxs-lookup"><span data-stu-id="ef010-334">Requirement</span></span>| <span data-ttu-id="ef010-335">Значение</span><span class="sxs-lookup"><span data-stu-id="ef010-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef010-336">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ef010-336">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef010-337">1.0</span><span class="sxs-lookup"><span data-stu-id="ef010-337">1.0</span></span>|
|[<span data-ttu-id="ef010-338">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ef010-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef010-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ef010-339">ReadItem</span></span>|
|[<span data-ttu-id="ef010-340">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ef010-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ef010-341">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ef010-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ef010-342">Пример</span><span class="sxs-lookup"><span data-stu-id="ef010-342">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="ef010-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="ef010-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="ef010-344">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="ef010-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="ef010-345">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="ef010-345">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ef010-346">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="ef010-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="ef010-347">В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="ef010-347">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="ef010-348">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="ef010-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="ef010-p113">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="ef010-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ef010-351">Параметры</span><span class="sxs-lookup"><span data-stu-id="ef010-351">Parameters:</span></span>

|<span data-ttu-id="ef010-352">Имя</span><span class="sxs-lookup"><span data-stu-id="ef010-352">Name</span></span>| <span data-ttu-id="ef010-353">Тип</span><span class="sxs-lookup"><span data-stu-id="ef010-353">Type</span></span>| <span data-ttu-id="ef010-354">Описание</span><span class="sxs-lookup"><span data-stu-id="ef010-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="ef010-355">String</span><span class="sxs-lookup"><span data-stu-id="ef010-355">String</span></span>|<span data-ttu-id="ef010-356">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="ef010-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ef010-357">Требования</span><span class="sxs-lookup"><span data-stu-id="ef010-357">Requirements</span></span>

|<span data-ttu-id="ef010-358">Requirement</span><span class="sxs-lookup"><span data-stu-id="ef010-358">Requirement</span></span>| <span data-ttu-id="ef010-359">Значение</span><span class="sxs-lookup"><span data-stu-id="ef010-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef010-360">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ef010-360">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef010-361">1.0</span><span class="sxs-lookup"><span data-stu-id="ef010-361">1.0</span></span>|
|[<span data-ttu-id="ef010-362">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ef010-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef010-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ef010-363">ReadItem</span></span>|
|[<span data-ttu-id="ef010-364">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ef010-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ef010-365">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ef010-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ef010-366">Пример</span><span class="sxs-lookup"><span data-stu-id="ef010-366">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="ef010-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="ef010-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="ef010-368">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="ef010-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ef010-369">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="ef010-369">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ef010-p114">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="ef010-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="ef010-p115">В Outlook Web App и Outlook Web App для устройств этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="ef010-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="ef010-p116">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="ef010-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="ef010-377">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="ef010-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ef010-378">Параметры</span><span class="sxs-lookup"><span data-stu-id="ef010-378">Parameters:</span></span>

|<span data-ttu-id="ef010-379">Имя</span><span class="sxs-lookup"><span data-stu-id="ef010-379">Name</span></span>| <span data-ttu-id="ef010-380">Тип</span><span class="sxs-lookup"><span data-stu-id="ef010-380">Type</span></span>| <span data-ttu-id="ef010-381">Описание</span><span class="sxs-lookup"><span data-stu-id="ef010-381">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="ef010-382">Object</span><span class="sxs-lookup"><span data-stu-id="ef010-382">Object</span></span> | <span data-ttu-id="ef010-383">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="ef010-383">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="ef010-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="ef010-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="ef010-p117">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="ef010-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="ef010-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="ef010-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="ef010-p118">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="ef010-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="ef010-390">Date</span><span class="sxs-lookup"><span data-stu-id="ef010-390">Date</span></span> | <span data-ttu-id="ef010-391">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="ef010-391">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="ef010-392">Date</span><span class="sxs-lookup"><span data-stu-id="ef010-392">Date</span></span> | <span data-ttu-id="ef010-393">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="ef010-393">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="ef010-394">String</span><span class="sxs-lookup"><span data-stu-id="ef010-394">String</span></span> | <span data-ttu-id="ef010-p119">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="ef010-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="ef010-397">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="ef010-397">Array.&lt;String&gt;</span></span> | <span data-ttu-id="ef010-p120">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="ef010-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="ef010-400">String</span><span class="sxs-lookup"><span data-stu-id="ef010-400">String</span></span> | <span data-ttu-id="ef010-p121">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="ef010-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="ef010-403">String</span><span class="sxs-lookup"><span data-stu-id="ef010-403">String</span></span> | <span data-ttu-id="ef010-p122">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="ef010-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ef010-406">Требования</span><span class="sxs-lookup"><span data-stu-id="ef010-406">Requirements</span></span>

|<span data-ttu-id="ef010-407">Requirement</span><span class="sxs-lookup"><span data-stu-id="ef010-407">Requirement</span></span>| <span data-ttu-id="ef010-408">Значение</span><span class="sxs-lookup"><span data-stu-id="ef010-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef010-409">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ef010-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef010-410">1.0</span><span class="sxs-lookup"><span data-stu-id="ef010-410">1.0</span></span>|
|[<span data-ttu-id="ef010-411">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ef010-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef010-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ef010-412">ReadItem</span></span>|
|[<span data-ttu-id="ef010-413">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ef010-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ef010-414">Чтение</span><span class="sxs-lookup"><span data-stu-id="ef010-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ef010-415">Пример</span><span class="sxs-lookup"><span data-stu-id="ef010-415">Example</span></span>

```js
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="ef010-416">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="ef010-416">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="ef010-417">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="ef010-417">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="ef010-p123">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный токен с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="ef010-p123">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="ef010-420">Рекомендуется использовать, что надстройки API-интерфейсы REST вместо веб-служб Exchange по возможности.</span><span class="sxs-lookup"><span data-stu-id="ef010-420">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="ef010-421">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="ef010-421">**REST Tokens**</span></span>

<span data-ttu-id="ef010-p124">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="ef010-p124">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="ef010-425">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="ef010-425">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="ef010-426">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="ef010-426">**EWS Tokens**</span></span>

<span data-ttu-id="ef010-p125">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="ef010-p125">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="ef010-429">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="ef010-429">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ef010-430">Параметры:</span><span class="sxs-lookup"><span data-stu-id="ef010-430">Parameters:</span></span>

|<span data-ttu-id="ef010-431">Имя</span><span class="sxs-lookup"><span data-stu-id="ef010-431">Name</span></span>| <span data-ttu-id="ef010-432">Тип</span><span class="sxs-lookup"><span data-stu-id="ef010-432">Type</span></span>| <span data-ttu-id="ef010-433">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ef010-433">Attributes</span></span>| <span data-ttu-id="ef010-434">Описание</span><span class="sxs-lookup"><span data-stu-id="ef010-434">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="ef010-435">Объект</span><span class="sxs-lookup"><span data-stu-id="ef010-435">Object</span></span> | <span data-ttu-id="ef010-436">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ef010-436">&lt;optional&gt;</span></span> | <span data-ttu-id="ef010-437">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="ef010-437">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="ef010-438">Boolean</span><span class="sxs-lookup"><span data-stu-id="ef010-438">Boolean</span></span> |  <span data-ttu-id="ef010-439">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ef010-439">&lt;optional&gt;</span></span> | <span data-ttu-id="ef010-p126">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="ef010-p126">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="ef010-442">Объект</span><span class="sxs-lookup"><span data-stu-id="ef010-442">Object</span></span> |  <span data-ttu-id="ef010-443">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ef010-443">&lt;optional&gt;</span></span> | <span data-ttu-id="ef010-444">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="ef010-444">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="ef010-445">function</span><span class="sxs-lookup"><span data-stu-id="ef010-445">function</span></span>||<span data-ttu-id="ef010-p127">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ef010-p127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ef010-448">Требования</span><span class="sxs-lookup"><span data-stu-id="ef010-448">Requirements</span></span>

|<span data-ttu-id="ef010-449">Requirement</span><span class="sxs-lookup"><span data-stu-id="ef010-449">Requirement</span></span>| <span data-ttu-id="ef010-450">Значение</span><span class="sxs-lookup"><span data-stu-id="ef010-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef010-451">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ef010-451">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef010-452">1.5</span><span class="sxs-lookup"><span data-stu-id="ef010-452">1.5</span></span> |
|[<span data-ttu-id="ef010-453">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ef010-453">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef010-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ef010-454">ReadItem</span></span>|
|[<span data-ttu-id="ef010-455">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ef010-455">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ef010-456">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="ef010-456">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="ef010-457">Пример</span><span class="sxs-lookup"><span data-stu-id="ef010-457">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="ef010-458">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ef010-458">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="ef010-459">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="ef010-459">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="ef010-p128">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный токен с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="ef010-p128">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="ef010-p129">Вы можете передать сторонней системе токен и идентификатор вложения или элемента. Сторонняя система использует этот токен как токен авторизации, чтобы вызвать операцию [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="ef010-p129">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="ef010-465">Для вызова метода `getCallbackTokenAsync` в режиме чтения манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="ef010-465">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="ef010-p130">Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="ef010-p130">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ef010-468">Параметры</span><span class="sxs-lookup"><span data-stu-id="ef010-468">Parameters:</span></span>

|<span data-ttu-id="ef010-469">Имя</span><span class="sxs-lookup"><span data-stu-id="ef010-469">Name</span></span>| <span data-ttu-id="ef010-470">Тип</span><span class="sxs-lookup"><span data-stu-id="ef010-470">Type</span></span>| <span data-ttu-id="ef010-471">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ef010-471">Attributes</span></span>| <span data-ttu-id="ef010-472">Описание</span><span class="sxs-lookup"><span data-stu-id="ef010-472">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="ef010-473">function</span><span class="sxs-lookup"><span data-stu-id="ef010-473">function</span></span>||<span data-ttu-id="ef010-p131">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ef010-p131">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="ef010-476">Объект</span><span class="sxs-lookup"><span data-stu-id="ef010-476">Object</span></span>| <span data-ttu-id="ef010-477">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ef010-477">&lt;optional&gt;</span></span>|<span data-ttu-id="ef010-478">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="ef010-478">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ef010-479">Требования</span><span class="sxs-lookup"><span data-stu-id="ef010-479">Requirements</span></span>

|<span data-ttu-id="ef010-480">Requirement</span><span class="sxs-lookup"><span data-stu-id="ef010-480">Requirement</span></span>| <span data-ttu-id="ef010-481">Значение</span><span class="sxs-lookup"><span data-stu-id="ef010-481">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef010-482">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ef010-482">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef010-483">1.3</span><span class="sxs-lookup"><span data-stu-id="ef010-483">1.3</span></span>|
|[<span data-ttu-id="ef010-484">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ef010-484">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef010-485">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ef010-485">ReadItem</span></span>|
|[<span data-ttu-id="ef010-486">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ef010-486">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ef010-487">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="ef010-487">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="ef010-488">Пример</span><span class="sxs-lookup"><span data-stu-id="ef010-488">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="ef010-489">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ef010-489">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="ef010-490">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="ef010-490">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="ef010-491">Метод `getUserIdentityTokenAsync` возвращает токен, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="ef010-491">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="ef010-492">Параметры</span><span class="sxs-lookup"><span data-stu-id="ef010-492">Parameters:</span></span>

|<span data-ttu-id="ef010-493">Имя</span><span class="sxs-lookup"><span data-stu-id="ef010-493">Name</span></span>| <span data-ttu-id="ef010-494">Тип</span><span class="sxs-lookup"><span data-stu-id="ef010-494">Type</span></span>| <span data-ttu-id="ef010-495">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ef010-495">Attributes</span></span>| <span data-ttu-id="ef010-496">Описание</span><span class="sxs-lookup"><span data-stu-id="ef010-496">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="ef010-497">function</span><span class="sxs-lookup"><span data-stu-id="ef010-497">function</span></span>||<span data-ttu-id="ef010-498">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ef010-498">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ef010-499">Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ef010-499">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="ef010-500">Object</span><span class="sxs-lookup"><span data-stu-id="ef010-500">Object</span></span>| <span data-ttu-id="ef010-501">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ef010-501">&lt;optional&gt;</span></span>|<span data-ttu-id="ef010-502">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="ef010-502">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ef010-503">Требования</span><span class="sxs-lookup"><span data-stu-id="ef010-503">Requirements</span></span>

|<span data-ttu-id="ef010-504">Requirement</span><span class="sxs-lookup"><span data-stu-id="ef010-504">Requirement</span></span>| <span data-ttu-id="ef010-505">Значение</span><span class="sxs-lookup"><span data-stu-id="ef010-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef010-506">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ef010-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef010-507">1.0</span><span class="sxs-lookup"><span data-stu-id="ef010-507">1.0</span></span>|
|[<span data-ttu-id="ef010-508">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ef010-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef010-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ef010-509">ReadItem</span></span>|
|[<span data-ttu-id="ef010-510">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ef010-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ef010-511">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ef010-511">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ef010-512">Пример</span><span class="sxs-lookup"><span data-stu-id="ef010-512">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="ef010-513">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ef010-513">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="ef010-514">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="ef010-514">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="ef010-515">Этот метод не поддерживается в следующих случаях.</span><span class="sxs-lookup"><span data-stu-id="ef010-515">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="ef010-516">В Outlook для операций ввода-вывода или Outlook для Android (en)</span><span class="sxs-lookup"><span data-stu-id="ef010-516">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="ef010-517">Когда надстройки загружается в почтовом ящике Gmail</span><span class="sxs-lookup"><span data-stu-id="ef010-517">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="ef010-518">В таких случаях надстроек следует [использовать API -интерфейсы REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) к почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="ef010-518">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="ef010-519">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="ef010-519">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="ef010-520">[Вызов веб-служб из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) в разделе список поддерживаемые операции веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="ef010-520">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="ef010-521">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="ef010-521">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="ef010-522">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="ef010-522">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="ef010-p133">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="ef010-p133">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="ef010-525">Администратор сервера должен установить `OAuthAuthentication` имеет значение true в каталоге Client Access Server EWS, чтобы включить `makeEwsRequestAsync` запрашивает метод веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="ef010-525">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="ef010-526">Различия версий</span><span class="sxs-lookup"><span data-stu-id="ef010-526">Version differences</span></span>

<span data-ttu-id="ef010-527">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="ef010-527">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="ef010-p134">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="ef010-p134">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ef010-531">Параметры</span><span class="sxs-lookup"><span data-stu-id="ef010-531">Parameters:</span></span>

|<span data-ttu-id="ef010-532">Имя</span><span class="sxs-lookup"><span data-stu-id="ef010-532">Name</span></span>| <span data-ttu-id="ef010-533">Тип</span><span class="sxs-lookup"><span data-stu-id="ef010-533">Type</span></span>| <span data-ttu-id="ef010-534">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ef010-534">Attributes</span></span>| <span data-ttu-id="ef010-535">Описание</span><span class="sxs-lookup"><span data-stu-id="ef010-535">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="ef010-536">String</span><span class="sxs-lookup"><span data-stu-id="ef010-536">String</span></span>||<span data-ttu-id="ef010-537">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="ef010-537">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="ef010-538">function</span><span class="sxs-lookup"><span data-stu-id="ef010-538">function</span></span>||<span data-ttu-id="ef010-539">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ef010-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ef010-540">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ef010-540">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="ef010-541">Если результат превышает 1 МБ по размеру, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="ef010-541">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="ef010-542">Объект</span><span class="sxs-lookup"><span data-stu-id="ef010-542">Object</span></span>| <span data-ttu-id="ef010-543">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ef010-543">&lt;optional&gt;</span></span>|<span data-ttu-id="ef010-544">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="ef010-544">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ef010-545">Требования</span><span class="sxs-lookup"><span data-stu-id="ef010-545">Requirements</span></span>

|<span data-ttu-id="ef010-546">Requirement</span><span class="sxs-lookup"><span data-stu-id="ef010-546">Requirement</span></span>| <span data-ttu-id="ef010-547">Значение</span><span class="sxs-lookup"><span data-stu-id="ef010-547">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef010-548">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ef010-548">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef010-549">1.0</span><span class="sxs-lookup"><span data-stu-id="ef010-549">1.0</span></span>|
|[<span data-ttu-id="ef010-550">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ef010-550">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef010-551">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="ef010-551">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="ef010-552">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ef010-552">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ef010-553">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ef010-553">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ef010-554">Пример</span><span class="sxs-lookup"><span data-stu-id="ef010-554">Example</span></span>

<span data-ttu-id="ef010-555">В следующем примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="ef010-555">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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