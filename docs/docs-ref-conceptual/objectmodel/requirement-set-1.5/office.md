# <a name="office"></a><span data-ttu-id="79728-101">Office</span><span class="sxs-lookup"><span data-stu-id="79728-101">Office</span></span>

<span data-ttu-id="79728-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="79728-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="79728-104">Требования</span><span class="sxs-lookup"><span data-stu-id="79728-104">Requirements</span></span>

|<span data-ttu-id="79728-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="79728-105">Requirement</span></span>| <span data-ttu-id="79728-106">Значение</span><span class="sxs-lookup"><span data-stu-id="79728-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="79728-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="79728-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79728-108">1.0</span><span class="sxs-lookup"><span data-stu-id="79728-108">1.0</span></span>|
|[<span data-ttu-id="79728-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="79728-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79728-110">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="79728-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="79728-111">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="79728-111">Members and methods</span></span>

| <span data-ttu-id="79728-112">Элемент</span><span class="sxs-lookup"><span data-stu-id="79728-112">Member</span></span> | <span data-ttu-id="79728-113">Тип</span><span class="sxs-lookup"><span data-stu-id="79728-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="79728-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="79728-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="79728-115">Член</span><span class="sxs-lookup"><span data-stu-id="79728-115">Member</span></span> |
| [<span data-ttu-id="79728-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="79728-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="79728-117">Член</span><span class="sxs-lookup"><span data-stu-id="79728-117">Member</span></span> |
| [<span data-ttu-id="79728-118">EventType</span><span class="sxs-lookup"><span data-stu-id="79728-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="79728-119">Член</span><span class="sxs-lookup"><span data-stu-id="79728-119">Member</span></span> |
| [<span data-ttu-id="79728-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="79728-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="79728-121">Член</span><span class="sxs-lookup"><span data-stu-id="79728-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="79728-122">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="79728-122">Namespaces</span></span>

<span data-ttu-id="79728-123">[контекст](office.context.md): предоставляет общедоступные интерфейсы из пространства имен контекста API надстройки Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="79728-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="79728-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype). Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="79728-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="79728-125">Элементы</span><span class="sxs-lookup"><span data-stu-id="79728-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="79728-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="79728-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="79728-127">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="79728-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="79728-128">Тип:</span><span class="sxs-lookup"><span data-stu-id="79728-128">Type:</span></span>

*   <span data-ttu-id="79728-129">String</span><span class="sxs-lookup"><span data-stu-id="79728-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="79728-130">Свойства:</span><span class="sxs-lookup"><span data-stu-id="79728-130">Properties:</span></span>

|<span data-ttu-id="79728-131">Имя</span><span class="sxs-lookup"><span data-stu-id="79728-131">Name</span></span>| <span data-ttu-id="79728-132">Тип</span><span class="sxs-lookup"><span data-stu-id="79728-132">Type</span></span>| <span data-ttu-id="79728-133">Описание</span><span class="sxs-lookup"><span data-stu-id="79728-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="79728-134">String</span><span class="sxs-lookup"><span data-stu-id="79728-134">String</span></span>|<span data-ttu-id="79728-135">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="79728-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="79728-136">String</span><span class="sxs-lookup"><span data-stu-id="79728-136">String</span></span>|<span data-ttu-id="79728-137">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="79728-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="79728-138">Требования</span><span class="sxs-lookup"><span data-stu-id="79728-138">Requirements</span></span>

|<span data-ttu-id="79728-139">Requirement</span><span class="sxs-lookup"><span data-stu-id="79728-139">Requirement</span></span>| <span data-ttu-id="79728-140">Значение</span><span class="sxs-lookup"><span data-stu-id="79728-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="79728-141">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="79728-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79728-142">1.0</span><span class="sxs-lookup"><span data-stu-id="79728-142">1.0</span></span>|
|[<span data-ttu-id="79728-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="79728-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79728-144">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="79728-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="79728-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="79728-145">CoercionType :String</span></span>

<span data-ttu-id="79728-146">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="79728-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="79728-147">Тип:</span><span class="sxs-lookup"><span data-stu-id="79728-147">Type:</span></span>

*   <span data-ttu-id="79728-148">String</span><span class="sxs-lookup"><span data-stu-id="79728-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="79728-149">Свойства:</span><span class="sxs-lookup"><span data-stu-id="79728-149">Properties:</span></span>

|<span data-ttu-id="79728-150">Имя</span><span class="sxs-lookup"><span data-stu-id="79728-150">Name</span></span>| <span data-ttu-id="79728-151">Тип</span><span class="sxs-lookup"><span data-stu-id="79728-151">Type</span></span>| <span data-ttu-id="79728-152">Описание</span><span class="sxs-lookup"><span data-stu-id="79728-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="79728-153">String</span><span class="sxs-lookup"><span data-stu-id="79728-153">String</span></span>|<span data-ttu-id="79728-154">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="79728-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="79728-155">String</span><span class="sxs-lookup"><span data-stu-id="79728-155">String</span></span>|<span data-ttu-id="79728-156">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="79728-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="79728-157">Требования</span><span class="sxs-lookup"><span data-stu-id="79728-157">Requirements</span></span>

|<span data-ttu-id="79728-158">Requirement</span><span class="sxs-lookup"><span data-stu-id="79728-158">Requirement</span></span>| <span data-ttu-id="79728-159">Значение</span><span class="sxs-lookup"><span data-stu-id="79728-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="79728-160">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="79728-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79728-161">1.0</span><span class="sxs-lookup"><span data-stu-id="79728-161">1.0</span></span>|
|[<span data-ttu-id="79728-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="79728-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79728-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="79728-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="79728-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="79728-164">EventType :String</span></span>

<span data-ttu-id="79728-165">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="79728-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="79728-166">Тип:</span><span class="sxs-lookup"><span data-stu-id="79728-166">Type:</span></span>

*   <span data-ttu-id="79728-167">String</span><span class="sxs-lookup"><span data-stu-id="79728-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="79728-168">Свойства:</span><span class="sxs-lookup"><span data-stu-id="79728-168">Properties:</span></span>

| <span data-ttu-id="79728-169">Имя</span><span class="sxs-lookup"><span data-stu-id="79728-169">Name</span></span> | <span data-ttu-id="79728-170">Тип</span><span class="sxs-lookup"><span data-stu-id="79728-170">Type</span></span> | <span data-ttu-id="79728-171">Описание</span><span class="sxs-lookup"><span data-stu-id="79728-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="79728-172">String</span><span class="sxs-lookup"><span data-stu-id="79728-172">String</span></span> | <span data-ttu-id="79728-173">Выбранный элемент изменился.</span><span class="sxs-lookup"><span data-stu-id="79728-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="79728-174">Требования</span><span class="sxs-lookup"><span data-stu-id="79728-174">Requirements</span></span>

|<span data-ttu-id="79728-175">Requirement</span><span class="sxs-lookup"><span data-stu-id="79728-175">Requirement</span></span>| <span data-ttu-id="79728-176">Значение</span><span class="sxs-lookup"><span data-stu-id="79728-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="79728-177">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="79728-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79728-178">1.5</span><span class="sxs-lookup"><span data-stu-id="79728-178">1.5</span></span> |
|[<span data-ttu-id="79728-179">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="79728-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79728-180">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="79728-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="79728-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="79728-181">SourceProperty :String</span></span>

<span data-ttu-id="79728-182">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="79728-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="79728-183">Тип:</span><span class="sxs-lookup"><span data-stu-id="79728-183">Type:</span></span>

*   <span data-ttu-id="79728-184">String</span><span class="sxs-lookup"><span data-stu-id="79728-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="79728-185">Свойства:</span><span class="sxs-lookup"><span data-stu-id="79728-185">Properties:</span></span>

|<span data-ttu-id="79728-186">Имя</span><span class="sxs-lookup"><span data-stu-id="79728-186">Name</span></span>| <span data-ttu-id="79728-187">Тип</span><span class="sxs-lookup"><span data-stu-id="79728-187">Type</span></span>| <span data-ttu-id="79728-188">Описание</span><span class="sxs-lookup"><span data-stu-id="79728-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="79728-189">String</span><span class="sxs-lookup"><span data-stu-id="79728-189">String</span></span>|<span data-ttu-id="79728-190">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="79728-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="79728-191">String</span><span class="sxs-lookup"><span data-stu-id="79728-191">String</span></span>|<span data-ttu-id="79728-192">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="79728-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="79728-193">Требования</span><span class="sxs-lookup"><span data-stu-id="79728-193">Requirements</span></span>

|<span data-ttu-id="79728-194">Requirement</span><span class="sxs-lookup"><span data-stu-id="79728-194">Requirement</span></span>| <span data-ttu-id="79728-195">Значение</span><span class="sxs-lookup"><span data-stu-id="79728-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="79728-196">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="79728-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79728-197">1.0</span><span class="sxs-lookup"><span data-stu-id="79728-197">1.0</span></span>|
|[<span data-ttu-id="79728-198">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="79728-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79728-199">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="79728-199">Compose or read</span></span>|