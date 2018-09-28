 

# <a name="office"></a><span data-ttu-id="3cecd-101">Office</span><span class="sxs-lookup"><span data-stu-id="3cecd-101">Office</span></span>

<span data-ttu-id="3cecd-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="3cecd-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3cecd-104">Требования</span><span class="sxs-lookup"><span data-stu-id="3cecd-104">Requirements</span></span>

|<span data-ttu-id="3cecd-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="3cecd-105">Requirement</span></span>| <span data-ttu-id="3cecd-106">Значение</span><span class="sxs-lookup"><span data-stu-id="3cecd-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="3cecd-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3cecd-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3cecd-108">1.0</span><span class="sxs-lookup"><span data-stu-id="3cecd-108">1.0</span></span>|
|[<span data-ttu-id="3cecd-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3cecd-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3cecd-110">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3cecd-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3cecd-111">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="3cecd-111">Members and methods</span></span>

| <span data-ttu-id="3cecd-112">Элемент</span><span class="sxs-lookup"><span data-stu-id="3cecd-112">Member</span></span> | <span data-ttu-id="3cecd-113">Тип</span><span class="sxs-lookup"><span data-stu-id="3cecd-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3cecd-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="3cecd-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="3cecd-115">Член</span><span class="sxs-lookup"><span data-stu-id="3cecd-115">Member</span></span> |
| [<span data-ttu-id="3cecd-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="3cecd-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="3cecd-117">Член</span><span class="sxs-lookup"><span data-stu-id="3cecd-117">Member</span></span> |
| [<span data-ttu-id="3cecd-118">EventType</span><span class="sxs-lookup"><span data-stu-id="3cecd-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="3cecd-119">Член</span><span class="sxs-lookup"><span data-stu-id="3cecd-119">Member</span></span> |
| [<span data-ttu-id="3cecd-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="3cecd-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="3cecd-121">Член</span><span class="sxs-lookup"><span data-stu-id="3cecd-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="3cecd-122">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="3cecd-122">Namespaces</span></span>

<span data-ttu-id="3cecd-123">[контекст](office.context.md): предоставляет общедоступные интерфейсы из пространства имен контекста API надстройки Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="3cecd-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="3cecd-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype). Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="3cecd-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="3cecd-125">Элементы</span><span class="sxs-lookup"><span data-stu-id="3cecd-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="3cecd-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="3cecd-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="3cecd-127">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="3cecd-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="3cecd-128">Тип:</span><span class="sxs-lookup"><span data-stu-id="3cecd-128">Type:</span></span>

*   <span data-ttu-id="3cecd-129">String</span><span class="sxs-lookup"><span data-stu-id="3cecd-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3cecd-130">Свойства:</span><span class="sxs-lookup"><span data-stu-id="3cecd-130">Properties:</span></span>

|<span data-ttu-id="3cecd-131">Имя</span><span class="sxs-lookup"><span data-stu-id="3cecd-131">Name</span></span>| <span data-ttu-id="3cecd-132">Тип</span><span class="sxs-lookup"><span data-stu-id="3cecd-132">Type</span></span>| <span data-ttu-id="3cecd-133">Описание</span><span class="sxs-lookup"><span data-stu-id="3cecd-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="3cecd-134">String</span><span class="sxs-lookup"><span data-stu-id="3cecd-134">String</span></span>|<span data-ttu-id="3cecd-135">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="3cecd-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="3cecd-136">String</span><span class="sxs-lookup"><span data-stu-id="3cecd-136">String</span></span>|<span data-ttu-id="3cecd-137">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="3cecd-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3cecd-138">Требования</span><span class="sxs-lookup"><span data-stu-id="3cecd-138">Requirements</span></span>

|<span data-ttu-id="3cecd-139">Requirement</span><span class="sxs-lookup"><span data-stu-id="3cecd-139">Requirement</span></span>| <span data-ttu-id="3cecd-140">Значение</span><span class="sxs-lookup"><span data-stu-id="3cecd-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="3cecd-141">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3cecd-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3cecd-142">1.0</span><span class="sxs-lookup"><span data-stu-id="3cecd-142">1.0</span></span>|
|[<span data-ttu-id="3cecd-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3cecd-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3cecd-144">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3cecd-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="3cecd-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="3cecd-145">CoercionType :String</span></span>

<span data-ttu-id="3cecd-146">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="3cecd-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3cecd-147">Тип:</span><span class="sxs-lookup"><span data-stu-id="3cecd-147">Type:</span></span>

*   <span data-ttu-id="3cecd-148">String</span><span class="sxs-lookup"><span data-stu-id="3cecd-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3cecd-149">Свойства:</span><span class="sxs-lookup"><span data-stu-id="3cecd-149">Properties:</span></span>

|<span data-ttu-id="3cecd-150">Имя</span><span class="sxs-lookup"><span data-stu-id="3cecd-150">Name</span></span>| <span data-ttu-id="3cecd-151">Тип</span><span class="sxs-lookup"><span data-stu-id="3cecd-151">Type</span></span>| <span data-ttu-id="3cecd-152">Описание</span><span class="sxs-lookup"><span data-stu-id="3cecd-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="3cecd-153">String</span><span class="sxs-lookup"><span data-stu-id="3cecd-153">String</span></span>|<span data-ttu-id="3cecd-154">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="3cecd-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="3cecd-155">String</span><span class="sxs-lookup"><span data-stu-id="3cecd-155">String</span></span>|<span data-ttu-id="3cecd-156">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="3cecd-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3cecd-157">Требования</span><span class="sxs-lookup"><span data-stu-id="3cecd-157">Requirements</span></span>

|<span data-ttu-id="3cecd-158">Requirement</span><span class="sxs-lookup"><span data-stu-id="3cecd-158">Requirement</span></span>| <span data-ttu-id="3cecd-159">Значение</span><span class="sxs-lookup"><span data-stu-id="3cecd-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="3cecd-160">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3cecd-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3cecd-161">1.0</span><span class="sxs-lookup"><span data-stu-id="3cecd-161">1.0</span></span>|
|[<span data-ttu-id="3cecd-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3cecd-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3cecd-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3cecd-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="3cecd-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="3cecd-164">EventType :String</span></span>

<span data-ttu-id="3cecd-165">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="3cecd-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="3cecd-166">Тип:</span><span class="sxs-lookup"><span data-stu-id="3cecd-166">Type:</span></span>

*   <span data-ttu-id="3cecd-167">String</span><span class="sxs-lookup"><span data-stu-id="3cecd-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3cecd-168">Свойства:</span><span class="sxs-lookup"><span data-stu-id="3cecd-168">Properties:</span></span>

| <span data-ttu-id="3cecd-169">Имя</span><span class="sxs-lookup"><span data-stu-id="3cecd-169">Name</span></span> | <span data-ttu-id="3cecd-170">Тип</span><span class="sxs-lookup"><span data-stu-id="3cecd-170">Type</span></span> | <span data-ttu-id="3cecd-171">Описание</span><span class="sxs-lookup"><span data-stu-id="3cecd-171">Description</span></span> | <span data-ttu-id="3cecd-172">Минимальное требование set</span><span class="sxs-lookup"><span data-stu-id="3cecd-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="3cecd-173">String</span><span class="sxs-lookup"><span data-stu-id="3cecd-173">String</span></span> | <span data-ttu-id="3cecd-174">Дата или время выбранной встречи или серии, была изменена.</span><span class="sxs-lookup"><span data-stu-id="3cecd-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="3cecd-175">1.7</span><span class="sxs-lookup"><span data-stu-id="3cecd-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="3cecd-176">String</span><span class="sxs-lookup"><span data-stu-id="3cecd-176">String</span></span> | <span data-ttu-id="3cecd-177">Выбранный элемент изменился.</span><span class="sxs-lookup"><span data-stu-id="3cecd-177">The selected item has changed.</span></span> | <span data-ttu-id="3cecd-178">1.5</span><span class="sxs-lookup"><span data-stu-id="3cecd-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="3cecd-179">String</span><span class="sxs-lookup"><span data-stu-id="3cecd-179">String</span></span> | <span data-ttu-id="3cecd-180">Список получателей в выбранное расположение элемента или встречи, была изменена.</span><span class="sxs-lookup"><span data-stu-id="3cecd-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="3cecd-181">1.7</span><span class="sxs-lookup"><span data-stu-id="3cecd-181">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="3cecd-182">String</span><span class="sxs-lookup"><span data-stu-id="3cecd-182">String</span></span> | <span data-ttu-id="3cecd-183">Шаблон повторения выбранного серии, была изменена.</span><span class="sxs-lookup"><span data-stu-id="3cecd-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="3cecd-184">1.7</span><span class="sxs-lookup"><span data-stu-id="3cecd-184">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3cecd-185">Требования</span><span class="sxs-lookup"><span data-stu-id="3cecd-185">Requirements</span></span>

|<span data-ttu-id="3cecd-186">Requirement</span><span class="sxs-lookup"><span data-stu-id="3cecd-186">Requirement</span></span>| <span data-ttu-id="3cecd-187">Значение</span><span class="sxs-lookup"><span data-stu-id="3cecd-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="3cecd-188">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3cecd-188">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3cecd-189">1.5</span><span class="sxs-lookup"><span data-stu-id="3cecd-189">1.5</span></span> |
|[<span data-ttu-id="3cecd-190">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3cecd-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3cecd-191">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3cecd-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="3cecd-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="3cecd-192">SourceProperty :String</span></span>

<span data-ttu-id="3cecd-193">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="3cecd-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3cecd-194">Тип:</span><span class="sxs-lookup"><span data-stu-id="3cecd-194">Type:</span></span>

*   <span data-ttu-id="3cecd-195">String</span><span class="sxs-lookup"><span data-stu-id="3cecd-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3cecd-196">Свойства:</span><span class="sxs-lookup"><span data-stu-id="3cecd-196">Properties:</span></span>

|<span data-ttu-id="3cecd-197">Имя</span><span class="sxs-lookup"><span data-stu-id="3cecd-197">Name</span></span>| <span data-ttu-id="3cecd-198">Тип</span><span class="sxs-lookup"><span data-stu-id="3cecd-198">Type</span></span>| <span data-ttu-id="3cecd-199">Описание</span><span class="sxs-lookup"><span data-stu-id="3cecd-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="3cecd-200">String</span><span class="sxs-lookup"><span data-stu-id="3cecd-200">String</span></span>|<span data-ttu-id="3cecd-201">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="3cecd-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="3cecd-202">String</span><span class="sxs-lookup"><span data-stu-id="3cecd-202">String</span></span>|<span data-ttu-id="3cecd-203">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="3cecd-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3cecd-204">Требования</span><span class="sxs-lookup"><span data-stu-id="3cecd-204">Requirements</span></span>

|<span data-ttu-id="3cecd-205">Requirement</span><span class="sxs-lookup"><span data-stu-id="3cecd-205">Requirement</span></span>| <span data-ttu-id="3cecd-206">Значение</span><span class="sxs-lookup"><span data-stu-id="3cecd-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="3cecd-207">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3cecd-207">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3cecd-208">1.0</span><span class="sxs-lookup"><span data-stu-id="3cecd-208">1.0</span></span>|
|[<span data-ttu-id="3cecd-209">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3cecd-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3cecd-210">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3cecd-210">Compose or read</span></span>|