 

# <a name="office"></a><span data-ttu-id="93cae-101">Office</span><span class="sxs-lookup"><span data-stu-id="93cae-101">Office</span></span>

<span data-ttu-id="93cae-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="93cae-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="93cae-104">Требования</span><span class="sxs-lookup"><span data-stu-id="93cae-104">Requirements</span></span>

|<span data-ttu-id="93cae-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="93cae-105">Requirement</span></span>| <span data-ttu-id="93cae-106">Значение</span><span class="sxs-lookup"><span data-stu-id="93cae-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="93cae-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="93cae-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93cae-108">1.0</span><span class="sxs-lookup"><span data-stu-id="93cae-108">1.0</span></span>|
|[<span data-ttu-id="93cae-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="93cae-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="93cae-110">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="93cae-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="93cae-111">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="93cae-111">Members and methods</span></span>

| <span data-ttu-id="93cae-112">Элемент</span><span class="sxs-lookup"><span data-stu-id="93cae-112">Member</span></span> | <span data-ttu-id="93cae-113">Тип</span><span class="sxs-lookup"><span data-stu-id="93cae-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="93cae-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="93cae-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="93cae-115">Член</span><span class="sxs-lookup"><span data-stu-id="93cae-115">Member</span></span> |
| [<span data-ttu-id="93cae-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="93cae-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="93cae-117">Член</span><span class="sxs-lookup"><span data-stu-id="93cae-117">Member</span></span> |
| [<span data-ttu-id="93cae-118">EventType</span><span class="sxs-lookup"><span data-stu-id="93cae-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="93cae-119">Член</span><span class="sxs-lookup"><span data-stu-id="93cae-119">Member</span></span> |
| [<span data-ttu-id="93cae-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="93cae-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="93cae-121">Член</span><span class="sxs-lookup"><span data-stu-id="93cae-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="93cae-122">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="93cae-122">Namespaces</span></span>

<span data-ttu-id="93cae-123">[контекст](office.context.md): предоставляет общедоступные интерфейсы из пространства имен контекста API надстройки Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="93cae-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="93cae-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype). Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="93cae-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="93cae-125">Элементы</span><span class="sxs-lookup"><span data-stu-id="93cae-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="93cae-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="93cae-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="93cae-127">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="93cae-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="93cae-128">Тип:</span><span class="sxs-lookup"><span data-stu-id="93cae-128">Type:</span></span>

*   <span data-ttu-id="93cae-129">String</span><span class="sxs-lookup"><span data-stu-id="93cae-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="93cae-130">Свойства:</span><span class="sxs-lookup"><span data-stu-id="93cae-130">Properties:</span></span>

|<span data-ttu-id="93cae-131">Имя</span><span class="sxs-lookup"><span data-stu-id="93cae-131">Name</span></span>| <span data-ttu-id="93cae-132">Тип</span><span class="sxs-lookup"><span data-stu-id="93cae-132">Type</span></span>| <span data-ttu-id="93cae-133">Описание</span><span class="sxs-lookup"><span data-stu-id="93cae-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="93cae-134">String</span><span class="sxs-lookup"><span data-stu-id="93cae-134">String</span></span>|<span data-ttu-id="93cae-135">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="93cae-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="93cae-136">String</span><span class="sxs-lookup"><span data-stu-id="93cae-136">String</span></span>|<span data-ttu-id="93cae-137">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="93cae-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="93cae-138">Требования</span><span class="sxs-lookup"><span data-stu-id="93cae-138">Requirements</span></span>

|<span data-ttu-id="93cae-139">Requirement</span><span class="sxs-lookup"><span data-stu-id="93cae-139">Requirement</span></span>| <span data-ttu-id="93cae-140">Значение</span><span class="sxs-lookup"><span data-stu-id="93cae-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="93cae-141">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="93cae-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93cae-142">1.0</span><span class="sxs-lookup"><span data-stu-id="93cae-142">1.0</span></span>|
|[<span data-ttu-id="93cae-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="93cae-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="93cae-144">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="93cae-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="93cae-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="93cae-145">CoercionType :String</span></span>

<span data-ttu-id="93cae-146">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="93cae-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="93cae-147">Тип:</span><span class="sxs-lookup"><span data-stu-id="93cae-147">Type:</span></span>

*   <span data-ttu-id="93cae-148">String</span><span class="sxs-lookup"><span data-stu-id="93cae-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="93cae-149">Свойства:</span><span class="sxs-lookup"><span data-stu-id="93cae-149">Properties:</span></span>

|<span data-ttu-id="93cae-150">Имя</span><span class="sxs-lookup"><span data-stu-id="93cae-150">Name</span></span>| <span data-ttu-id="93cae-151">Тип</span><span class="sxs-lookup"><span data-stu-id="93cae-151">Type</span></span>| <span data-ttu-id="93cae-152">Описание</span><span class="sxs-lookup"><span data-stu-id="93cae-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="93cae-153">String</span><span class="sxs-lookup"><span data-stu-id="93cae-153">String</span></span>|<span data-ttu-id="93cae-154">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="93cae-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="93cae-155">String</span><span class="sxs-lookup"><span data-stu-id="93cae-155">String</span></span>|<span data-ttu-id="93cae-156">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="93cae-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="93cae-157">Требования</span><span class="sxs-lookup"><span data-stu-id="93cae-157">Requirements</span></span>

|<span data-ttu-id="93cae-158">Requirement</span><span class="sxs-lookup"><span data-stu-id="93cae-158">Requirement</span></span>| <span data-ttu-id="93cae-159">Значение</span><span class="sxs-lookup"><span data-stu-id="93cae-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="93cae-160">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="93cae-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93cae-161">1.0</span><span class="sxs-lookup"><span data-stu-id="93cae-161">1.0</span></span>|
|[<span data-ttu-id="93cae-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="93cae-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="93cae-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="93cae-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="93cae-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="93cae-164">EventType :String</span></span>

<span data-ttu-id="93cae-165">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="93cae-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="93cae-166">Тип:</span><span class="sxs-lookup"><span data-stu-id="93cae-166">Type:</span></span>

*   <span data-ttu-id="93cae-167">String</span><span class="sxs-lookup"><span data-stu-id="93cae-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="93cae-168">Свойства:</span><span class="sxs-lookup"><span data-stu-id="93cae-168">Properties:</span></span>

| <span data-ttu-id="93cae-169">Имя</span><span class="sxs-lookup"><span data-stu-id="93cae-169">Name</span></span> | <span data-ttu-id="93cae-170">Тип</span><span class="sxs-lookup"><span data-stu-id="93cae-170">Type</span></span> | <span data-ttu-id="93cae-171">Описание</span><span class="sxs-lookup"><span data-stu-id="93cae-171">Description</span></span> | <span data-ttu-id="93cae-172">Минимальное требование set</span><span class="sxs-lookup"><span data-stu-id="93cae-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="93cae-173">String</span><span class="sxs-lookup"><span data-stu-id="93cae-173">String</span></span> | <span data-ttu-id="93cae-174">Дата или время выбранной встречи или серии, была изменена.</span><span class="sxs-lookup"><span data-stu-id="93cae-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="93cae-175">1.7</span><span class="sxs-lookup"><span data-stu-id="93cae-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="93cae-176">String</span><span class="sxs-lookup"><span data-stu-id="93cae-176">String</span></span> | <span data-ttu-id="93cae-177">Выбранный элемент изменился.</span><span class="sxs-lookup"><span data-stu-id="93cae-177">The selected item has changed.</span></span> | <span data-ttu-id="93cae-178">1.5</span><span class="sxs-lookup"><span data-stu-id="93cae-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="93cae-179">String</span><span class="sxs-lookup"><span data-stu-id="93cae-179">String</span></span> | <span data-ttu-id="93cae-180">Выбранный элемент изменился.</span><span class="sxs-lookup"><span data-stu-id="93cae-180">The selected item has changed.</span></span> | <span data-ttu-id="93cae-181">Preview</span><span class="sxs-lookup"><span data-stu-id="93cae-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="93cae-182">String</span><span class="sxs-lookup"><span data-stu-id="93cae-182">String</span></span> | <span data-ttu-id="93cae-183">Список получателей в выбранное расположение элемента или встречи, была изменена.</span><span class="sxs-lookup"><span data-stu-id="93cae-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="93cae-184">1.7</span><span class="sxs-lookup"><span data-stu-id="93cae-184">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="93cae-185">String</span><span class="sxs-lookup"><span data-stu-id="93cae-185">String</span></span> | <span data-ttu-id="93cae-186">Шаблон повторения выбранного серии, была изменена.</span><span class="sxs-lookup"><span data-stu-id="93cae-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="93cae-187">1.7</span><span class="sxs-lookup"><span data-stu-id="93cae-187">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="93cae-188">Требования</span><span class="sxs-lookup"><span data-stu-id="93cae-188">Requirements</span></span>

|<span data-ttu-id="93cae-189">Requirement</span><span class="sxs-lookup"><span data-stu-id="93cae-189">Requirement</span></span>| <span data-ttu-id="93cae-190">Значение</span><span class="sxs-lookup"><span data-stu-id="93cae-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="93cae-191">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="93cae-191">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93cae-192">1.5</span><span class="sxs-lookup"><span data-stu-id="93cae-192">1.5</span></span> |
|[<span data-ttu-id="93cae-193">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="93cae-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="93cae-194">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="93cae-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="93cae-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="93cae-195">SourceProperty :String</span></span>

<span data-ttu-id="93cae-196">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="93cae-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="93cae-197">Тип:</span><span class="sxs-lookup"><span data-stu-id="93cae-197">Type:</span></span>

*   <span data-ttu-id="93cae-198">String</span><span class="sxs-lookup"><span data-stu-id="93cae-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="93cae-199">Свойства:</span><span class="sxs-lookup"><span data-stu-id="93cae-199">Properties:</span></span>

|<span data-ttu-id="93cae-200">Имя</span><span class="sxs-lookup"><span data-stu-id="93cae-200">Name</span></span>| <span data-ttu-id="93cae-201">Тип</span><span class="sxs-lookup"><span data-stu-id="93cae-201">Type</span></span>| <span data-ttu-id="93cae-202">Описание</span><span class="sxs-lookup"><span data-stu-id="93cae-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="93cae-203">String</span><span class="sxs-lookup"><span data-stu-id="93cae-203">String</span></span>|<span data-ttu-id="93cae-204">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="93cae-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="93cae-205">String</span><span class="sxs-lookup"><span data-stu-id="93cae-205">String</span></span>|<span data-ttu-id="93cae-206">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="93cae-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="93cae-207">Требования</span><span class="sxs-lookup"><span data-stu-id="93cae-207">Requirements</span></span>

|<span data-ttu-id="93cae-208">Requirement</span><span class="sxs-lookup"><span data-stu-id="93cae-208">Requirement</span></span>| <span data-ttu-id="93cae-209">Значение</span><span class="sxs-lookup"><span data-stu-id="93cae-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="93cae-210">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="93cae-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93cae-211">1.0</span><span class="sxs-lookup"><span data-stu-id="93cae-211">1.0</span></span>|
|[<span data-ttu-id="93cae-212">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="93cae-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="93cae-213">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="93cae-213">Compose or read</span></span>|