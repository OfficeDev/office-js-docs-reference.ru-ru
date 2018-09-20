 

# <a name="office"></a><span data-ttu-id="0a913-101">Office</span><span class="sxs-lookup"><span data-stu-id="0a913-101">Office</span></span>

<span data-ttu-id="0a913-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="0a913-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0a913-104">Требования</span><span class="sxs-lookup"><span data-stu-id="0a913-104">Requirements</span></span>

|<span data-ttu-id="0a913-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="0a913-105">Requirement</span></span>| <span data-ttu-id="0a913-106">Значение</span><span class="sxs-lookup"><span data-stu-id="0a913-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="0a913-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0a913-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0a913-108">1.0</span><span class="sxs-lookup"><span data-stu-id="0a913-108">1.0</span></span>|
|[<span data-ttu-id="0a913-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0a913-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0a913-110">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0a913-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="0a913-111">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="0a913-111">Namespaces</span></span>

<span data-ttu-id="0a913-112">[контекст](office.context.md): предоставляет общедоступные интерфейсы из пространства имен контекста API надстройки Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="0a913-112">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="0a913-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype). Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="0a913-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="0a913-114">Элементы</span><span class="sxs-lookup"><span data-stu-id="0a913-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="0a913-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="0a913-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="0a913-116">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="0a913-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="0a913-117">Тип:</span><span class="sxs-lookup"><span data-stu-id="0a913-117">Type:</span></span>

*   <span data-ttu-id="0a913-118">String</span><span class="sxs-lookup"><span data-stu-id="0a913-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0a913-119">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0a913-119">Properties:</span></span>

|<span data-ttu-id="0a913-120">Имя</span><span class="sxs-lookup"><span data-stu-id="0a913-120">Name</span></span>| <span data-ttu-id="0a913-121">Тип</span><span class="sxs-lookup"><span data-stu-id="0a913-121">Type</span></span>| <span data-ttu-id="0a913-122">Описание</span><span class="sxs-lookup"><span data-stu-id="0a913-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="0a913-123">String</span><span class="sxs-lookup"><span data-stu-id="0a913-123">String</span></span>|<span data-ttu-id="0a913-124">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="0a913-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="0a913-125">String</span><span class="sxs-lookup"><span data-stu-id="0a913-125">String</span></span>|<span data-ttu-id="0a913-126">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="0a913-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0a913-127">Требования</span><span class="sxs-lookup"><span data-stu-id="0a913-127">Requirements</span></span>

|<span data-ttu-id="0a913-128">Requirement</span><span class="sxs-lookup"><span data-stu-id="0a913-128">Requirement</span></span>| <span data-ttu-id="0a913-129">Значение</span><span class="sxs-lookup"><span data-stu-id="0a913-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="0a913-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0a913-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0a913-131">1.0</span><span class="sxs-lookup"><span data-stu-id="0a913-131">1.0</span></span>|
|[<span data-ttu-id="0a913-132">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0a913-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0a913-133">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0a913-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="0a913-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="0a913-134">CoercionType :String</span></span>

<span data-ttu-id="0a913-135">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="0a913-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0a913-136">Тип:</span><span class="sxs-lookup"><span data-stu-id="0a913-136">Type:</span></span>

*   <span data-ttu-id="0a913-137">String</span><span class="sxs-lookup"><span data-stu-id="0a913-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0a913-138">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0a913-138">Properties:</span></span>

|<span data-ttu-id="0a913-139">Имя</span><span class="sxs-lookup"><span data-stu-id="0a913-139">Name</span></span>| <span data-ttu-id="0a913-140">Тип</span><span class="sxs-lookup"><span data-stu-id="0a913-140">Type</span></span>| <span data-ttu-id="0a913-141">Описание</span><span class="sxs-lookup"><span data-stu-id="0a913-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="0a913-142">String</span><span class="sxs-lookup"><span data-stu-id="0a913-142">String</span></span>|<span data-ttu-id="0a913-143">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="0a913-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="0a913-144">String</span><span class="sxs-lookup"><span data-stu-id="0a913-144">String</span></span>|<span data-ttu-id="0a913-145">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="0a913-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0a913-146">Требования</span><span class="sxs-lookup"><span data-stu-id="0a913-146">Requirements</span></span>

|<span data-ttu-id="0a913-147">Requirement</span><span class="sxs-lookup"><span data-stu-id="0a913-147">Requirement</span></span>| <span data-ttu-id="0a913-148">Значение</span><span class="sxs-lookup"><span data-stu-id="0a913-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="0a913-149">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0a913-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0a913-150">1.0</span><span class="sxs-lookup"><span data-stu-id="0a913-150">1.0</span></span>|
|[<span data-ttu-id="0a913-151">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0a913-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0a913-152">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0a913-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="0a913-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="0a913-153">SourceProperty :String</span></span>

<span data-ttu-id="0a913-154">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="0a913-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0a913-155">Тип:</span><span class="sxs-lookup"><span data-stu-id="0a913-155">Type:</span></span>

*   <span data-ttu-id="0a913-156">String</span><span class="sxs-lookup"><span data-stu-id="0a913-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0a913-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0a913-157">Properties:</span></span>

|<span data-ttu-id="0a913-158">Имя</span><span class="sxs-lookup"><span data-stu-id="0a913-158">Name</span></span>| <span data-ttu-id="0a913-159">Тип</span><span class="sxs-lookup"><span data-stu-id="0a913-159">Type</span></span>| <span data-ttu-id="0a913-160">Описание</span><span class="sxs-lookup"><span data-stu-id="0a913-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="0a913-161">String</span><span class="sxs-lookup"><span data-stu-id="0a913-161">String</span></span>|<span data-ttu-id="0a913-162">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="0a913-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="0a913-163">String</span><span class="sxs-lookup"><span data-stu-id="0a913-163">String</span></span>|<span data-ttu-id="0a913-164">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="0a913-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0a913-165">Требования</span><span class="sxs-lookup"><span data-stu-id="0a913-165">Requirements</span></span>

|<span data-ttu-id="0a913-166">Requirement</span><span class="sxs-lookup"><span data-stu-id="0a913-166">Requirement</span></span>| <span data-ttu-id="0a913-167">Значение</span><span class="sxs-lookup"><span data-stu-id="0a913-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="0a913-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0a913-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0a913-169">1.0</span><span class="sxs-lookup"><span data-stu-id="0a913-169">1.0</span></span>|
|[<span data-ttu-id="0a913-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0a913-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0a913-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0a913-171">Compose or read</span></span>|