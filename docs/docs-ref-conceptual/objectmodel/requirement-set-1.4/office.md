 

# <a name="office"></a><span data-ttu-id="24fb0-101">Office</span><span class="sxs-lookup"><span data-stu-id="24fb0-101">Office</span></span>

<span data-ttu-id="24fb0-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="24fb0-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="24fb0-104">Требования</span><span class="sxs-lookup"><span data-stu-id="24fb0-104">Requirements</span></span>

|<span data-ttu-id="24fb0-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="24fb0-105">Requirement</span></span>| <span data-ttu-id="24fb0-106">Значение</span><span class="sxs-lookup"><span data-stu-id="24fb0-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="24fb0-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="24fb0-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24fb0-108">1.0</span><span class="sxs-lookup"><span data-stu-id="24fb0-108">1.0</span></span>|
|[<span data-ttu-id="24fb0-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="24fb0-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="24fb0-110">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="24fb0-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="24fb0-111">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="24fb0-111">Namespaces</span></span>

<span data-ttu-id="24fb0-112">[контекст](Office.context.md): предоставляет общедоступные интерфейсы из пространства имен контекста API надстройки Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="24fb0-112">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="24fb0-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype). Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="24fb0-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="24fb0-114">Элементы</span><span class="sxs-lookup"><span data-stu-id="24fb0-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="24fb0-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="24fb0-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="24fb0-116">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="24fb0-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="24fb0-117">Тип:</span><span class="sxs-lookup"><span data-stu-id="24fb0-117">Type:</span></span>

*   <span data-ttu-id="24fb0-118">String</span><span class="sxs-lookup"><span data-stu-id="24fb0-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="24fb0-119">Свойства:</span><span class="sxs-lookup"><span data-stu-id="24fb0-119">Properties:</span></span>

|<span data-ttu-id="24fb0-120">Имя</span><span class="sxs-lookup"><span data-stu-id="24fb0-120">Name</span></span>| <span data-ttu-id="24fb0-121">Тип</span><span class="sxs-lookup"><span data-stu-id="24fb0-121">Type</span></span>| <span data-ttu-id="24fb0-122">Описание</span><span class="sxs-lookup"><span data-stu-id="24fb0-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="24fb0-123">String</span><span class="sxs-lookup"><span data-stu-id="24fb0-123">String</span></span>|<span data-ttu-id="24fb0-124">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="24fb0-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="24fb0-125">String</span><span class="sxs-lookup"><span data-stu-id="24fb0-125">String</span></span>|<span data-ttu-id="24fb0-126">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="24fb0-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="24fb0-127">Требования</span><span class="sxs-lookup"><span data-stu-id="24fb0-127">Requirements</span></span>

|<span data-ttu-id="24fb0-128">Requirement</span><span class="sxs-lookup"><span data-stu-id="24fb0-128">Requirement</span></span>| <span data-ttu-id="24fb0-129">Значение</span><span class="sxs-lookup"><span data-stu-id="24fb0-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="24fb0-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="24fb0-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24fb0-131">1.0</span><span class="sxs-lookup"><span data-stu-id="24fb0-131">1.0</span></span>|
|[<span data-ttu-id="24fb0-132">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="24fb0-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="24fb0-133">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="24fb0-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="24fb0-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="24fb0-134">CoercionType :String</span></span>

<span data-ttu-id="24fb0-135">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="24fb0-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="24fb0-136">Тип:</span><span class="sxs-lookup"><span data-stu-id="24fb0-136">Type:</span></span>

*   <span data-ttu-id="24fb0-137">String</span><span class="sxs-lookup"><span data-stu-id="24fb0-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="24fb0-138">Свойства:</span><span class="sxs-lookup"><span data-stu-id="24fb0-138">Properties:</span></span>

|<span data-ttu-id="24fb0-139">Имя</span><span class="sxs-lookup"><span data-stu-id="24fb0-139">Name</span></span>| <span data-ttu-id="24fb0-140">Тип</span><span class="sxs-lookup"><span data-stu-id="24fb0-140">Type</span></span>| <span data-ttu-id="24fb0-141">Описание</span><span class="sxs-lookup"><span data-stu-id="24fb0-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="24fb0-142">String</span><span class="sxs-lookup"><span data-stu-id="24fb0-142">String</span></span>|<span data-ttu-id="24fb0-143">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="24fb0-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="24fb0-144">String</span><span class="sxs-lookup"><span data-stu-id="24fb0-144">String</span></span>|<span data-ttu-id="24fb0-145">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="24fb0-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="24fb0-146">Требования</span><span class="sxs-lookup"><span data-stu-id="24fb0-146">Requirements</span></span>

|<span data-ttu-id="24fb0-147">Requirement</span><span class="sxs-lookup"><span data-stu-id="24fb0-147">Requirement</span></span>| <span data-ttu-id="24fb0-148">Значение</span><span class="sxs-lookup"><span data-stu-id="24fb0-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="24fb0-149">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="24fb0-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24fb0-150">1.0</span><span class="sxs-lookup"><span data-stu-id="24fb0-150">1.0</span></span>|
|[<span data-ttu-id="24fb0-151">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="24fb0-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="24fb0-152">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="24fb0-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="24fb0-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="24fb0-153">SourceProperty :String</span></span>

<span data-ttu-id="24fb0-154">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="24fb0-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="24fb0-155">Тип:</span><span class="sxs-lookup"><span data-stu-id="24fb0-155">Type:</span></span>

*   <span data-ttu-id="24fb0-156">String</span><span class="sxs-lookup"><span data-stu-id="24fb0-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="24fb0-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="24fb0-157">Properties:</span></span>

|<span data-ttu-id="24fb0-158">Имя</span><span class="sxs-lookup"><span data-stu-id="24fb0-158">Name</span></span>| <span data-ttu-id="24fb0-159">Тип</span><span class="sxs-lookup"><span data-stu-id="24fb0-159">Type</span></span>| <span data-ttu-id="24fb0-160">Описание</span><span class="sxs-lookup"><span data-stu-id="24fb0-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="24fb0-161">String</span><span class="sxs-lookup"><span data-stu-id="24fb0-161">String</span></span>|<span data-ttu-id="24fb0-162">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="24fb0-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="24fb0-163">String</span><span class="sxs-lookup"><span data-stu-id="24fb0-163">String</span></span>|<span data-ttu-id="24fb0-164">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="24fb0-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="24fb0-165">Требования</span><span class="sxs-lookup"><span data-stu-id="24fb0-165">Requirements</span></span>

|<span data-ttu-id="24fb0-166">Requirement</span><span class="sxs-lookup"><span data-stu-id="24fb0-166">Requirement</span></span>| <span data-ttu-id="24fb0-167">Значение</span><span class="sxs-lookup"><span data-stu-id="24fb0-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="24fb0-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="24fb0-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24fb0-169">1.0</span><span class="sxs-lookup"><span data-stu-id="24fb0-169">1.0</span></span>|
|[<span data-ttu-id="24fb0-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="24fb0-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="24fb0-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="24fb0-171">Compose or read</span></span>|