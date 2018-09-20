 

# <a name="office"></a><span data-ttu-id="a5304-101">Office</span><span class="sxs-lookup"><span data-stu-id="a5304-101">Office</span></span>

<span data-ttu-id="a5304-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="a5304-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a5304-104">Требования</span><span class="sxs-lookup"><span data-stu-id="a5304-104">Requirements</span></span>

|<span data-ttu-id="a5304-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="a5304-105">Requirement</span></span>| <span data-ttu-id="a5304-106">Значение</span><span class="sxs-lookup"><span data-stu-id="a5304-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5304-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a5304-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5304-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a5304-108">1.0</span></span>|
|[<span data-ttu-id="a5304-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a5304-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5304-110">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a5304-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="a5304-111">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="a5304-111">Namespaces</span></span>

<span data-ttu-id="a5304-112">[контекст](Office.context.md): предоставляет общедоступные интерфейсы из пространства имен контекста API надстройки Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="a5304-112">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="a5304-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype). Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="a5304-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="a5304-114">Элементы</span><span class="sxs-lookup"><span data-stu-id="a5304-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="a5304-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="a5304-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="a5304-116">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="a5304-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a5304-117">Тип:</span><span class="sxs-lookup"><span data-stu-id="a5304-117">Type:</span></span>

*   <span data-ttu-id="a5304-118">String</span><span class="sxs-lookup"><span data-stu-id="a5304-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a5304-119">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a5304-119">Properties:</span></span>

|<span data-ttu-id="a5304-120">Имя</span><span class="sxs-lookup"><span data-stu-id="a5304-120">Name</span></span>| <span data-ttu-id="a5304-121">Тип</span><span class="sxs-lookup"><span data-stu-id="a5304-121">Type</span></span>| <span data-ttu-id="a5304-122">Описание</span><span class="sxs-lookup"><span data-stu-id="a5304-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a5304-123">String</span><span class="sxs-lookup"><span data-stu-id="a5304-123">String</span></span>|<span data-ttu-id="a5304-124">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="a5304-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a5304-125">String</span><span class="sxs-lookup"><span data-stu-id="a5304-125">String</span></span>|<span data-ttu-id="a5304-126">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="a5304-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5304-127">Требования</span><span class="sxs-lookup"><span data-stu-id="a5304-127">Requirements</span></span>

|<span data-ttu-id="a5304-128">Requirement</span><span class="sxs-lookup"><span data-stu-id="a5304-128">Requirement</span></span>| <span data-ttu-id="a5304-129">Значение</span><span class="sxs-lookup"><span data-stu-id="a5304-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5304-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a5304-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5304-131">1.0</span><span class="sxs-lookup"><span data-stu-id="a5304-131">1.0</span></span>|
|[<span data-ttu-id="a5304-132">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a5304-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5304-133">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a5304-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="a5304-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="a5304-134">CoercionType :String</span></span>

<span data-ttu-id="a5304-135">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="a5304-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a5304-136">Тип:</span><span class="sxs-lookup"><span data-stu-id="a5304-136">Type:</span></span>

*   <span data-ttu-id="a5304-137">String</span><span class="sxs-lookup"><span data-stu-id="a5304-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a5304-138">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a5304-138">Properties:</span></span>

|<span data-ttu-id="a5304-139">Имя</span><span class="sxs-lookup"><span data-stu-id="a5304-139">Name</span></span>| <span data-ttu-id="a5304-140">Тип</span><span class="sxs-lookup"><span data-stu-id="a5304-140">Type</span></span>| <span data-ttu-id="a5304-141">Описание</span><span class="sxs-lookup"><span data-stu-id="a5304-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a5304-142">String</span><span class="sxs-lookup"><span data-stu-id="a5304-142">String</span></span>|<span data-ttu-id="a5304-143">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="a5304-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a5304-144">String</span><span class="sxs-lookup"><span data-stu-id="a5304-144">String</span></span>|<span data-ttu-id="a5304-145">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="a5304-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5304-146">Требования</span><span class="sxs-lookup"><span data-stu-id="a5304-146">Requirements</span></span>

|<span data-ttu-id="a5304-147">Requirement</span><span class="sxs-lookup"><span data-stu-id="a5304-147">Requirement</span></span>| <span data-ttu-id="a5304-148">Значение</span><span class="sxs-lookup"><span data-stu-id="a5304-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5304-149">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a5304-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5304-150">1.0</span><span class="sxs-lookup"><span data-stu-id="a5304-150">1.0</span></span>|
|[<span data-ttu-id="a5304-151">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a5304-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5304-152">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a5304-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="a5304-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="a5304-153">SourceProperty :String</span></span>

<span data-ttu-id="a5304-154">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="a5304-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a5304-155">Тип:</span><span class="sxs-lookup"><span data-stu-id="a5304-155">Type:</span></span>

*   <span data-ttu-id="a5304-156">String</span><span class="sxs-lookup"><span data-stu-id="a5304-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a5304-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a5304-157">Properties:</span></span>

|<span data-ttu-id="a5304-158">Имя</span><span class="sxs-lookup"><span data-stu-id="a5304-158">Name</span></span>| <span data-ttu-id="a5304-159">Тип</span><span class="sxs-lookup"><span data-stu-id="a5304-159">Type</span></span>| <span data-ttu-id="a5304-160">Описание</span><span class="sxs-lookup"><span data-stu-id="a5304-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a5304-161">String</span><span class="sxs-lookup"><span data-stu-id="a5304-161">String</span></span>|<span data-ttu-id="a5304-162">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="a5304-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a5304-163">String</span><span class="sxs-lookup"><span data-stu-id="a5304-163">String</span></span>|<span data-ttu-id="a5304-164">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="a5304-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5304-165">Требования</span><span class="sxs-lookup"><span data-stu-id="a5304-165">Requirements</span></span>

|<span data-ttu-id="a5304-166">Requirement</span><span class="sxs-lookup"><span data-stu-id="a5304-166">Requirement</span></span>| <span data-ttu-id="a5304-167">Значение</span><span class="sxs-lookup"><span data-stu-id="a5304-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5304-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a5304-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5304-169">1.0</span><span class="sxs-lookup"><span data-stu-id="a5304-169">1.0</span></span>|
|[<span data-ttu-id="a5304-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a5304-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5304-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a5304-171">Compose or read</span></span>|