
# <a name="item"></a><span data-ttu-id="5d10e-101">item</span><span class="sxs-lookup"><span data-stu-id="5d10e-101">item</span></span>

### <span data-ttu-id="5d10e-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="5d10e-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="5d10e-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="5d10e-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="5d10e-106">Requirements</span></span>

|<span data-ttu-id="5d10e-107">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-107">Requirement</span></span>| <span data-ttu-id="5d10e-108">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-109">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-110">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-110">1.0</span></span>|
|[<span data-ttu-id="5d10e-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-112">Restricted</span><span class="sxs-lookup"><span data-stu-id="5d10e-112">Restricted</span></span>|
|[<span data-ttu-id="5d10e-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="5d10e-115">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-115">Example</span></span>

<span data-ttu-id="5d10e-116">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="5d10e-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a><span data-ttu-id="5d10e-117">Элементы</span><span class="sxs-lookup"><span data-stu-id="5d10e-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="5d10e-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="5d10e-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="5d10e-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5d10e-121">Определенные типы файлов блокируемых в Outlook из-за потенциальных проблем безопасности и поэтому не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="5d10e-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="5d10e-122">Для получения дополнительных сведений см [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="5d10e-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-123">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-123">Type:</span></span>

*   <span data-ttu-id="5d10e-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="5d10e-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-125">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-125">Requirements</span></span>

|<span data-ttu-id="5d10e-126">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-126">Requirement</span></span>| <span data-ttu-id="5d10e-127">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-128">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-128">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-129">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-129">1.0</span></span>|
|[<span data-ttu-id="5d10e-130">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-131">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-132">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-133">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-134">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-134">Example</span></span>

<span data-ttu-id="5d10e-135">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5d10e-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```JavaScript
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="5d10e-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d10e-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="5d10e-137">Получает объект, который предоставляет методы для получения или обновления получателей в строке (Скрытая копия) скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="5d10e-138">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5d10e-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-139">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-139">Type:</span></span>

*   [<span data-ttu-id="5d10e-140">Recipients</span><span class="sxs-lookup"><span data-stu-id="5d10e-140">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="5d10e-141">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-141">Requirements</span></span>

|<span data-ttu-id="5d10e-142">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-142">Requirement</span></span>| <span data-ttu-id="5d10e-143">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-145">1.1</span><span class="sxs-lookup"><span data-stu-id="5d10e-145">1.1</span></span>|
|[<span data-ttu-id="5d10e-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-147">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-149">Создание</span><span class="sxs-lookup"><span data-stu-id="5d10e-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-150">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="5d10e-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="5d10e-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="5d10e-152">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="5d10e-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-153">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-153">Type:</span></span>

*   [<span data-ttu-id="5d10e-154">Body</span><span class="sxs-lookup"><span data-stu-id="5d10e-154">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="5d10e-155">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-155">Requirements</span></span>

|<span data-ttu-id="5d10e-156">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-156">Requirement</span></span>| <span data-ttu-id="5d10e-157">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-159">1.1</span><span class="sxs-lookup"><span data-stu-id="5d10e-159">1.1</span></span>|
|[<span data-ttu-id="5d10e-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-161">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="5d10e-164">cc: массив. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[получателей](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d10e-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="5d10e-165">Предоставляет доступ к «копия» (копия) получателей сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="5d10e-166">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5d10e-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5d10e-167">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5d10e-167">Read mode</span></span>

<span data-ttu-id="5d10e-p107">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5d10e-170">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5d10e-170">Compose mode</span></span>

<span data-ttu-id="5d10e-171">`cc` Возвращает свойство `Recipients` объект, предоставляющий методы для получения или обновления получателей в строке **копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-171">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-172">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-172">Type:</span></span>

*   <span data-ttu-id="5d10e-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d10e-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-174">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-174">Requirements</span></span>

|<span data-ttu-id="5d10e-175">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-175">Requirement</span></span>| <span data-ttu-id="5d10e-176">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-177">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-178">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-178">1.0</span></span>|
|[<span data-ttu-id="5d10e-179">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-180">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-181">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-182">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-183">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="5d10e-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="5d10e-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="5d10e-185">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="5d10e-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="5d10e-p108">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="5d10e-p109">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-190">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-190">Type:</span></span>

*   <span data-ttu-id="5d10e-191">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-192">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-192">Requirements</span></span>

|<span data-ttu-id="5d10e-193">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-193">Requirement</span></span>| <span data-ttu-id="5d10e-194">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-195">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-195">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-196">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-196">1.0</span></span>|
|[<span data-ttu-id="5d10e-197">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-198">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="5d10e-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="5d10e-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="5d10e-p110">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-204">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-204">Type:</span></span>

*   <span data-ttu-id="5d10e-205">Date</span><span class="sxs-lookup"><span data-stu-id="5d10e-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-206">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-206">Requirements</span></span>

|<span data-ttu-id="5d10e-207">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-207">Requirement</span></span>| <span data-ttu-id="5d10e-208">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-209">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-209">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-210">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-210">1.0</span></span>|
|[<span data-ttu-id="5d10e-211">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-212">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-214">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-215">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="5d10e-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="5d10e-216">dateTimeModified :Date</span></span>

<span data-ttu-id="5d10e-p111">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5d10e-219">Этот член не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5d10e-219">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-220">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-220">Type:</span></span>

*   <span data-ttu-id="5d10e-221">Date</span><span class="sxs-lookup"><span data-stu-id="5d10e-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-222">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-222">Requirements</span></span>

|<span data-ttu-id="5d10e-223">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-223">Requirement</span></span>| <span data-ttu-id="5d10e-224">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-225">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-225">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-226">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-226">1.0</span></span>|
|[<span data-ttu-id="5d10e-227">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-228">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-229">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-230">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-231">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="5d10e-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="5d10e-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="5d10e-233">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="5d10e-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="5d10e-p112">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="5d10e-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5d10e-236">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5d10e-236">Read mode</span></span>

<span data-ttu-id="5d10e-237">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5d10e-238">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5d10e-238">Compose mode</span></span>

<span data-ttu-id="5d10e-239">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="5d10e-240">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="5d10e-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-241">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-241">Type:</span></span>

*   <span data-ttu-id="5d10e-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="5d10e-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-243">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-243">Requirements</span></span>

|<span data-ttu-id="5d10e-244">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-244">Requirement</span></span>| <span data-ttu-id="5d10e-245">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-246">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-246">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-247">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-247">1.0</span></span>|
|[<span data-ttu-id="5d10e-248">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-249">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-250">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-251">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-252">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-252">Example</span></span>

<span data-ttu-id="5d10e-253">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5d10e-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="5d10e-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5d10e-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="5d10e-p113">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="5d10e-p114">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="5d10e-259">`recipientType` Свойства `EmailAddressDetails` объект в `from` — это свойство `undefined`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-259">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-260">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-260">Type:</span></span>

*   [<span data-ttu-id="5d10e-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5d10e-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5d10e-262">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-262">Requirements</span></span>

|<span data-ttu-id="5d10e-263">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-263">Requirement</span></span>| <span data-ttu-id="5d10e-264">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-265">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-266">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-266">1.0</span></span>|
|[<span data-ttu-id="5d10e-267">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-268">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-269">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-270">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="5d10e-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="5d10e-271">internetMessageId :String</span></span>

<span data-ttu-id="5d10e-p115">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-274">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-274">Type:</span></span>

*   <span data-ttu-id="5d10e-275">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-276">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-276">Requirements</span></span>

|<span data-ttu-id="5d10e-277">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-277">Requirement</span></span>| <span data-ttu-id="5d10e-278">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-279">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-279">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-280">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-280">1.0</span></span>|
|[<span data-ttu-id="5d10e-281">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-282">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-283">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-284">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-285">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="5d10e-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="5d10e-286">itemClass :String</span></span>

<span data-ttu-id="5d10e-p116">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="5d10e-p117">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="5d10e-291">Тип</span><span class="sxs-lookup"><span data-stu-id="5d10e-291">Type</span></span> | <span data-ttu-id="5d10e-292">Описание</span><span class="sxs-lookup"><span data-stu-id="5d10e-292">Description</span></span> | <span data-ttu-id="5d10e-293">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="5d10e-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="5d10e-294">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="5d10e-294">Appointment items</span></span> | <span data-ttu-id="5d10e-295">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="5d10e-296">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="5d10e-296">Message items</span></span> | <span data-ttu-id="5d10e-297">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="5d10e-298">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-299">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-299">Type:</span></span>

*   <span data-ttu-id="5d10e-300">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-301">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-301">Requirements</span></span>

|<span data-ttu-id="5d10e-302">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-302">Requirement</span></span>| <span data-ttu-id="5d10e-303">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-304">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-304">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-305">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-305">1.0</span></span>|
|[<span data-ttu-id="5d10e-306">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-307">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-308">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-309">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-310">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="5d10e-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="5d10e-311">(nullable) itemId :String</span></span>

<span data-ttu-id="5d10e-p118">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5d10e-314">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="5d10e-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="5d10e-315">`itemId` Свойство не совпадать с Идентификатором, используемым API-Интерфейс REST Outlook или идентификатор записи Outlook.</span><span class="sxs-lookup"><span data-stu-id="5d10e-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="5d10e-316">Прежде чем API-Интерфейс REST вызовы с использованием это значение, необходимо преобразовать с помощью `Office.context.mailbox.convertToRestId`, который появился в требование задано 1.3.</span><span class="sxs-lookup"><span data-stu-id="5d10e-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="5d10e-317">Для получения дополнительных сведений показано [Использование API REST Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="5d10e-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-318">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-318">Type:</span></span>

*   <span data-ttu-id="5d10e-319">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-320">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-320">Requirements</span></span>

|<span data-ttu-id="5d10e-321">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-321">Requirement</span></span>| <span data-ttu-id="5d10e-322">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-323">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-323">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-324">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-324">1.0</span></span>|
|[<span data-ttu-id="5d10e-325">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-326">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-327">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-328">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-329">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-329">Example</span></span>

<span data-ttu-id="5d10e-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="5d10e-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="5d10e-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="5d10e-333">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="5d10e-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="5d10e-334">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="5d10e-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-335">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-335">Type:</span></span>

*   [<span data-ttu-id="5d10e-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="5d10e-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="5d10e-337">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-337">Requirements</span></span>

|<span data-ttu-id="5d10e-338">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-338">Requirement</span></span>| <span data-ttu-id="5d10e-339">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-340">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-341">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-341">1.0</span></span>|
|[<span data-ttu-id="5d10e-342">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-343">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-344">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-345">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-346">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="5d10e-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="5d10e-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="5d10e-348">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="5d10e-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5d10e-349">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5d10e-349">Read mode</span></span>

<span data-ttu-id="5d10e-350">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="5d10e-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5d10e-351">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5d10e-351">Compose mode</span></span>

<span data-ttu-id="5d10e-352">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="5d10e-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-353">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-353">Type:</span></span>

*   <span data-ttu-id="5d10e-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="5d10e-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-355">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-355">Requirements</span></span>

|<span data-ttu-id="5d10e-356">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-356">Requirement</span></span>| <span data-ttu-id="5d10e-357">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-358">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-358">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-359">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-359">1.0</span></span>|
|[<span data-ttu-id="5d10e-360">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-361">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-362">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-363">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-364">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="5d10e-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="5d10e-365">normalizedSubject :String</span></span>

<span data-ttu-id="5d10e-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="5d10e-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject).</span><span class="sxs-lookup"><span data-stu-id="5d10e-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-370">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-370">Type:</span></span>

*   <span data-ttu-id="5d10e-371">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-372">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-372">Requirements</span></span>

|<span data-ttu-id="5d10e-373">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-373">Requirement</span></span>| <span data-ttu-id="5d10e-374">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-375">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-375">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-376">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-376">1.0</span></span>|
|[<span data-ttu-id="5d10e-377">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-378">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-379">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-380">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-381">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="5d10e-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d10e-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="5d10e-383">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="5d10e-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="5d10e-384">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5d10e-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5d10e-385">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5d10e-385">Read mode</span></span>

<span data-ttu-id="5d10e-386">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="5d10e-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5d10e-387">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5d10e-387">Compose mode</span></span>

<span data-ttu-id="5d10e-388">`optionalAttendees` Возвращает свойство `Recipients` объект, предоставляющий методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="5d10e-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-389">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-389">Type:</span></span>

*   <span data-ttu-id="5d10e-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d10e-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-391">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-391">Requirements</span></span>

|<span data-ttu-id="5d10e-392">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-392">Requirement</span></span>| <span data-ttu-id="5d10e-393">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-394">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-394">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-395">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-395">1.0</span></span>|
|[<span data-ttu-id="5d10e-396">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-397">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-398">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-399">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-400">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="5d10e-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5d10e-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="5d10e-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-404">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-404">Type:</span></span>

*   [<span data-ttu-id="5d10e-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5d10e-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5d10e-406">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-406">Requirements</span></span>

|<span data-ttu-id="5d10e-407">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-407">Requirement</span></span>| <span data-ttu-id="5d10e-408">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-409">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-410">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-410">1.0</span></span>|
|[<span data-ttu-id="5d10e-411">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-412">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-413">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-414">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-415">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="5d10e-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d10e-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="5d10e-417">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="5d10e-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="5d10e-418">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5d10e-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5d10e-419">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5d10e-419">Read mode</span></span>

<span data-ttu-id="5d10e-420">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="5d10e-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5d10e-421">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5d10e-421">Compose mode</span></span>

<span data-ttu-id="5d10e-422">`requiredAttendees` Возвращает свойство `Recipients` объект, предоставляющий методы для получения или обновления обязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="5d10e-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-423">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-423">Type:</span></span>

*   <span data-ttu-id="5d10e-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d10e-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-425">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-425">Requirements</span></span>

|<span data-ttu-id="5d10e-426">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-426">Requirement</span></span>| <span data-ttu-id="5d10e-427">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-428">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-428">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-429">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-429">1.0</span></span>|
|[<span data-ttu-id="5d10e-430">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-431">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-432">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-433">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-434">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="5d10e-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5d10e-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="5d10e-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="5d10e-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="5d10e-440">`recipientType` Свойства `EmailAddressDetails` объект в `sender` — это свойство `undefined`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-440">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-441">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-441">Type:</span></span>

*   [<span data-ttu-id="5d10e-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5d10e-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5d10e-443">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-443">Requirements</span></span>

|<span data-ttu-id="5d10e-444">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-444">Requirement</span></span>| <span data-ttu-id="5d10e-445">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-446">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-446">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-447">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-447">1.0</span></span>|
|[<span data-ttu-id="5d10e-448">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-449">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-450">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-451">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-452">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="5d10e-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="5d10e-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="5d10e-454">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="5d10e-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="5d10e-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="5d10e-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5d10e-457">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5d10e-457">Read mode</span></span>

<span data-ttu-id="5d10e-458">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5d10e-459">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5d10e-459">Compose mode</span></span>

<span data-ttu-id="5d10e-460">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="5d10e-461">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="5d10e-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-462">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-462">Type:</span></span>

*   <span data-ttu-id="5d10e-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="5d10e-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-464">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-464">Requirements</span></span>

|<span data-ttu-id="5d10e-465">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-465">Requirement</span></span>| <span data-ttu-id="5d10e-466">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-467">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-467">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-468">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-468">1.0</span></span>|
|[<span data-ttu-id="5d10e-469">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-470">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-471">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-472">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-473">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-473">Example</span></span>

<span data-ttu-id="5d10e-474">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5d10e-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="5d10e-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="5d10e-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="5d10e-476">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="5d10e-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="5d10e-477">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="5d10e-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5d10e-478">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5d10e-478">Read mode</span></span>

<span data-ttu-id="5d10e-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="5d10e-481">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5d10e-481">Compose mode</span></span>

<span data-ttu-id="5d10e-482">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="5d10e-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="5d10e-483">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-483">Type:</span></span>

*   <span data-ttu-id="5d10e-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="5d10e-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-485">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-485">Requirements</span></span>

|<span data-ttu-id="5d10e-486">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-486">Requirement</span></span>| <span data-ttu-id="5d10e-487">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-488">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-488">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-489">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-489">1.0</span></span>|
|[<span data-ttu-id="5d10e-490">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-491">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-492">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-493">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="5d10e-494">Чтобы: массив. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[получателей](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d10e-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="5d10e-495">Предоставляет доступ к получателей в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="5d10e-496">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5d10e-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5d10e-497">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5d10e-497">Read mode</span></span>

<span data-ttu-id="5d10e-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5d10e-500">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5d10e-500">Compose mode</span></span>

<span data-ttu-id="5d10e-501">`to` Возвращает свойство `Recipients` объект, предоставляющий методы для получения или обновления получателей в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="5d10e-502">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d10e-502">Type:</span></span>

*   <span data-ttu-id="5d10e-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d10e-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-504">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-504">Requirements</span></span>

|<span data-ttu-id="5d10e-505">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-505">Requirement</span></span>| <span data-ttu-id="5d10e-506">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-507">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-508">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-508">1.0</span></span>|
|[<span data-ttu-id="5d10e-509">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-510">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-511">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-512">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-513">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="5d10e-514">Методы</span><span class="sxs-lookup"><span data-stu-id="5d10e-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="5d10e-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5d10e-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="5d10e-516">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="5d10e-517">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="5d10e-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="5d10e-518">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="5d10e-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d10e-519">Параметры</span><span class="sxs-lookup"><span data-stu-id="5d10e-519">Parameters:</span></span>

|<span data-ttu-id="5d10e-520">Имя</span><span class="sxs-lookup"><span data-stu-id="5d10e-520">Name</span></span>| <span data-ttu-id="5d10e-521">Тип</span><span class="sxs-lookup"><span data-stu-id="5d10e-521">Type</span></span>| <span data-ttu-id="5d10e-522">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5d10e-522">Attributes</span></span>| <span data-ttu-id="5d10e-523">Описание</span><span class="sxs-lookup"><span data-stu-id="5d10e-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="5d10e-524">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-524">String</span></span>||<span data-ttu-id="5d10e-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="5d10e-527">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-527">String</span></span>||<span data-ttu-id="5d10e-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="5d10e-530">Object</span><span class="sxs-lookup"><span data-stu-id="5d10e-530">Object</span></span>| <span data-ttu-id="5d10e-531">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-531">&lt;optional&gt;</span></span>|<span data-ttu-id="5d10e-532">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5d10e-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5d10e-533">Object</span><span class="sxs-lookup"><span data-stu-id="5d10e-533">Object</span></span>| <span data-ttu-id="5d10e-534">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-534">&lt;optional&gt;</span></span>|<span data-ttu-id="5d10e-535">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5d10e-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5d10e-536">function</span><span class="sxs-lookup"><span data-stu-id="5d10e-536">function</span></span>| <span data-ttu-id="5d10e-537">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-537">&lt;optional&gt;</span></span>|<span data-ttu-id="5d10e-538">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d10e-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5d10e-539">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="5d10e-540">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="5d10e-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5d10e-541">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5d10e-541">Errors</span></span>

| <span data-ttu-id="5d10e-542">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5d10e-542">Error code</span></span> | <span data-ttu-id="5d10e-543">Описание</span><span class="sxs-lookup"><span data-stu-id="5d10e-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="5d10e-544">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="5d10e-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="5d10e-545">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="5d10e-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="5d10e-546">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="5d10e-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5d10e-547">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-547">Requirements</span></span>

|<span data-ttu-id="5d10e-548">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-548">Requirement</span></span>| <span data-ttu-id="5d10e-549">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-550">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-550">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-551">1.1</span><span class="sxs-lookup"><span data-stu-id="5d10e-551">1.1</span></span>|
|[<span data-ttu-id="5d10e-552">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="5d10e-554">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-555">Создание</span><span class="sxs-lookup"><span data-stu-id="5d10e-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-556">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-556">Example</span></span>

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="5d10e-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5d10e-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="5d10e-558">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="5d10e-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="5d10e-562">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="5d10e-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="5d10e-563">Если надстройки Office работает в Outlook Web App, `addItemAttachmentAsync` метод могут прикреплять элементов для элементов, отличных от элемента, который вы изменяете; Однако это не поддерживается и не рекомендуется.</span><span class="sxs-lookup"><span data-stu-id="5d10e-563">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d10e-564">Параметры</span><span class="sxs-lookup"><span data-stu-id="5d10e-564">Parameters:</span></span>

|<span data-ttu-id="5d10e-565">Имя</span><span class="sxs-lookup"><span data-stu-id="5d10e-565">Name</span></span>| <span data-ttu-id="5d10e-566">Тип</span><span class="sxs-lookup"><span data-stu-id="5d10e-566">Type</span></span>| <span data-ttu-id="5d10e-567">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5d10e-567">Attributes</span></span>| <span data-ttu-id="5d10e-568">Описание</span><span class="sxs-lookup"><span data-stu-id="5d10e-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="5d10e-569">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-569">String</span></span>||<span data-ttu-id="5d10e-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="5d10e-572">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-572">String</span></span>||<span data-ttu-id="5d10e-p136">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="5d10e-575">Object</span><span class="sxs-lookup"><span data-stu-id="5d10e-575">Object</span></span>| <span data-ttu-id="5d10e-576">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-576">&lt;optional&gt;</span></span>|<span data-ttu-id="5d10e-577">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5d10e-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5d10e-578">Object</span><span class="sxs-lookup"><span data-stu-id="5d10e-578">Object</span></span>| <span data-ttu-id="5d10e-579">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-579">&lt;optional&gt;</span></span>|<span data-ttu-id="5d10e-580">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5d10e-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5d10e-581">function</span><span class="sxs-lookup"><span data-stu-id="5d10e-581">function</span></span>| <span data-ttu-id="5d10e-582">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-582">&lt;optional&gt;</span></span>|<span data-ttu-id="5d10e-583">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d10e-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5d10e-584">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="5d10e-585">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="5d10e-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5d10e-586">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5d10e-586">Errors</span></span>

| <span data-ttu-id="5d10e-587">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5d10e-587">Error code</span></span> | <span data-ttu-id="5d10e-588">Описание</span><span class="sxs-lookup"><span data-stu-id="5d10e-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="5d10e-589">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="5d10e-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5d10e-590">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-590">Requirements</span></span>

|<span data-ttu-id="5d10e-591">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-591">Requirement</span></span>| <span data-ttu-id="5d10e-592">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-593">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-593">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-594">1.1</span><span class="sxs-lookup"><span data-stu-id="5d10e-594">1.1</span></span>|
|[<span data-ttu-id="5d10e-595">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="5d10e-597">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-598">Создание</span><span class="sxs-lookup"><span data-stu-id="5d10e-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-599">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-599">Example</span></span>

<span data-ttu-id="5d10e-600">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="5d10e-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="5d10e-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="5d10e-602">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="5d10e-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5d10e-603">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5d10e-603">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5d10e-604">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="5d10e-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="5d10e-605">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="5d10e-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="5d10e-p137">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d10e-609">Параметры</span><span class="sxs-lookup"><span data-stu-id="5d10e-609">Parameters:</span></span>

|<span data-ttu-id="5d10e-610">Имя</span><span class="sxs-lookup"><span data-stu-id="5d10e-610">Name</span></span>| <span data-ttu-id="5d10e-611">Тип</span><span class="sxs-lookup"><span data-stu-id="5d10e-611">Type</span></span>| <span data-ttu-id="5d10e-612">Описание</span><span class="sxs-lookup"><span data-stu-id="5d10e-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="5d10e-613">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="5d10e-613">String &#124; Object</span></span>| |<span data-ttu-id="5d10e-p138">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="5d10e-616">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="5d10e-616">**OR**</span></span><br/><span data-ttu-id="5d10e-p139">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="5d10e-619">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-619">String</span></span> | <span data-ttu-id="5d10e-620">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-620">&lt;optional&gt;</span></span> | <span data-ttu-id="5d10e-p140">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="5d10e-623">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-623">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="5d10e-624">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-624">&lt;optional&gt;</span></span> | <span data-ttu-id="5d10e-625">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="5d10e-625">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="5d10e-626">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-626">String</span></span> | | <span data-ttu-id="5d10e-p141">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="5d10e-629">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-629">String</span></span> | | <span data-ttu-id="5d10e-630">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5d10e-630">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="5d10e-631">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-631">String</span></span> | | <span data-ttu-id="5d10e-p142">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="5d10e-634">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-634">String</span></span> | | <span data-ttu-id="5d10e-p143">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="5d10e-638">function</span><span class="sxs-lookup"><span data-stu-id="5d10e-638">function</span></span> | <span data-ttu-id="5d10e-639">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-639">&lt;optional&gt;</span></span> | <span data-ttu-id="5d10e-640">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d10e-640">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5d10e-641">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-641">Requirements</span></span>

|<span data-ttu-id="5d10e-642">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-642">Requirement</span></span>| <span data-ttu-id="5d10e-643">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-644">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-644">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-645">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-645">1.0</span></span>|
|[<span data-ttu-id="5d10e-646">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-647">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-648">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-649">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-649">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="5d10e-650">Примеры</span><span class="sxs-lookup"><span data-stu-id="5d10e-650">Examples</span></span>

<span data-ttu-id="5d10e-651">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-651">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="5d10e-652">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-652">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="5d10e-653">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-653">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="5d10e-654">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="5d10e-654">Reply with a body and a file attachment.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="5d10e-655">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="5d10e-655">Reply with a body and an item attachment.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="5d10e-656">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="5d10e-656">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="5d10e-657">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="5d10e-657">displayReplyForm(formData)</span></span>

<span data-ttu-id="5d10e-658">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="5d10e-658">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5d10e-659">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5d10e-659">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5d10e-660">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="5d10e-660">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="5d10e-661">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="5d10e-661">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="5d10e-p144">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d10e-665">Параметры</span><span class="sxs-lookup"><span data-stu-id="5d10e-665">Parameters:</span></span>

|<span data-ttu-id="5d10e-666">Имя</span><span class="sxs-lookup"><span data-stu-id="5d10e-666">Name</span></span>| <span data-ttu-id="5d10e-667">Тип</span><span class="sxs-lookup"><span data-stu-id="5d10e-667">Type</span></span>| <span data-ttu-id="5d10e-668">Описание</span><span class="sxs-lookup"><span data-stu-id="5d10e-668">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="5d10e-669">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="5d10e-669">String &#124; Object</span></span>| | <span data-ttu-id="5d10e-p145">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="5d10e-672">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="5d10e-672">**OR**</span></span><br/><span data-ttu-id="5d10e-p146">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="5d10e-675">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-675">String</span></span> | <span data-ttu-id="5d10e-676">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-676">&lt;optional&gt;</span></span> | <span data-ttu-id="5d10e-p147">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="5d10e-679">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-679">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="5d10e-680">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-680">&lt;optional&gt;</span></span> | <span data-ttu-id="5d10e-681">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="5d10e-681">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="5d10e-682">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-682">String</span></span> | | <span data-ttu-id="5d10e-p148">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="5d10e-685">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-685">String</span></span> | | <span data-ttu-id="5d10e-686">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5d10e-686">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="5d10e-687">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-687">String</span></span> | | <span data-ttu-id="5d10e-p149">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="5d10e-690">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-690">String</span></span> | | <span data-ttu-id="5d10e-p150">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="5d10e-694">function</span><span class="sxs-lookup"><span data-stu-id="5d10e-694">function</span></span> | <span data-ttu-id="5d10e-695">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-695">&lt;optional&gt;</span></span> | <span data-ttu-id="5d10e-696">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d10e-696">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5d10e-697">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-697">Requirements</span></span>

|<span data-ttu-id="5d10e-698">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-698">Requirement</span></span>| <span data-ttu-id="5d10e-699">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-699">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-700">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-700">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-701">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-701">1.0</span></span>|
|[<span data-ttu-id="5d10e-702">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-702">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-703">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-703">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-704">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-704">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-705">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-705">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="5d10e-706">Примеры</span><span class="sxs-lookup"><span data-stu-id="5d10e-706">Examples</span></span>

<span data-ttu-id="5d10e-707">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-707">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="5d10e-708">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-708">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="5d10e-709">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-709">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="5d10e-710">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="5d10e-710">Reply with a body and a file attachment.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="5d10e-711">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="5d10e-711">Reply with a body and an item attachment.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="5d10e-712">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="5d10e-712">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="5d10e-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="5d10e-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="5d10e-714">Возвращает сущности, обнаруженные в тело выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5d10e-714">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="5d10e-715">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5d10e-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-716">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-716">Requirements</span></span>

|<span data-ttu-id="5d10e-717">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-717">Requirement</span></span>| <span data-ttu-id="5d10e-718">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-719">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-719">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-720">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-720">1.0</span></span>|
|[<span data-ttu-id="5d10e-721">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-721">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-722">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-723">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-723">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-724">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-724">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5d10e-725">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5d10e-725">Returns:</span></span>

<span data-ttu-id="5d10e-726">Тип: [Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="5d10e-726">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="5d10e-727">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-727">Example</span></span>

<span data-ttu-id="5d10e-728">Этот пример ссылается сущностей контакты в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5d10e-728">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="5d10e-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="5d10e-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="5d10e-730">Получает массив всех сущностей указанного типа, обнаруженных в тело выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5d10e-730">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="5d10e-731">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5d10e-731">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d10e-732">Параметры</span><span class="sxs-lookup"><span data-stu-id="5d10e-732">Parameters:</span></span>

|<span data-ttu-id="5d10e-733">Имя</span><span class="sxs-lookup"><span data-stu-id="5d10e-733">Name</span></span>| <span data-ttu-id="5d10e-734">Тип</span><span class="sxs-lookup"><span data-stu-id="5d10e-734">Type</span></span>| <span data-ttu-id="5d10e-735">Описание</span><span class="sxs-lookup"><span data-stu-id="5d10e-735">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="5d10e-736">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="5d10e-736">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="5d10e-737">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="5d10e-737">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5d10e-738">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-738">Requirements</span></span>

|<span data-ttu-id="5d10e-739">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-739">Requirement</span></span>| <span data-ttu-id="5d10e-740">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-741">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-741">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-742">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-742">1.0</span></span>|
|[<span data-ttu-id="5d10e-743">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-743">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-744">Restricted</span><span class="sxs-lookup"><span data-stu-id="5d10e-744">Restricted</span></span>|
|[<span data-ttu-id="5d10e-745">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-745">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-746">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-746">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5d10e-747">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5d10e-747">Returns:</span></span>

<span data-ttu-id="5d10e-748">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="5d10e-748">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="5d10e-749">Если сущности указанного типа отсутствуют в основной текст элемента, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="5d10e-749">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="5d10e-750">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-750">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="5d10e-751">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="5d10e-751">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="5d10e-752">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="5d10e-752">Value of `entityType`</span></span> | <span data-ttu-id="5d10e-753">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="5d10e-753">Type of objects in returned array</span></span> | <span data-ttu-id="5d10e-754">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-754">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="5d10e-755">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-755">String</span></span> | <span data-ttu-id="5d10e-756">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="5d10e-756">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="5d10e-757">Contact</span><span class="sxs-lookup"><span data-stu-id="5d10e-757">Contact</span></span> | <span data-ttu-id="5d10e-758">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5d10e-758">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="5d10e-759">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-759">String</span></span> | <span data-ttu-id="5d10e-760">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5d10e-760">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="5d10e-761">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="5d10e-761">MeetingSuggestion</span></span> | <span data-ttu-id="5d10e-762">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5d10e-762">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="5d10e-763">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="5d10e-763">PhoneNumber</span></span> | <span data-ttu-id="5d10e-764">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="5d10e-764">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="5d10e-765">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="5d10e-765">TaskSuggestion</span></span> | <span data-ttu-id="5d10e-766">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5d10e-766">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="5d10e-767">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-767">String</span></span> | <span data-ttu-id="5d10e-768">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="5d10e-768">**Restricted**</span></span> |

<span data-ttu-id="5d10e-769">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="5d10e-769">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="5d10e-770">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-770">Example</span></span>

<span data-ttu-id="5d10e-771">Следующем примере показано, как получить доступ к массив строк, представляющих почтовых адресов в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5d10e-771">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="5d10e-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="5d10e-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="5d10e-773">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5d10e-773">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5d10e-774">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5d10e-774">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5d10e-775">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-775">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d10e-776">Параметры</span><span class="sxs-lookup"><span data-stu-id="5d10e-776">Parameters:</span></span>

|<span data-ttu-id="5d10e-777">Имя</span><span class="sxs-lookup"><span data-stu-id="5d10e-777">Name</span></span>| <span data-ttu-id="5d10e-778">Тип</span><span class="sxs-lookup"><span data-stu-id="5d10e-778">Type</span></span>| <span data-ttu-id="5d10e-779">Описание</span><span class="sxs-lookup"><span data-stu-id="5d10e-779">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="5d10e-780">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-780">String</span></span>|<span data-ttu-id="5d10e-781">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="5d10e-781">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5d10e-782">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-782">Requirements</span></span>

|<span data-ttu-id="5d10e-783">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-783">Requirement</span></span>| <span data-ttu-id="5d10e-784">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-784">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-785">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-785">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-786">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-786">1.0</span></span>|
|[<span data-ttu-id="5d10e-787">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-787">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-788">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-788">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-789">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-789">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-790">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-790">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5d10e-791">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5d10e-791">Returns:</span></span>

<span data-ttu-id="5d10e-p152">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="5d10e-794">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="5d10e-794">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="5d10e-795">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="5d10e-795">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="5d10e-796">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5d10e-796">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5d10e-797">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5d10e-797">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5d10e-p153">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="5d10e-801">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-801">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="5d10e-802">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-802">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="5d10e-p154">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d10e-805">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-805">Requirements</span></span>

|<span data-ttu-id="5d10e-806">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-806">Requirement</span></span>| <span data-ttu-id="5d10e-807">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-808">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-808">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-809">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-809">1.0</span></span>|
|[<span data-ttu-id="5d10e-810">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-810">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-811">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-812">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-812">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-813">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-813">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5d10e-814">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5d10e-814">Returns:</span></span>

<span data-ttu-id="5d10e-p155">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="5d10e-817">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="5d10e-817">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5d10e-818">Object</span><span class="sxs-lookup"><span data-stu-id="5d10e-818">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5d10e-819">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-819">Example</span></span>

<span data-ttu-id="5d10e-820">В примере ниже показано, как получить доступ к массиву совпадений для элементов <rule> регулярного выражения `fruits` и `veggies`, которые указаны в манифесте.</rule></span><span class="sxs-lookup"><span data-stu-id="5d10e-820">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="5d10e-821">getRegExMatchesByName(name) пункты (допускает значение NULL) {массива. < String >}</span><span class="sxs-lookup"><span data-stu-id="5d10e-821">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="5d10e-822">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5d10e-822">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5d10e-823">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5d10e-823">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5d10e-824">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-824">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="5d10e-p156">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d10e-827">Параметры</span><span class="sxs-lookup"><span data-stu-id="5d10e-827">Parameters:</span></span>

|<span data-ttu-id="5d10e-828">Имя</span><span class="sxs-lookup"><span data-stu-id="5d10e-828">Name</span></span>| <span data-ttu-id="5d10e-829">Тип</span><span class="sxs-lookup"><span data-stu-id="5d10e-829">Type</span></span>| <span data-ttu-id="5d10e-830">Описание</span><span class="sxs-lookup"><span data-stu-id="5d10e-830">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="5d10e-831">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-831">String</span></span>|<span data-ttu-id="5d10e-832">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="5d10e-832">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5d10e-833">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-833">Requirements</span></span>

|<span data-ttu-id="5d10e-834">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-834">Requirement</span></span>| <span data-ttu-id="5d10e-835">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-836">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-836">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-837">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-837">1.0</span></span>|
|[<span data-ttu-id="5d10e-838">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-839">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-839">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-840">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-841">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-841">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5d10e-842">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5d10e-842">Returns:</span></span>

<span data-ttu-id="5d10e-843">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5d10e-843">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="5d10e-844">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="5d10e-844">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5d10e-845">Массив. < String ></span><span class="sxs-lookup"><span data-stu-id="5d10e-845">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5d10e-846">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-846">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="5d10e-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="5d10e-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="5d10e-848">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-848">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="5d10e-p157">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d10e-851">Параметры</span><span class="sxs-lookup"><span data-stu-id="5d10e-851">Parameters:</span></span>

|<span data-ttu-id="5d10e-852">Имя</span><span class="sxs-lookup"><span data-stu-id="5d10e-852">Name</span></span>| <span data-ttu-id="5d10e-853">Тип</span><span class="sxs-lookup"><span data-stu-id="5d10e-853">Type</span></span>| <span data-ttu-id="5d10e-854">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5d10e-854">Attributes</span></span>| <span data-ttu-id="5d10e-855">Описание</span><span class="sxs-lookup"><span data-stu-id="5d10e-855">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="5d10e-856">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="5d10e-856">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="5d10e-p158">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="5d10e-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="5d10e-860">Object</span><span class="sxs-lookup"><span data-stu-id="5d10e-860">Object</span></span>| <span data-ttu-id="5d10e-861">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-861">&lt;optional&gt;</span></span>|<span data-ttu-id="5d10e-862">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5d10e-862">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5d10e-863">Object</span><span class="sxs-lookup"><span data-stu-id="5d10e-863">Object</span></span>| <span data-ttu-id="5d10e-864">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-864">&lt;optional&gt;</span></span>|<span data-ttu-id="5d10e-865">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5d10e-865">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5d10e-866">function</span><span class="sxs-lookup"><span data-stu-id="5d10e-866">function</span></span>||<span data-ttu-id="5d10e-867">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d10e-867">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5d10e-868">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-868">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="5d10e-869">Для доступа к свойству источника, выделение, поступающих из источников, вызовите `asyncResult.value.sourceProperty`, который может быть либо `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-869">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5d10e-870">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-870">Requirements</span></span>

|<span data-ttu-id="5d10e-871">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-871">Requirement</span></span>| <span data-ttu-id="5d10e-872">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-873">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5d10e-873">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-874">1.2</span><span class="sxs-lookup"><span data-stu-id="5d10e-874">1.2</span></span>|
|[<span data-ttu-id="5d10e-875">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="5d10e-877">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-878">Создание</span><span class="sxs-lookup"><span data-stu-id="5d10e-878">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="5d10e-879">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5d10e-879">Returns:</span></span>

<span data-ttu-id="5d10e-880">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-880">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="5d10e-881">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="5d10e-881">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5d10e-882">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-882">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5d10e-883">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-883">Example</span></span>

```JavaScript
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="5d10e-884">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="5d10e-884">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="5d10e-885">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5d10e-885">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="5d10e-p160">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d10e-889">Параметры</span><span class="sxs-lookup"><span data-stu-id="5d10e-889">Parameters:</span></span>

|<span data-ttu-id="5d10e-890">Имя</span><span class="sxs-lookup"><span data-stu-id="5d10e-890">Name</span></span>| <span data-ttu-id="5d10e-891">Тип</span><span class="sxs-lookup"><span data-stu-id="5d10e-891">Type</span></span>| <span data-ttu-id="5d10e-892">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5d10e-892">Attributes</span></span>| <span data-ttu-id="5d10e-893">Описание</span><span class="sxs-lookup"><span data-stu-id="5d10e-893">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="5d10e-894">function</span><span class="sxs-lookup"><span data-stu-id="5d10e-894">function</span></span>||<span data-ttu-id="5d10e-895">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d10e-895">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5d10e-896">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-896">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="5d10e-897">Этот объект можно использовать для получения, задания и удаление настраиваемых свойств из элемента и сохранение изменений для настраиваемого свойства, задайте обратно на сервер.</span><span class="sxs-lookup"><span data-stu-id="5d10e-897">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="5d10e-898">Объект</span><span class="sxs-lookup"><span data-stu-id="5d10e-898">Object</span></span>| <span data-ttu-id="5d10e-899">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-899">&lt;optional&gt;</span></span>|<span data-ttu-id="5d10e-900">Разработчики могут предоставлять любого объекта, которые следует получить доступ к в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5d10e-900">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="5d10e-901">Этот объект можно получить доступ с `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5d10e-901">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5d10e-902">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-902">Requirements</span></span>

|<span data-ttu-id="5d10e-903">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-903">Requirement</span></span>| <span data-ttu-id="5d10e-904">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-905">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-906">1.0</span><span class="sxs-lookup"><span data-stu-id="5d10e-906">1.0</span></span>|
|[<span data-ttu-id="5d10e-907">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-908">ReadItem</span></span>|
|[<span data-ttu-id="5d10e-909">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-910">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d10e-910">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-911">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-911">Example</span></span>

<span data-ttu-id="5d10e-p163">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```JavaScript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="5d10e-915">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5d10e-915">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="5d10e-916">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="5d10e-916">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="5d10e-p164">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d10e-921">Параметры</span><span class="sxs-lookup"><span data-stu-id="5d10e-921">Parameters:</span></span>

|<span data-ttu-id="5d10e-922">Имя</span><span class="sxs-lookup"><span data-stu-id="5d10e-922">Name</span></span>| <span data-ttu-id="5d10e-923">Тип</span><span class="sxs-lookup"><span data-stu-id="5d10e-923">Type</span></span>| <span data-ttu-id="5d10e-924">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5d10e-924">Attributes</span></span>| <span data-ttu-id="5d10e-925">Описание</span><span class="sxs-lookup"><span data-stu-id="5d10e-925">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="5d10e-926">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-926">String</span></span>||<span data-ttu-id="5d10e-p165">Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p165">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="5d10e-929">Object</span><span class="sxs-lookup"><span data-stu-id="5d10e-929">Object</span></span>| <span data-ttu-id="5d10e-930">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-930">&lt;optional&gt;</span></span>|<span data-ttu-id="5d10e-931">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5d10e-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5d10e-932">Object</span><span class="sxs-lookup"><span data-stu-id="5d10e-932">Object</span></span>| <span data-ttu-id="5d10e-933">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-933">&lt;optional&gt;</span></span>|<span data-ttu-id="5d10e-934">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5d10e-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5d10e-935">function</span><span class="sxs-lookup"><span data-stu-id="5d10e-935">function</span></span>| <span data-ttu-id="5d10e-936">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-936">&lt;optional&gt;</span></span>|<span data-ttu-id="5d10e-937">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d10e-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5d10e-938">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="5d10e-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5d10e-939">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5d10e-939">Errors</span></span>

| <span data-ttu-id="5d10e-940">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5d10e-940">Error code</span></span> | <span data-ttu-id="5d10e-941">Описание</span><span class="sxs-lookup"><span data-stu-id="5d10e-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="5d10e-942">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="5d10e-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5d10e-943">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-943">Requirements</span></span>

|<span data-ttu-id="5d10e-944">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-944">Requirement</span></span>| <span data-ttu-id="5d10e-945">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-946">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d10e-946">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-947">1.1</span><span class="sxs-lookup"><span data-stu-id="5d10e-947">1.1</span></span>|
|[<span data-ttu-id="5d10e-948">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="5d10e-950">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-951">Создание</span><span class="sxs-lookup"><span data-stu-id="5d10e-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-952">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-952">Example</span></span>

<span data-ttu-id="5d10e-953">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="5d10e-953">The following code removes an attachment with an identifier of '0'.</span></span>

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="5d10e-954">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="5d10e-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="5d10e-955">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d10e-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="5d10e-p166">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d10e-959">Параметры</span><span class="sxs-lookup"><span data-stu-id="5d10e-959">Parameters:</span></span>

|<span data-ttu-id="5d10e-960">Имя</span><span class="sxs-lookup"><span data-stu-id="5d10e-960">Name</span></span>| <span data-ttu-id="5d10e-961">Тип</span><span class="sxs-lookup"><span data-stu-id="5d10e-961">Type</span></span>| <span data-ttu-id="5d10e-962">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5d10e-962">Attributes</span></span>| <span data-ttu-id="5d10e-963">Описание</span><span class="sxs-lookup"><span data-stu-id="5d10e-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="5d10e-964">String</span><span class="sxs-lookup"><span data-stu-id="5d10e-964">String</span></span>||<span data-ttu-id="5d10e-p167">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="5d10e-968">Object</span><span class="sxs-lookup"><span data-stu-id="5d10e-968">Object</span></span>| <span data-ttu-id="5d10e-969">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-969">&lt;optional&gt;</span></span>|<span data-ttu-id="5d10e-970">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5d10e-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5d10e-971">Object</span><span class="sxs-lookup"><span data-stu-id="5d10e-971">Object</span></span>| <span data-ttu-id="5d10e-972">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-972">&lt;optional&gt;</span></span>|<span data-ttu-id="5d10e-973">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="5d10e-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="5d10e-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="5d10e-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="5d10e-975">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d10e-975">&lt;optional&gt;</span></span>|<span data-ttu-id="5d10e-p168">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="5d10e-p169">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="5d10e-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="5d10e-980">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="5d10e-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="5d10e-981">функция</span><span class="sxs-lookup"><span data-stu-id="5d10e-981">function</span></span>||<span data-ttu-id="5d10e-982">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d10e-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5d10e-983">Требования</span><span class="sxs-lookup"><span data-stu-id="5d10e-983">Requirements</span></span>

|<span data-ttu-id="5d10e-984">Requirement</span><span class="sxs-lookup"><span data-stu-id="5d10e-984">Requirement</span></span>| <span data-ttu-id="5d10e-985">Значение</span><span class="sxs-lookup"><span data-stu-id="5d10e-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d10e-986">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5d10e-986">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d10e-987">1.2</span><span class="sxs-lookup"><span data-stu-id="5d10e-987">1.2</span></span>|
|[<span data-ttu-id="5d10e-988">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d10e-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d10e-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5d10e-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="5d10e-990">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d10e-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d10e-991">Создание</span><span class="sxs-lookup"><span data-stu-id="5d10e-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5d10e-992">Пример</span><span class="sxs-lookup"><span data-stu-id="5d10e-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```